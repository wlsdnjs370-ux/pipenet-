"""Deterministic supply-to-head calculation tree (coordinates/lengths in mm).

This is a route selection model, not a hydraulic solution of a looped network.
Source geometry and fitting ports must be retained separately by consumers.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from collections.abc import Iterable, Mapping, Sequence
from dataclasses import dataclass
import hashlib
import heapq
import math

Edge = tuple[int, int]


def edge_key(a: int, b: int) -> Edge:
    """Return an undirected source edge identity."""
    return (min(a, b), max(a, b))


@dataclass
class FlowTree:
    """One parent per reachable vertex; loads include all unique head positions."""

    roots: tuple[int, ...]
    parent: dict[int, int]
    distance_mm: dict[int, float]
    hop: dict[int, int]
    head_node: dict[int, int]
    representatives: dict[int, int]
    lengths_mm: dict[Edge, float]
    loads: dict[Edge, int]
    excluded: dict[Edge, str]
    revision: str

    @property
    def edges(self) -> set[Edge]:
        """Head-carrying tree edges only (zero-head twigs removed)."""
        return set(self.loads)

    def path(self, node: int) -> list[int]:
        """Return the recorded root-to-node route, without a new path search."""
        if node not in self.distance_mm:
            return []
        route = [node]
        while route[-1] in self.parent:
            route.append(self.parent[route[-1]])
        return route[::-1]

    def between(self, a: int, b: int) -> list[Edge]:
        """Unique source path between two nodes, or an empty list if disconnected."""
        pa, pb = self.path(a), self.path(b)
        if not pa or not pb or pa[0] != pb[0]:
            return []
        i = 0
        while i < min(len(pa), len(pb)) and pa[i] == pb[i]:
            i += 1
        nodes = pa[i - 1:][::-1] + pb[i:]
        return [edge_key(u, v) for u, v in zip(nodes, nodes[1:])]

    def selected_loads(self, heads: Iterable[int]) -> dict[Edge, int]:
        """Count only selected heads on the already fixed tree, without rerouting."""
        count: Counter[Edge] = Counter()
        for hi in sorted({self.representatives[h] for h in heads if h in self.head_node}):
            path = self.path(self.head_node[hi])
            count.update(edge_key(a, b) for a, b in zip(path, path[1:]))
        return dict(count)

    def report(self) -> dict:
        """JSON-compatible audit summary; exclusions never mean physically absent."""
        return {"revision": self.revision, "roots": list(self.roots),
                "heads": len(set(self.representatives.values())),
                "edges": len(self.loads), "excluded": dict(Counter(self.excluded.values())),
                "max_load_all": max(self.loads.values(), default=0),
                "length_m": round(sum(self.lengths_mm[e] for e in self.loads) / 1000, 3)}


def build_flow_tree(
    points: Sequence[Sequence[float]], edges: Iterable[Sequence[int]],
    head_nodes: Sequence[Iterable[int]], roots: Sequence[int], *,
    head_xy: Sequence[Sequence[float]] | None = None,
    lengths_mm: Mapping[Edge, float] | None = None,
    inferred_edges: Iterable[Edge] = (),
) -> FlowTree:
    """Choose routes by (inferred-link count, real length, hops), then stable IDs.

    Multiple roots are supported only for legacy diagnostics; calculation callers
    must select one valve. No geometry-based crossing/nearby joining occurs here.
    Cycles are first made acyclic; full downstream loads are then counted, before
    zone/K filtering. Zero-length aliases cannot produce parent cycles.
    """
    roots = tuple(sorted(set(roots)))
    if not roots or any(not isinstance(r, int) or not 0 <= r < len(points) for r in roots):
        raise ValueError("알람밸브 기준 노드가 없습니다. 급수원을 먼저 지정하세요.")
    weights = {}
    adj: dict[int, list[int]] = defaultdict(list)
    declared = {edge_key(*e): float(v) for e, v in (lengths_mm or {}).items()}
    inferred = {edge_key(*e) for e in inferred_edges}
    for raw in edges:
        a, b = (int(v) for v in raw)
        if not (0 <= a < len(points) and 0 <= b < len(points)):
            raise ValueError("배관의 원본 노드 참조가 올바르지 않습니다.")
        if a == b:
            continue
        e = edge_key(a, b)
        length = declared.get(e, math.dist(points[a][:2], points[b][:2]))
        if not math.isfinite(length) or length < 0:
            raise ValueError("배관 길이는 0 이상의 유한한 mm 값이어야 합니다.")
        weights[e] = length
    for a, b in sorted(weights):
        adj[a].append(b)
        adj[b].append(a)
    best = {r: (0, 0.0, 0) for r in roots}
    queue = [(0, 0.0, 0, r) for r in roots]
    heapq.heapify(queue)
    parent: dict[int, int] = {}
    while queue:
        cost, dist, hop, u = heapq.heappop(queue)
        if best[u] != (cost, dist, hop):
            continue
        for v in adj[u]:
            e = edge_key(u, v)
            candidate = (cost + int(e in inferred), dist + weights[e], hop + 1)
            if v not in best or candidate < best[v]:
                best[v], parent[v] = candidate, u
                heapq.heappush(queue, (*candidate, v))
    head_node = {}
    reps = {}
    places = {}
    for hi, nodes in enumerate(head_nodes):
        wet = [n for n in nodes if n in best]
        if not wet:
            continue
        node = min(wet, key=lambda n: (*best[n], n))
        place = (tuple(round(float(c), 1) for c in head_xy[hi][:2])
                 if head_xy is not None and hi < len(head_xy) else ("node", node))
        reps[hi] = places.setdefault(place, hi)
        head_node[hi] = node
    counts = Counter(head_node[hi] for hi in set(reps.values()))
    loads = {}
    # Parent always precedes child in the lexicographic Dijkstra cost.
    for u in sorted(best, key=lambda n: (*best[n], n), reverse=True):
        if u in parent:
            p = parent[u]
            counts[p] += counts[u]
            if counts[u]:
                loads[edge_key(u, p)] = counts[u]
    tree_edges = {edge_key(u, p) for u, p in parent.items()}
    excluded = {e: ("unreachable" if e[0] not in best else
                    "no_head" if e in tree_edges else "alternate_route")
                for e in weights if e not in loads}
    signature = (tuple(tuple(p[:2]) for p in points), roots, sorted(parent.items()), sorted(head_node.items()),
                 sorted(reps.items()), sorted(weights.items()), sorted(inferred))
    revision = hashlib.sha256(repr(signature).encode()).hexdigest()[:16]
    return FlowTree(roots, parent, {n: v[1] for n, v in best.items()},
                    {n: v[2] for n, v in best.items()}, head_node, reps, weights,
                    loads, excluded, revision)
