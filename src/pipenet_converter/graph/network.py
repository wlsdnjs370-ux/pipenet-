"""Cycle-preserving source extraction, independent of hydraulic calculations.

The block-cut forest selects the union of all simple supply-to-terminal paths.
It retains transfer branches of grids without enumerating exponentially many
paths. Crossings are never joined here. Loop/grid names are user declarations,
not classifications inferred from cycle counts or the appearance of a drawing.
"""
from __future__ import annotations

from collections import Counter, deque
from collections.abc import Iterable, Sequence
from dataclasses import asdict, dataclass
import hashlib
import math

import networkx as nx

from .flow import Edge, FlowTree, edge_key

NETWORK_MODES = ("tree", "loop", "grid")
CALCULATION_PENDING = (
    "루프·그리드에는 트리 규약 관경을 적용하지 않습니다. "
    "수리계산에서 검토용 기본 관경을 직접 지정하고 표를 확정하면 미확정 SDF를 출력할 수 있습니다. "
    "유량 분배·실제 표고·티 손실·최불리 여부는 별도 검증이 필요합니다."
)


def network_mode(value: object = "tree") -> str:
    """Validate an explicit mode; old drawings without this field remain trees."""
    mode = str(value)
    if mode not in NETWORK_MODES:
        raise ValueError("배관망 방식은 tree, loop, grid 중 하나여야 합니다.")
    return mode


def require_tree_calculation(mode: object = "tree") -> None:
    """Fail closed before a cyclic extraction enters tree-only conversion."""
    if network_mode(mode) != "tree":
        raise ValueError(CALCULATION_PENDING)


@dataclass(frozen=True)
class PathBlocks:
    """Biconnected edge blocks with a block/vertex incidence forest."""

    blocks: tuple[frozenset[Edge], ...]
    links: dict[tuple[str, int], frozenset[tuple[str, int]]]

    @classmethod
    def build(cls, edges: Iterable[Edge]) -> PathBlocks:
        """Build in O(V+E), using the repository's existing networkx dependency."""
        graph = nx.Graph()
        graph.add_edges_from(sorted(edges))
        blocks = tuple(sorted(
            (frozenset(edge_key(a, b) for a, b in block)
             for block in nx.biconnected_component_edges(graph)),
            key=lambda block: min(block)))
        links: dict[tuple[str, int], set[tuple[str, int]]] = {}
        for bi, block in enumerate(blocks):
            bn = ("block", bi)
            vertices = {v for edge in block for v in edge}
            links[bn] = {("node", v) for v in vertices}
            for v in vertices:
                links.setdefault(("node", v), set()).add(bn)
        return cls(blocks, {key: frozenset(v) for key, v in links.items()})

    def paths(self, root: int, terminals: Iterable[int]) -> frozenset[Edge]:
        """Retain every edge on a simple root→terminal path, not dangling rings."""
        terminals = set(terminals)
        if not terminals:
            return frozenset()
        required = {("node", v) for v in terminals | {root}}
        degree = {key: len(v) for key, v in self.links.items()}
        removed = set()
        queue = deque(k for k, d in degree.items() if d <= 1 and k not in required)
        while queue:
            key = queue.popleft()
            if key in removed:
                continue
            removed.add(key)
            for other in self.links[key]:
                if other not in removed:
                    degree[other] -= 1
                    if degree[other] <= 1 and other not in required:
                        queue.append(other)
        return frozenset(e for bi, block in enumerate(self.blocks)
                         if ("block", bi) not in removed for e in block)


@dataclass(frozen=True)
class NetworkNode:
    """Source vertex; unknown real elevation must not be silently set to zero."""

    id: int
    x_real_mm: float
    y_real_mm: float
    z_real_m: float | None = None


@dataclass(frozen=True)
class NetworkPipe:
    """One source edge; endpoints are identities, not solved flow directions."""

    id: str
    node_a: int
    node_b: int
    plan_length_m: float
    inferred_connection: bool
    length_m: float | None = None
    rise_m: float | None = None
    bore_m: float | None = None
    flow_m3s: float | None = None


@dataclass(frozen=True)
class NetworkNozzle:
    """Physical head attachment and separate operating-scenario selection."""

    id: int
    attachment_node: int
    selected: bool
    flow_m3s: float | None = None


@dataclass(frozen=True)
class FlowNetwork:
    """Preserved source connectivity; the tree is ONLY a reference for ranking."""

    mode: str
    reference: FlowTree
    blocks: PathBlocks
    edges: frozenset[Edge]
    excluded: dict[Edge, str]
    inferred: frozenset[Edge]
    revision: str

    @classmethod
    def build(cls, reference: FlowTree, mode: str, *,
              inferred_edges: Iterable[Edge] = ()) -> FlowNetwork:
        """Preserve alternate paths without inventing ports, crossings or flow."""
        mode = network_mode(mode)
        if mode == "tree" or len(reference.roots) != 1:
            raise ValueError("연결 보존 추출에는 루프/그리드 방식과 급수원 하나가 필요합니다.")
        reachable = {e for e in reference.lengths_mm
                     if e[0] in reference.distance_mm and e[1] in reference.distance_mm}
        blocks = PathBlocks.build(reachable)
        heads = {reference.head_node[h] for h in set(reference.representatives.values())}
        edges = blocks.paths(reference.roots[0], heads)
        excluded = {e: "unreachable" if e not in reachable else "no_head_path"
                    for e in reference.lengths_mm if e not in edges}
        inferred = frozenset(edge_key(*e) for e in inferred_edges)
        revision = hashlib.sha256(repr(("network-v1", mode, reference.revision,
                                       sorted(edges), sorted(inferred))).encode()).hexdigest()[:16]
        return cls(mode, reference, blocks, edges, excluded, inferred, revision)

    def selected_edges(self, heads: Iterable[int]) -> frozenset[Edge]:
        """Prune non-operating dead ends but preserve ALL alternate supply paths."""
        requested = set(heads)
        if not requested <= self.reference.head_node.keys():
            raise ValueError("선정 헤드가 현재 급수원에 연결되어 있지 않습니다.")
        refs = {self.reference.representatives[h] for h in requested}
        nodes = {self.reference.head_node[h] for h in refs}
        return self.blocks.paths(self.reference.roots[0], nodes)

    def resolve_selection(self, heads: Iterable[int], *,
                          source_edges: Iterable[Edge] | None = None) -> frozenset[Edge]:
        """Validate an explicit scoped graph without expanding external loops."""
        requested = set(heads)
        if not requested or not requested <= self.reference.head_node.keys():
            raise ValueError("선정 헤드가 현재 급수원에 연결되어 있지 않습니다.")
        if source_edges is None:
            return self.selected_edges(requested)
        edges = frozenset(edge_key(*e) for e in source_edges)
        if not edges <= self.edges:
            raise ValueError("추출 범위에 원본 공급망에 없는 배관이 있습니다.")
        graph = nx.Graph()
        graph.add_edges_from(edges)
        graph.add_nodes_from(self.reference.roots)
        graph.add_nodes_from(self.reference.head_node[h] for h in requested)
        if not nx.is_connected(graph):
            raise ValueError("추출 범위의 배관과 헤드가 급수원까지 연결되지 않았습니다.")
        return edges

    def report(self) -> dict:
        """Audit summary; cycle rank is NOT a sprinkler layout classification."""
        nodes = {v for edge in self.edges for v in edge} | set(self.reference.roots)
        return {"revision": self.revision, "roots": list(self.reference.roots),
                "network_mode": self.mode, "layout_basis": "user_selected",
                "heads": len(set(self.reference.representatives.values())),
                "edges": len(self.edges), "excluded": dict(Counter(self.excluded.values())),
                "cycle_rank": len(self.edges) - len(nodes) + 1,
                "inferred_connections": len(self.inferred & self.edges),
                "length_basis": "source_plan_not_3d",
                "max_load_all": None, "hydraulics_solved": False,
                "calculation_ready": False, "calculation_note": CALCULATION_PENDING,
                "length_m": round(sum(self.reference.lengths_mm[e] for e in self.edges) / 1000, 3)}

    def extraction(self, points: Sequence[Sequence[float]], *,
                   selected_heads: Iterable[int] | None = None,
                   source_edges: Iterable[Edge] | None = None) -> dict:
        """Export typed, inspectable 2D intermediate data, never a pretend SDF."""
        ref = self.reference
        requested = set(ref.head_node) if selected_heads is None else set(selected_heads)
        if not requested <= ref.head_node.keys():
            raise ValueError("선정 헤드가 현재 급수원에 연결되어 있지 않습니다.")
        chosen = {ref.representatives[h] for h in requested}
        edges = (self.resolve_selection(chosen, source_edges=source_edges)
                 if source_edges is not None else self.selected_edges(chosen))
        used = {v for e in edges for v in e} | set(ref.roots)
        nodes = [NetworkNode(n, float(points[n][0]), float(points[n][1])) for n in sorted(used)]
        if any(not math.isfinite(c) for n in nodes for c in (n.x_real_mm, n.y_real_mm)):
            raise ValueError("추출 노드 좌표는 유한한 mm 값이어야 합니다.")
        pipes = [NetworkPipe(f"E{a}-{b}", a, b, ref.lengths_mm[a, b] / 1000,
                             (a, b) in self.inferred) for a, b in sorted(edges)]
        nozzles = [NetworkNozzle(h, ref.head_node[h], h in chosen)
                   for h in sorted(set(ref.representatives.values())) if ref.head_node[h] in used]
        return {"schema": "module-f-source-network-v1", "calculation_ready": False,
                "stage": "plan_connectivity", "network_mode": self.mode,
                "revision": self.revision, "root": ref.roots[0],
                "scope": ("region_and_supply" if source_edges is not None else
                          "all_heads" if selected_heads is None else "selected_heads"),
                "orientation": "undirected_not_hydraulic_flow",
                "cycle_rank": len(edges) - len(used) + 1,
                "nodes": [asdict(n) for n in nodes], "pipes": [asdict(p) for p in pipes],
                "nozzles": [asdict(n) for n in nozzles],
                "pending": ["elevation_and_3d_lengths", "diameters", "hydraulic_flow_split",
                            "tee_branch_losses", "hydraulically_worst_area"],
                "note": CALCULATION_PENDING}
