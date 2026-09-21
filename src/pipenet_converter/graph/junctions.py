"""Normalize evidenced CAD junctions before choosing calculation paths.

Only explicitly linked coincident vertices and collinear overlapping links
sharing an endpoint are normalized. Unconnected crossings are never joined.
Coordinates are millimetres and are not modified.
"""
from __future__ import annotations

from collections.abc import Iterable, Sequence
from dataclasses import dataclass
import math

Edge = tuple[int, int]


@dataclass(frozen=True)
class JunctionNormalization:
    """Normalized source links plus inspectable vertex and overlap provenance."""

    edges: frozenset[Edge]
    aliases: dict[int, int]
    splits: tuple[tuple[int, int, int], ...]


def normalize_junctions(points: Sequence[Sequence[float]], edges: Iterable[Edge],
                        *, tolerance_mm: float = 1e-6) -> JunctionNormalization:
    """Remove duplicate through-links; retain the actual branch intersection.

    A zero-length edge is explicit connection evidence, not proximity guessing.
    A through-link is split only at a connected neighbour strictly inside it.
    Cross-product distance avoids cancellation at large drawing coordinates.
    """
    if not math.isfinite(tolerance_mm) or tolerance_mm <= 0:
        raise ValueError("Junction tolerance must be finite and positive")
    original = {tuple(sorted(e)) for e in edges if e[0] != e[1]}
    parents: dict[int, int] = {}

    def root(v: int) -> int:
        while v in parents:
            v = parents[v]
        return v

    current = original
    splits: list[tuple[int, int, int]] = []
    while True:
        # Splitting duplicate coincident hubs can expose another zero link;
        # resolve it in the same pass so applying this function is idempotent.
        for a, b in sorted(current):
            if math.dist(points[a][:2], points[b][:2]) <= tolerance_mm:
                ra, rb = root(a), root(b)
                if ra != rb:
                    parents[max(ra, rb)] = min(ra, rb)
        current = {tuple(sorted((root(a), root(b)))) for a, b in current
                   if root(a) != root(b)}
        adjacency: dict[int, set[int]] = {}
        for a, b in current:
            adjacency.setdefault(a, set()).add(b)
            adjacency.setdefault(b, set()).add(a)
        replacements = []
        for a, b in sorted(current):
            ax, ay = points[a][:2]
            dx, dy = points[b][0] - ax, points[b][1] - ay
            length = math.hypot(dx, dy)
            if length <= tolerance_mm:
                continue
            inside = []
            for h in (adjacency[a] | adjacency[b]) - {a, b}:
                if len(adjacency[h]) < 2:
                    continue
                wx, wy = points[h][0] - ax, points[h][1] - ay
                along = (wx * dx + wy * dy) / length
                lateral = abs(wx * dy - wy * dx) / length
                if tolerance_mm < along < length - tolerance_mm and lateral <= tolerance_mm:
                    inside.append((along, h))
            if inside:
                chain = [a] + [h for _, h in sorted(inside)] + [b]
                replacements.append(((a, b), chain))
        if not replacements:
            break
        for edge, chain in replacements:
            current.discard(edge)
            current.update(tuple(sorted((a, b))) for a, b in zip(chain, chain[1:]))
            splits.extend((edge[0], edge[1], h) for h in chain[1:-1])
    return JunctionNormalization(frozenset(current),
                                 {v: root(v) for v in parents}, tuple(splits))


def arc_connection_pairs(points: Sequence[Sequence[float]], arms: Sequence[int],
                         existing: Iterable[Edge], candidates: Iterable[Edge]) -> list[Edge]:
    """Join components within ONE arc symbol, never creating a local triangle.

    An existing short arm may be represented by two endpoints. Connecting both
    to the opposite arm fabricates two tees. Prefer the shorter candidate;
    existing connections and genuine loops outside this symbol are untouched.
    """
    local = set(arms)
    parents = {v: v for v in local}

    def root(v: int) -> int:
        while parents[v] != v:
            v = parents[v]
        return v

    def join(a: int, b: int) -> bool:
        ra, rb = root(a), root(b)
        if ra == rb:
            return False
        parents[max(ra, rb)] = min(ra, rb)
        return True

    for a, b in existing:
        if a in local and b in local:
            join(a, b)
    ordered = sorted({tuple(sorted(e)) for e in candidates},
                     key=lambda e: (math.dist(points[e[0]][:2], points[e[1]][:2]), e))
    return [(a, b) for a, b in ordered if join(a, b)]
