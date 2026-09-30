"""Conservative evidence for a centre tick inside superimposed arc symbols.

This does not remove source geometry. It excludes a proven, unselected symbol
stroke from physical-port evidence only. Selected pipes and explicit edits win.
"""
from __future__ import annotations

from dataclasses import dataclass
import math
from typing import Mapping, Sequence


@dataclass(frozen=True)
class ArcStroke:
    """Source-mm evidence for two dead-end halves of one drawn centre tick."""

    edges: frozenset[tuple[int, int]]
    drawn_path: tuple[int, ...]
    arc_xy: tuple[float, float]
    radius_mm: float


def _open(arc: Mapping, vector: Sequence[float]) -> bool:
    angle = math.degrees(math.atan2(vector[1], vector[0])) % 360
    return (angle - float(arc['sa'])) % 360 > float(arc['sweep']) + 1e-8


def _unit(vector: Sequence[float]) -> tuple[float, float]:
    length = math.hypot(*vector)
    return vector[0] / length, vector[1] / length


def overlapping_arc_stroke(
    points: Sequence[Sequence[float]], adjacency: Mapping[int, set[int]],
    drawn_adjacency: Mapping[int, set[int]], cluster: Sequence[int],
    symbols: Sequence[Mapping], *, selected_nodes: set[int],
    protected_nodes: set[int],
) -> ArcStroke | None:
    """Recognize a bounded, symmetric, dead-end centre tick with raw proof.

    Require a 240–300 degree bowl overlaid by two contained quarter arcs, a
    drawn straight tick wholly within half its radius, and three other external
    ports (opposite main ports and one open-side branch). A real fourth pipe,
    head, explicit edit, asymmetry or missing original line keeps the review flag.
    """
    cluster_set = set(cluster)
    protected_ends = selected_nodes | protected_nodes
    if cluster_set & protected_nodes:
        return None
    for bowl in symbols:
        if not 240 <= float(bowl['sweep']) <= 300:
            continue
        xy = float(bowl['cx']), float(bowl['cy'])
        radius = float(bowl['r'])
        if radius <= 0:
            continue
        quarters = []
        for arc in symbols:
            if not 60 <= float(arc['sweep']) <= 120:
                continue
            if (math.dist(xy, (arc['cx'], arc['cy'])) > .2 or
                    abs(float(arc['r']) - radius) > .2):
                continue
            if all((float(arc['sa']) + t * float(arc['sweep']) -
                    float(bowl['sa'])) % 360 <= float(bowl['sweep']) + 1e-6
                   for t in (0, .5, 1)):
                quarters.append(arc)
        if not any(abs((float(a['sa']) - float(b['sa'])) % 360 - 180) < 2
                   for a in quarters for b in quarters):
            continue
        external = [(n, other) for n in cluster for other in sorted(adjacency.get(n, set()))
                    if other not in cluster_set]
        short = [(n, other) for n, other in external
                 if other not in protected_ends
                 and len(adjacency.get(other, ())) == 1
                 and math.dist(points[n][:2], xy) <= .2
                 and 1 < math.dist(points[other][:2], xy) <= radius / 2]
        if len(short) != 2:
            continue
        ends = [other for _, other in short]
        vectors = [tuple(points[e][i] - xy[i] for i in (0, 1)) for e in ends]
        unit = [_unit(v) for v in vectors]
        if (sum(a * b for a, b in zip(*unit)) > -.999 or
                abs(math.hypot(*vectors[0]) - math.hypot(*vectors[1])) > .2 or
                _open(bowl, unit[0]) == _open(bowl, unit[1])):
            continue
        # Recover the original short LINE, possibly split at its centre before
        # joining. A mere inferred dead end is not sufficient evidence.
        path = [ends[0]]
        while len(path) <= 8 and path[-1] != ends[1]:
            candidates = drawn_adjacency.get(path[-1], set()) - set(path)
            if len(candidates) != 1:
                break
            other = next(iter(candidates))
            v = tuple(points[other][i] - xy[i] for i in (0, 1))
            if (other in protected_nodes or math.hypot(*v) > radius / 2 or
                    abs(v[0] * unit[0][1] - v[1] * unit[0][0]) > .2):
                break
            path.append(other)
        if path[-1] != ends[1] or len(drawn_adjacency.get(ends[1], ())) != 1:
            continue
        # Remaining real ports must leave the symbol and match tee + elbow.
        directions = []
        for n, other in external:
            if (n, other) in short:
                continue
            v = tuple(points[other][i] - points[n][i] for i in (0, 1))
            if math.hypot(*v) <= radius:
                break
            u = _unit(v)
            if not any(sum(a * b for a, b in zip(u, p)) > .996 for p in directions):
                directions.append(u)
        else:
            opened = [v for v in directions if _open(bowl, v)]
            closed = [v for v in directions if not _open(bowl, v)]
            if (len(opened) == 1 and len(closed) == 2 and
                    sum(a * b for a, b in zip(*closed)) < -.996):
                return ArcStroke(frozenset(tuple(sorted(e)) for e in short),
                                 tuple(path), xy, radius)
    return None
