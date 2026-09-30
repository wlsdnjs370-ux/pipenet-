"""Fill all bounded parts of a freehand stroke without using its convex hull.

Split intersections into a planar graph, walk its faces, and retain every
bounded face regardless of stroke direction. Repeated edges and open tails do
not create holes or add selectable area. All coordinates are CAD millimetres.
"""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from functools import lru_cache
import math

Point = tuple[float, float]
Segment = tuple[Point, Point]
EPS = 1e-7


@dataclass(frozen=True)
class EnclosedGeometry:
    """Disjoint bounded faces and the actual outer boundary of their union."""

    faces: tuple[tuple[Point, ...], ...]
    boundary: tuple[Segment, ...]
    area: float


def _cross(a: Point, b: Point) -> float:
    return a[0] * b[1] - a[1] * b[0]


@lru_cache(maxsize=32)
def enclosed_geometry(points: tuple[Point, ...]) -> EnclosedGeometry:
    """Polygonize a validated, at-most-512-point stroke, implicitly closing it.

    Normalizing/limiting external input belongs to ``regions.normalize_zones``.
    Translate first to keep large CAD origins out of intersection arithmetic.
    """
    if len(points) < 3:
        return EnclosedGeometry((), (), 0.0)
    ox, oy = points[0]
    local = [(x - ox, y - oy) for x, y in points]
    sides = [(a, b) for a, b in zip(local, local[1:] + local[:1]) if math.dist(a, b) > EPS]
    splits = [[(0.0, a), (1.0, b)] for a, b in sides]
    for i, (a, b) in enumerate(sides):
        v = (b[0] - a[0], b[1] - a[1])
        lv = math.hypot(*v)
        for j in range(i):
            c, d = sides[j]
            if (max(a[0], b[0]) + EPS < min(c[0], d[0]) or max(c[0], d[0]) + EPS < min(a[0], b[0])
                    or max(a[1], b[1]) + EPS < min(c[1], d[1]) or max(c[1], d[1]) + EPS < min(a[1], b[1])):
                continue
            w, q = (d[0] - c[0], d[1] - c[1]), (c[0] - a[0], c[1] - a[1])
            lw, det = math.hypot(*w), _cross(v, w)
            if abs(det) > 1e-12 * lv * lw:
                t, u = _cross(q, w) / det, _cross(q, v) / det
                if -EPS/lv <= t <= 1+EPS/lv and -EPS/lw <= u <= 1+EPS/lw:
                    t, u = max(0.0, min(1.0, t)), max(0.0, min(1.0, u))
                    p = (a[0] + t*v[0], a[1] + t*v[1])
                    splits[i].append((t, p))
                    splits[j].append((u, p))
            elif abs(_cross(q, v)) <= EPS * lv:
                # Split collinear overlaps too, so retracing is one graph edge.
                for index, start, vector, length, ends in ((i, a, v, lv, (c, d)), (j, c, w, lw, (a, b))):
                    for p in ends:
                        t = ((p[0]-start[0])*vector[0] + (p[1]-start[1])*vector[1]) / length**2
                        if -EPS/length <= t <= 1+EPS/length:
                            splits[index].append((max(0.0, min(1.0, t)), p))

    vertices: list[Point] = []
    ids: dict[Point, int] = {}
    neighbors: dict[int, set[int]] = {}

    def vertex(point: Point) -> int:
        key = (round(point[0], 6), round(point[1], 6))
        if key not in ids:
            ids[key] = len(vertices)
            vertices.append(point)
        return ids[key]

    for split in splits:
        ordered = [vertex(p) for _, p in sorted(split, key=lambda item: item[0])]
        for a, b in zip(ordered, ordered[1:]):
            if a != b:
                neighbors.setdefault(a, set()).add(b)
                neighbors.setdefault(b, set()).add(a)
    ordered_neighbors = {a: sorted(bs, key=lambda b: math.atan2(vertices[b][1]-vertices[a][1],
                                                               vertices[b][0]-vertices[a][0]))
                         for a, bs in neighbors.items()}
    next_edge = {}
    for b, around in ordered_neighbors.items():
        for index, a in enumerate(around):
            next_edge[a, b] = (b, around[index-1])
    visited: set[tuple[int, int]] = set()
    faces = []
    boundary_counts: Counter[tuple[int, int]] = Counter()
    total_area = 0.0
    for start in next_edge:
        if start in visited:
            continue
        edge, walk = start, []
        while edge not in visited:
            visited.add(edge)
            walk.append(edge)
            edge = next_edge[edge]
        if edge != start:
            raise ValueError('펜 영역의 둘레를 해석하지 못했습니다. 다시 그려주세요.')
        ring = [vertices[a] for a, _ in walk]
        x0, y0 = ring[0]
        area = sum((a[0]-x0)*(b[1]-y0) - (b[0]-x0)*(a[1]-y0)
                   for a, b in zip(ring, ring[1:] + ring[:1])) / 2
        if area > 1e-8:  # Exterior walks have the opposite orientation.
            faces.append(tuple((x+ox, y+oy) for x, y in ring))
            total_area += area
            boundary_counts.update(tuple(sorted(e)) for e in walk)
    boundary = tuple(((vertices[a][0]+ox, vertices[a][1]+oy), (vertices[b][0]+ox, vertices[b][1]+oy))
                     for (a, b), count in boundary_counts.items() if count == 1)
    return EnclosedGeometry(tuple(faces), boundary, total_area)
