"""Small, explicit CAD symbol geometry; no layer or head-type guesses."""
from __future__ import annotations

import math
from typing import Sequence

Point = tuple[float, float]


def two_arc_circle(points: Sequence[Sequence[float]], bulges: Sequence[float],
                   closed: bool) -> tuple[Point, float] | None:
    """Recognize exactly two semicircles forming a closed circle (CAD mm)."""
    if not closed or len(points) != 2 or len(bulges) != 2:
        return None
    a, b = map(float, bulges)
    if not (math.isclose(abs(a), 1., abs_tol=1e-8) and
            math.isclose(a, b, abs_tol=1e-8)):
        return None
    p, q = points
    radius = math.hypot(q[0]-p[0], q[1]-p[1])/2
    if not math.isfinite(radius) or radius <= 0:
        return None
    return ((float(p[0]+q[0])/2, float(p[1]+q[1])/2), radius)


def triangle_contains(point: Point, vertices: Sequence[Point], tolerance: float = 1e-7) -> bool:
    """Return whether a point is in a nondegenerate triangle, including its rim."""
    if len(vertices) != 3:
        return False
    distances = []
    for a, b in zip(vertices, (*vertices[1:], vertices[0])):
        length = math.dist(a, b)
        if length <= tolerance:
            return False
        distances.append(((b[0]-a[0])*(point[1]-a[1]) -
                          (b[1]-a[1])*(point[0]-a[0]))/length)
    return (min(distances) >= -tolerance or max(distances) <= tolerance) and \
        max(abs(d) for d in distances) > tolerance
