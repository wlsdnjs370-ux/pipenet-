"""Non-destructive drawing crop, before symbol detection or graph construction.

Lines are clipped exactly, including concave boundaries. Freehand enclosed-fill
regions preserve whole circles/arcs whose centers are inside. Legacy regions
keep their conservative boundary-symbol exclusion for exact saved replay.
Text uses its insertion point. No original DXF or full-world cache is modified.
"""
from __future__ import annotations

import copy
import math
from typing import Any, Sequence

from ..graph.regions import normalize_zones, zone_contains, zone_points, _distance
from ..graph.freehand import enclosed_geometry


def _cross(a, b) -> float:
    return a[0] * b[1] - a[1] * b[0]


class WorkRegion:
    """A validated region, optionally filling every part enclosed by a pen stroke."""

    def __init__(self, zone: dict | list):
        self.zone = normalize_zones([zone])[0]
        self.points = zone_points(self.zone)
        self.sides = list(zip(self.points, self.points[1:] + self.points[:1]))
        self.bounds = (min(p[0] for p in self.points), min(p[1] for p in self.points),
                       max(p[0] for p in self.points), max(p[1] for p in self.points))
        self.enclosed = isinstance(self.zone, dict) and self.zone.get('fill_rule') == 'enclosed'
        if self.enclosed:
            geometry = enclosed_geometry(tuple(map(tuple, self.points)))
            self.sides = geometry.boundary
            self.faces = [({'points': ring}, (min(p[0] for p in ring), min(p[1] for p in ring),
                                             max(p[0] for p in ring), max(p[1] for p in ring)))
                          for ring in geometry.faces]
            return
        for i, (a, b) in enumerate(self.sides):
            for j, (c, d) in enumerate(self.sides[:i]):
                if i == j + 1 or (j == 0 and i == len(self.sides) - 1):
                    continue
                v, w = (b[0]-a[0], b[1]-a[1]), (d[0]-c[0], d[1]-c[1])
                q = (c[0]-a[0], c[1]-a[1])
                det = _cross(v, w)
                if abs(det) > 1e-12:
                    t, u = _cross(q, w)/det, _cross(q, v)/det
                    if 0 <= t <= 1 and 0 <= u <= 1:
                        raise ValueError("영역 선이 서로 교차합니다. 겹치지 않는 둘레로 다시 그려주세요.")
                elif min(_distance(c, a, b), _distance(d, a, b),
                         _distance(a, c, d), _distance(b, c, d)) < 1e-7:
                    raise ValueError("영역 선이 겹칩니다. 둘레를 다시 그려주세요.")

    def contains(self, point: Sequence[float]) -> bool:
        """Boundary-inclusive containment in CAD millimetres."""
        x0, y0, x1, y1 = self.bounds
        x, y = point
        if not (x0-1e-7 <= x <= x1+1e-7 and y0-1e-7 <= y <= y1+1e-7):
            return False
        if self.enclosed:
            return any(a-1e-7 <= x <= c+1e-7 and b-1e-7 <= y <= d+1e-7 and zone_contains(face, point)
                       for face, (a, b, c, d) in self.faces)
        return zone_contains(self.zone, point)

    def clip(self, a: Sequence[float], b: Sequence[float]) -> list:
        """Return only interior pieces; never join across a concave cut-out."""
        x0, y0, x1, y1 = self.bounds
        if max(a[0], b[0]) < x0 or min(a[0], b[0]) > x1 or max(a[1], b[1]) < y0 or min(a[1], b[1]) > y1:
            return []
        v = (b[0]-a[0], b[1]-a[1])
        length2 = v[0]**2 + v[1]**2
        if length2 < 1e-14:
            return []
        ts = {0.0, 1.0}
        for c, d in self.sides:
            w, q = (d[0]-c[0], d[1]-c[1]), (c[0]-a[0], c[1]-a[1])
            det = _cross(v, w)
            if abs(det) > 1e-12:
                t, u = _cross(q, w)/det, _cross(q, v)/det
                if -1e-9 <= t <= 1+1e-9 and -1e-9 <= u <= 1+1e-9:
                    ts.add(max(0.0, min(1.0, t)))
            elif abs(_cross(q, v)) < 1e-7:
                for p in (c, d):
                    ts.add(max(0.0, min(1.0, ((p[0]-a[0])*v[0]+(p[1]-a[1])*v[1])/length2)))
        def at(t):
            return (a[0]+t*v[0], a[1]+t*v[1])
        parts = []
        cuts = sorted(ts)
        for t, u in zip(cuts, cuts[1:]):
            if u-t > 1e-10 and self.contains(at((t+u)/2)):
                if parts and math.dist(parts[-1][1], at(t)) < 1e-7:
                    parts[-1] = (parts[-1][0], at(u))
                else:
                    parts.append((at(t), at(u)))
        return parts

    def whole_circle(self, x: float, y: float, radius: float) -> bool:
        """Keep complete symbols; new pen regions use center-inclusive selection."""
        if self.enclosed:
            return self.contains((x, y))
        return self.contains((x, y)) and all(_distance((x, y), a, b) >= radius-1e-7 for a, b in self.sides)

    def whole_segment(self, a: Sequence[float], b: Sequence[float]) -> bool:
        """Reject inferred shortcuts crossing a previously removed region."""
        if math.dist(a,b)<1e-7:
            return self.contains(a)
        parts=self.clip(a,b)
        return len(parts)==1 and math.dist(parts[0][0],a)<1e-6 and math.dist(parts[0][1],b)<1e-6


def crop_world(world: Any, zone: dict | list) -> tuple[Any, dict]:
    """Return an independent cropped world and counts, preserving source metadata."""
    region = WorkRegion(zone)
    result = copy.copy(world)
    before = {k: len(getattr(world, k, ())) for k in ('segs', 'circles', 'arcs', 'texts')}
    for key in ('segs', 'raw_segs'):
        setattr(result, key, [(ly, col, aa, bb) for ly, col, a, b in getattr(world, key, ())
                             for aa, bb in region.clip(a, b)])
    result.circles = [c for c in world.circles if region.whole_circle(*c[2:5])]
    indices = [i for i, c in enumerate(world.arcs) if region.whole_circle(*c[2:5])]
    result.arcs = [world.arcs[i] for i in indices]
    angles = getattr(world, 'arc_ang', ())
    result.arc_ang = [angles[i] if i < len(angles) else None for i in indices]
    result.texts = [t for t in world.texts if region.contains(t[2:4])]
    result._work_regions = list(getattr(world, '_work_regions', ())) + [region.zone]
    after = {k: len(getattr(result, k)) for k in before}
    return result, {'before': before, 'after': after, 'regions': result._work_regions,
                    'boundary_symbols_excluded': sum(region.contains(c[2:4]) and not region.whole_circle(*c[2:5])
                                                     for c in [*world.circles, *world.arcs])}
