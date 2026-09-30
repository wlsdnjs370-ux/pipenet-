"""Validated rectangle/freehand regions, in CAD millimetres, with inclusive edges."""
from __future__ import annotations

import json
import math
from typing import Any

from .freehand import enclosed_geometry

MAX_ZONES = 64
MAX_ZONE_POINTS = 512
COORD_LIMIT = 1e9


def normalize_zones(raw: Any) -> list:
    """Keep legacy rectangles and explicit polygons; never replace a polygon by a box."""
    if raw is None:
        return []
    if not isinstance(raw, (list, tuple)) or len(raw) > MAX_ZONES:
        raise ValueError(f"영역은 최대 {MAX_ZONES}곳의 목록이어야 합니다.")

    def number(v):
        if isinstance(v, bool):
            raise ValueError("영역 좌표는 유한한 숫자여야 합니다.")
        try:
            n = float(v)
        except (TypeError, ValueError, OverflowError) as exc:
            raise ValueError("영역 좌표는 유한한 숫자여야 합니다.") from exc
        if not math.isfinite(n) or abs(n) > COORD_LIMIT:
            raise ValueError("영역 좌표가 허용 범위를 벗어났습니다.")
        return n

    out = []
    for zone in raw:
        if isinstance(zone, dict):
            pts = zone.get("points")
            if zone.get("type") != "polygon" or not isinstance(pts, (list, tuple)) or not 3 <= len(pts) <= MAX_ZONE_POINTS:
                raise ValueError(f"자유곡선 영역은 3~{MAX_ZONE_POINTS}개 점이 필요합니다.")
            clean = []
            for p in pts:
                if not isinstance(p, (list, tuple)) or len(p) != 2:
                    raise ValueError("자유곡선 점은 [x,y] 형식이어야 합니다.")
                point = [number(p[0]), number(p[1])]
                if not clean or point != clean[-1]:
                    clean.append(point)
            if len(clean) > 1 and clean[0] == clean[-1]:
                clean.pop()
            enclosed = zone.get('fill_rule') == 'enclosed'
            if zone.get('fill_rule') not in (None, 'enclosed'):
                raise ValueError('알 수 없는 영역 채우기 방식입니다.')
            area = enclosed_geometry(tuple(map(tuple, clean))).area if enclosed else polygon_area(clean)
            if len(clean) < 3 or area <= 1e-8:
                raise ValueError("영역의 면적이 없습니다. 둘레를 넓게 그려주세요.")
            out.append({"type": "polygon", "points": clean, **({'fill_rule':'enclosed'} if enclosed else {})})
        elif isinstance(zone, (list, tuple)) and len(zone) == 4:
            a, b, c, d = map(number, zone)
            if a == c or b == d:
                raise ValueError("영역의 면적이 없습니다.")
            out.append([min(a, c), min(b, d), max(a, c), max(b, d)])
        else:
            raise ValueError("영역은 사각형 네 좌표 또는 자유곡선 점 목록이어야 합니다.")
    return out


def polygon_area(points: list) -> float:
    """Absolute shoelace area, translated to reduce large-coordinate cancellation."""
    if len(points) < 3:
        return 0.0
    ox, oy = points[0]
    return abs(sum((a[0]-ox)*(b[1]-oy) - (b[0]-ox)*(a[1]-oy)
                   for a, b in zip(points, points[1:] + points[:1]))) / 2


def zone_points(zone: list | dict) -> list:
    """Return boundary vertices (not the polygon's bounding box)."""
    if isinstance(zone, dict):
        return zone["points"]
    x0, y0, x1, y1 = zone
    return [[x0, y0], [x1, y0], [x1, y1], [x0, y1]]


def _distance(p, a, b) -> float:
    dx, dy = b[0]-a[0], b[1]-a[1]
    length2 = dx*dx + dy*dy
    t = max(0, min(1, ((p[0]-a[0])*dx+(p[1]-a[1])*dy)/length2)) if length2 else 0
    return math.hypot(p[0]-a[0]-t*dx, p[1]-a[1]-t*dy)


def zone_contains(zone: list | dict, point: tuple | list, margin_mm: float = 0) -> bool:
    """Even-odd polygon containment, including boundary; rectangles remain compatible."""
    x, y = float(point[0]), float(point[1])
    if not isinstance(zone, dict):
        a, b, c, d = zone
        return a-margin_mm <= x <= c+margin_mm and b-margin_mm <= y <= d+margin_mm
    if zone.get('fill_rule') == 'enclosed':
        return any(zone_contains({'points': ring}, point, margin_mm)
                   for ring in enclosed_geometry(tuple(map(tuple, zone['points']))).faces)
    pts = zone["points"]
    inside = False
    for a, b in zip(pts, pts[1:] + pts[:1]):
        if _distance((x, y), a, b) <= max(1e-7, margin_mm):
            return True
        if (a[1] > y) != (b[1] > y) and x < (b[0]-a[0])*(y-a[1])/(b[1]-a[1])+a[0]:
            inside = not inside
    return inside


def zones_signature(zones: list) -> str:
    """Stable full-precision fingerprint, including every freehand vertex."""
    return json.dumps(normalize_zones(zones), sort_keys=True, separators=(",", ":"))


class ZoneRegion:
    """HeadRegion-compatible union used by Module A's anchored selector in Module F."""

    def __init__(self, zones: list, margin_mm: float = 0):
        self.zones = normalize_zones(zones)
        self.margin_mm = margin_mm

    @property
    def pts(self) -> list:
        """All actual vertices for the surrounding pipe-work window only."""
        return [p for z in self.zones for p in zone_points(z)]

    def contains(self, point: tuple | list) -> bool:
        """Test the union without filling concave cut-outs."""
        return any(zone_contains(z, point, self.margin_mm) for z in self.zones)

    def dilate(self, mm: float) -> ZoneRegion:
        """Expand candidate-search margin without changing the stored user boundary."""
        if not math.isfinite(mm) or mm < 0:
            raise ValueError("영역 여유는 0 이상의 유한한 길이여야 합니다.")
        return ZoneRegion(self.zones, self.margin_mm + mm)

    def __bool__(self) -> bool:
        return bool(self.zones)
