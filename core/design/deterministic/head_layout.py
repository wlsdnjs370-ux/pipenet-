# -*- coding: utf-8 -*-
"""C4 — 헤드 배치 (지시서 §8).

실 하나에 격자를 여러 번 얹어 보고 가장 적은 헤드로 전부 덮는 배치를 고른다.
격자는 **정방형·장방형만** 쓴다(§8.2). 육각 배치가 피복은 낫지만 헤드가 어긋나
흩어지면 C5 의 가지배관이 직선으로 못 가고, 헤드 몇 개 아끼는 이득보다 배관
복잡도 손해가 크다.

살수장애는 두 갈래다(§8.4). 덕트·조명·배관은 반경 60cm 룰(2.7.7.1)이지만 **보는
아니다** — 보는 수평거리 0.75m 미만에서도 반사판을 보 하단보다 낮추는 조건으로
허용되므로(2.7.7.7 별표), 60cm 룰에 넣으면 보가 있는 모든 실에서 대량 오탐이 난다.

장애물 정보가 없으면 검사를 건너뛰되 조용히 통과시키지 않는다(§8.5). 검증하지
않은 것과 검증해서 통과한 것은 다르다 — `OBSTACLE_UNVERIFIED` 로 남긴다.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any

from ..recognize.spatial import (
    point_in_polygon, polygon_area, representative_point)
from .nftc_tables import beam_clearance

_MM_PER_M = 1000.0

# §8.1 sweep 규모 — 격자 주기를 8 등분해 오프셋을 훑는다.
_OFFSET_STEPS = 8

# 헤드를 **놓아 볼** 자리의 간격(R/4). 피복 판정에는 쓰지 않는다 — 판정은
# `coverage_witnesses` 가 표본 없이 한다. 큰 실에서 후보가 폭발하면 늘린다.
_SPOT_BUDGET = 4000

# 피복 판정 여유. 격자 대각선이 정확히 2R 이라 칸 중앙은 딱 R 이고, 증인점은
# 자기를 만든 원 위에 놓인다 — 여유가 없으면 부동소수 끝자리가 판정을 뒤집는다.
# 1mm 는 도면 좌표 단위이자 헤드 설치 오차보다 작은 값이다.
_COVERAGE_TOL_MM = 1.0

# 벽 이격 metric 은 그 벽 **앞에 헤드가 있어야 할 만큼 긴 벽**에서만 잰다.
# 0.3m 짜리 꺾임 앞의 이격을 보고해 봐야 읽는 사람만 헷갈린다.
_WALL_MIN_LEN_FACTOR = 1.0

# 장방형 격자의 변 비(§8.2). 대각선은 항상 2R 로 고정하고 비만 바꾼다 —
# 대각선을 2R 보다 좁히면 같은 피복에 헤드만 는다.
#
# 1.5 에서 끊는 근거: 비가 1.5 면 가지배관 간격이 헤드 간격의 2/3 이고, 그보다
# 더 벌리면 헤드 한두 개 아낀 이득을 가지배관 물량이 도로 먹는다(§8.2 의 취지).
# 역수까지 넣는 것은 축을 하나만 훑기 때문이다 — 90° 돌린 격자가 곧 역수다.
_GRID_RATIOS = (1.0, 1.15, 1.3, 1.5, 1.0 / 1.15, 1.0 / 1.3, 1.0 / 1.5)


@dataclass
class Head:
    """§8.3 — `row`/`col` 은 C5 가 가지배관 축을 정하는 근거다.

    `branch_axis` 는 여기서 채우지 않는다. 축은 구역의 OBB 와 밸브 위치를 함께
    봐야 정해지고(§9.2 C510), 그 둘은 C4 가 보는 범위 밖이다. 모르는 것을 "x" 로
    적어 두면 C5 가 이미 정해진 축인 줄 알고 그대로 쓴다.
    """

    id: str
    room_id: str
    x: float
    y: float
    row: int
    col: int
    branch_axis: str | None = None
    provenance: str = "grid"

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id, "room_id": self.room_id,
            "x": round(self.x, 1), "y": round(self.y, 1),
            "row": self.row, "col": self.col,
            "branch_axis": self.branch_axis, "provenance": self.provenance,
        }


@dataclass
class RoomLayout:
    """실 하나의 배치 결과. `flags` 가 비어야 배치가 근거를 갖춘 것이다."""

    room_id: str
    heads: list = field(default_factory=list)
    flags: list = field(default_factory=list)
    metrics: dict = field(default_factory=dict)
    beam_requirements: list = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "room_id": self.room_id,
            "heads": [h.to_dict() for h in self.heads],
            "flags": self.flags, "metrics": self.metrics,
            "beam_requirements": self.beam_requirements,
        }


# ────────────────────────────────────────────────────────────────────────────
# 기하
# ────────────────────────────────────────────────────────────────────────────

def _hull(points: list) -> list:
    """단조 사슬 볼록껍질. 최소 외접 사각형의 후보 각도를 여기서 얻는다."""
    pts = sorted(set((float(p[0]), float(p[1])) for p in points))
    if len(pts) < 3:
        return pts

    def build(seq):
        out: list = []
        for p in seq:
            while len(out) >= 2:
                (x1, y1), (x2, y2) = out[-2], out[-1]
                if (x2 - x1) * (p[1] - y1) - (y2 - y1) * (p[0] - x1) > 0:
                    break
                out.pop()
            out.append(p)
        return out[:-1]

    return build(pts) + build(reversed(pts))


def principal_axes(polygon: list) -> list[float]:
    """실의 주축 각도(rad) 둘. 최소 외접 사각형의 변 방향이다.

    무게중심 기준 관성축을 쓰면 ㄱ자 실에서 벽과 어긋난 축이 나온다. 벽과 나란한
    격자여야 가지배관이 벽을 따라간다.
    """
    hull = _hull(polygon)
    if len(hull) < 3:
        return [0.0, math.pi / 2]

    best = None
    for i, (x1, y1) in enumerate(hull):
        x2, y2 = hull[(i + 1) % len(hull)]
        theta = math.atan2(y2 - y1, x2 - x1)
        cos_t, sin_t = math.cos(-theta), math.sin(-theta)
        us = [p[0] * cos_t - p[1] * sin_t for p in hull]
        vs = [p[0] * sin_t + p[1] * cos_t for p in hull]
        area = (max(us) - min(us)) * (max(vs) - min(vs))
        if best is None or area < best[0]:
            best = (area, theta)

    theta = best[1] % (math.pi / 2)
    return [theta, theta + math.pi / 2]


def _rotate(x: float, y: float, theta: float) -> tuple[float, float]:
    cos_t, sin_t = math.cos(theta), math.sin(theta)
    return (x * cos_t - y * sin_t, x * sin_t + y * cos_t)


def _point_to_segment(px: float, py: float, ax: float, ay: float,
                      bx: float, by: float) -> float:
    dx, dy = bx - ax, by - ay
    span = dx * dx + dy * dy
    t = 0.0 if span == 0 else max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / span))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def distance_to_boundary(point, polygon: list) -> float:
    px, py = point[0], point[1]
    return min(
        _point_to_segment(px, py, polygon[i][0], polygon[i][1],
                          polygon[(i + 1) % len(polygon)][0],
                          polygon[(i + 1) % len(polygon)][1])
        for i in range(len(polygon)))


def wall_gap_mm(heads: list, polygon: list, min_edge_mm: float) -> float:
    """벽에서 가장 가까운 헤드까지의 **수직**거리 중 최댓값(2.7.3).

    점거리로 재면 안 된다 — 벽을 따라 헤드 사이 중간 지점은 어떤 적법한 배치에서도
    S/2 를 넘는다. 조문이 말하는 것은 벽과 헤드 열 사이의 간격이다. 그래서 벽에
    수직으로 투영되는 헤드만 그 벽의 것으로 센다.

    긴 벽인데 앞에 헤드가 하나도 없으면 `inf` 다 — 0 으로 접으면 헤드가 없는 벽이
    가장 좋은 배치로 읽힌다.
    """
    worst = 0.0
    for i in range(len(polygon)):
        ax, ay = float(polygon[i][0]), float(polygon[i][1])
        bx, by = float(polygon[(i + 1) % len(polygon)][0]), float(polygon[(i + 1) % len(polygon)][1])
        span = math.hypot(bx - ax, by - ay)
        if span < min_edge_mm:
            continue
        ux, uy = (bx - ax) / span, (by - ay) / span
        near = min(
            (abs((h.x - ax) * -uy + (h.y - ay) * ux)
             for h in heads if 0.0 <= (h.x - ax) * ux + (h.y - ay) * uy <= span),
            default=float("inf"))
        worst = max(worst, near)
    return worst


# ────────────────────────────────────────────────────────────────────────────
# 격자 sweep (§8.1)
# ────────────────────────────────────────────────────────────────────────────

def grid_spacings(radius_mm: float) -> list[tuple[float, float]]:
    """(S₁, S₂) 후보. 대각선이 정확히 2R 인 조합만 — §8.2 의 장방형 허용 범위다."""
    diag = 2.0 * radius_mm
    return [(diag * r / math.hypot(1.0, r), diag / math.hypot(1.0, r))
            for r in _GRID_RATIOS]


def _spots(polygon: list, step_mm: float, wall_min: float) -> list:
    """헤드를 더 놓아 볼 자리. 벽에서 `wall_min` 이상 떨어진 내부 점만.

    안 덮인 점 자체를 자리로 쓰면 구석에서 벽에 붙은 헤드가 나오는데, 그 자리에는
    실제로 헤드를 못 단다(2.7.7.2).
    """
    xs = [float(p[0]) for p in polygon]
    ys = [float(p[1]) for p in polygon]
    width, height = max(xs) - min(xs), max(ys) - min(ys)
    while (width / step_mm + 1) * (height / step_mm + 1) > _SPOT_BUDGET:
        step_mm *= 2.0

    points = []
    y = min(ys) + step_mm * 0.5
    while y < max(ys):
        x = min(xs) + step_mm * 0.5
        while x < max(xs):
            if (point_in_polygon((x, y), polygon)
                    and distance_to_boundary((x, y), polygon) >= wall_min):
                points.append((x, y))
            x += step_mm
        y += step_mm
    return points


def _circle_segment_points(cx: float, cy: float, r: float,
                           ax: float, ay: float, bx: float, by: float):
    dx, dy = bx - ax, by - ay
    a = dx * dx + dy * dy
    if a == 0.0:
        return
    b = 2.0 * (dx * (ax - cx) + dy * (ay - cy))
    c = (ax - cx) ** 2 + (ay - cy) ** 2 - r * r
    disc = b * b - 4.0 * a * c
    if disc < 0.0:
        return
    root = math.sqrt(disc)
    for t in ((-b - root) / (2.0 * a), (-b + root) / (2.0 * a)):
        if 0.0 <= t <= 1.0:
            yield (ax + dx * t, ay + dy * t)


def _circle_circle_points(x1: float, y1: float, x2: float, y2: float, r: float):
    dx, dy = x2 - x1, y2 - y1
    d = math.hypot(dx, dy)
    if d == 0.0 or d > 2.0 * r:
        return
    h = math.sqrt(max(0.0, r * r - d * d * 0.25))
    mx, my = (x1 + x2) * 0.5, (y1 + y2) * 0.5
    ux, uy = -dy / d * h, dx / d * h
    yield (mx + ux, my + uy)
    yield (mx - ux, my - uy)


def coverage_witnesses(points: list, polygon: list, radius: float) -> list:
    """수평거리 R 밖에 남은 점. 비어 있으면 **증명된** 100% 피복이다.

    [문서정합 §8.1] 명세는 실을 `R/4` 격자로 표본화해 피복을 확인하라고 적었다.
    그렇게 하면 서로 다른 헤드에 덮인 두 표본점 **사이**가 비어도 통과한다. 실제로
    10m 정사각 실에서 R/4 표본이 100% 로 통과시킨 배치의 최악점이 2.37m(> R=2.3m)
    였다. 표본을 촘촘히 해도 오차가 줄 뿐 없어지지 않고, 오차만큼 R 을 깎으면
    이번에는 헤드가 헛되이 는다.

    안 덮인 영역이 있으면 그 경계의 꼭짓점은 반드시 폴리곤 꼭짓점이거나, 원과
    변의 교점이거나, 두 원의 교점이다. 그 후보만 보면 표본 없이 판정된다.
    """
    tol = radius + _COVERAGE_TOL_MM
    grid = _PointGrid(radius)
    for x, y in points:
        grid.add(x, y)

    pairs = _PointGrid(2.0 * radius)
    for x, y in points:
        pairs.add(x, y)

    n = len(polygon)
    edges = [(float(polygon[i][0]), float(polygon[i][1]),
              float(polygon[(i + 1) % n][0]), float(polygon[(i + 1) % n][1]))
             for i in range(n)]

    # 경계 위의 후보는 폴리곤 안팎을 다시 묻지 않는다 — 변 위의 점에 대해
    # 반직선 판정은 보장이 없어서, 물으면 구석의 진짜 구멍이 도로 걸러진다.
    holes = [(ax, ay) for ax, ay, _bx, _by in edges
             if not grid.within(ax, ay, tol)]
    for cx, cy in points:
        for edge in edges:
            holes.extend(p for p in _circle_segment_points(cx, cy, radius, *edge)
                         if not grid.within(p[0], p[1], tol))
        for ox, oy in pairs.near(cx, cy, 2.0 * radius):
            if (ox, oy) != (cx, cy):
                holes.extend(
                    p for p in _circle_circle_points(cx, cy, ox, oy, radius)
                    if not grid.within(p[0], p[1], tol)
                    and point_in_polygon(p, polygon))
    return holes


class _PointGrid:
    """반경 검색용 격자. 실이 커지면 전수 비교는 후보마다 수십만 번이 된다.

    셀 크기를 검색 반경으로 잡아 3x3 이웃만 보면 된다.
    """

    def __init__(self, cell: float):
        self.cell = cell
        self.buckets: dict[tuple[int, int], list] = {}

    def add(self, x: float, y: float) -> None:
        self.buckets.setdefault((int(x // self.cell), int(y // self.cell)), []).append((x, y))

    def _neighbourhood(self, x: float, y: float):
        cx, cy = int(x // self.cell), int(y // self.cell)
        for i in range(cx - 1, cx + 2):
            for j in range(cy - 1, cy + 2):
                yield from self.buckets.get((i, j), ())

    def within(self, x: float, y: float, radius: float) -> bool:
        r2 = radius * radius
        return any((px - x) ** 2 + (py - y) ** 2 <= r2
                   for px, py in self._neighbourhood(x, y))

    def near(self, x: float, y: float, radius: float) -> list:
        r2 = radius * radius
        return [p for p in self._neighbourhood(x, y)
                if (p[0] - x) ** 2 + (p[1] - y) ** 2 <= r2]


def _place_grid(polygon: list, theta: float, ox: float, oy: float,
                s1: float, s2: float) -> list[tuple[float, float, int, int]]:
    """회전 격자에서 실 안에 떨어지는 점만. `(x, y, row, col)`."""
    us, vs = zip(*(_rotate(float(p[0]), float(p[1]), -theta) for p in polygon))
    out = []
    col0 = math.floor((min(us) - ox) / s1)
    row0 = math.floor((min(vs) - oy) / s2)
    col1 = math.ceil((max(us) - ox) / s1)
    row1 = math.ceil((max(vs) - oy) / s2)
    for row in range(row0, row1 + 1):
        v = oy + row * s2
        for col in range(col0, col1 + 1):
            x, y = _rotate(ox + col * s1, v, theta)
            if point_in_polygon((x, y), polygon):
                out.append((x, y, row, col))
    return out


def _complete(placed: list, polygon: list, spots: _PointGrid, radius: float,
              theta: float, ox: float, oy: float, s1: float, s2: float,
              limit: int) -> list | None:
    """격자가 못 덮은 곳을 메운다. 남은 구멍을 가장 많이 덮는 자리를 골라 반복한다.

    자리 후보를 매번 전부 채점하지 않는 이유는 sweep 이 실당 448회 돌기 때문이다.
    남은 첫 구멍을 덮을 수 있는 자리로 좁히면 한 번 고를 때마다 그 구멍이 반드시
    사라지므로 진행이 보장되고, 비용은 실 크기가 아니라 반경 안 후보 수에 걸린다.

    `limit` 개를 넘게 되면 `None` 이다 — 이미 더 나은 배치가 있으니 끝까지 메워
    봐야 버릴 후보다. 이 조기 포기가 sweep 을 실용 속도로 만든다.

    더한 헤드의 `row`/`col` 은 격자에 반올림해 붙인다. 격자 위가 아니므로 기존
    헤드와 같은 칸을 가리킬 수 있는데, C5 는 이 인덱스를 열 묶음에만 쓰므로
    같은 열에 하나 더 붙는 것으로 읽힌다.
    """
    r2 = radius * radius
    added: list = []
    used: set = set()
    while True:
        holes = coverage_witnesses([(p[0], p[1]) for p in placed + added],
                                   polygon, radius)
        if not holes:
            return added
        if len(added) >= limit:
            return None
        pool = [p for p in spots.near(holes[0][0], holes[0][1], radius)
                if p not in used]
        if not pool:
            return added   # 메울 자리가 없다. 호출자가 플래그로 남긴다.
        cx, cy = max(pool, key=lambda c: sum(
            1 for h in holes if (h[0] - c[0]) ** 2 + (h[1] - c[1]) ** 2 <= r2))
        used.add((cx, cy))
        u, v = _rotate(cx, cy, -theta)
        added.append((cx, cy, round((v - oy) / s2), round((u - ox) / s1)))


def layout_heads(room, constraints, *, room_index: int = 0) -> RoomLayout:
    """실 하나에 헤드를 놓는다. 살수장애는 보지 않는다 — `check_obstacles` 가 본다."""
    layout = RoomLayout(room_id=room.id)
    polygon = [(float(p[0]), float(p[1])) for p in (room.polygon or [])]
    if len(polygon) < 3:
        layout.flags.append({"code": "ROOM_POLYGON_INVALID", "room": room.id,
                             "message": f"{room.id}: 실 폴리곤이 3점 미만이라 배치할 수 없습니다."})
        return layout
    if room.head_exempt:
        layout.metrics["head_exempt"] = True
        return layout

    radius = constraints.horizontal_distance_m * _MM_PER_M
    wall_min = constraints.head_to_wall_clearance_m * _MM_PER_M
    spots = _PointGrid(radius)
    for point in _spots(polygon, radius / 4.0, wall_min):
        spots.add(point[0], point[1])

    # [문서정합 §8.1] 명세는 축 2개 × 오프셋 8 × 8 = 128 후보를 훑으라고 적었다.
    # 축을 90° 돌린 격자는 (S₁, S₂) 를 맞바꾼 격자와 같으므로, 변 비를 역수까지
    # 훑는 여기서는 축 하나로 두 축을 다 본 것이 된다. 같은 후보를 두 번 채점하지
    # 않는다. 대신 §8.2 가 허용하는 장방형까지 보므로 후보 수는 7 × 64 로 는다.
    theta = principal_axes(polygon)[0]
    grids = []
    for s1, s2 in grid_spacings(radius):
        for i in range(_OFFSET_STEPS):
            for j in range(_OFFSET_STEPS):
                ox, oy = s1 * i / _OFFSET_STEPS, s2 * j / _OFFSET_STEPS
                placed = _place_grid(polygon, theta, ox, oy, s1, s2)
                grids.append((len(placed), placed, ox, oy, s1, s2))

    # 격자만으로 적은 것부터 본다. 채우기는 헤드를 늘리기만 하므로, 먼저 좋은
    # 답을 잡아 두면 나머지 후보는 채우기 전에 잘린다.
    best = None
    for count, placed, ox, oy, s1, s2 in sorted(grids, key=lambda g: g[0]):
        if best is not None and count > best[0][0]:
            break
        added = _complete(placed, polygon, spots, radius, theta, ox, oy, s1, s2,
                          limit=(best[0][0] - count) if best else len(spots.buckets) * 9)
        if added is None:
            continue
        whole = placed + added
        score = (len(whole), _clearance_variance(whole, polygon))
        if best is None or score < best[0]:
            best = (score, whole, s1, s2)

    placed = best[1] if best else []
    if not placed:
        # H3 — 면적과 무관하게 실마다 최소 한 개. 격자점이 하나도 안 걸리는 좁은
        # 실에서도 헤드 없는 실을 내보내면 그 실은 아무도 못 지킨다.
        point = representative_point(polygon)
        if point is None:
            layout.flags.append({"code": "HEAD_PLACEMENT_FAILED", "room": room.id,
                                 "message": f"{room.id}: 실 안의 점을 찾지 못해 헤드를 놓지 못했습니다."})
            return layout
        placed = [(point[0], point[1], 0, 0)]

    rows = sorted({p[2] for p in placed})
    cols = sorted({p[3] for p in placed})
    layout.heads = [
        Head(id=f"H-{room_index:03d}-{n:03d}", room_id=room.id, x=x, y=y,
             row=rows.index(row), col=cols.index(col))
        for n, (x, y, row, col) in enumerate(sorted(placed, key=lambda p: (p[2], p[3])), 1)
    ]

    s1, s2 = (best[2], best[3]) if best else (0.0, 0.0)
    gap = wall_gap_mm(layout.heads, polygon, max(s1, s2) * _WALL_MIN_LEN_FACTOR)
    layout.metrics.update({
        "area_m2": round(polygon_area(polygon) / (_MM_PER_M ** 2), 2),
        "head_count": len(layout.heads),
        "wall_gap_m": None if math.isinf(gap) else round(gap / _MM_PER_M, 3),
        "grid_spacing_m": [round(s1 / _MM_PER_M, 3), round(s2 / _MM_PER_M, 3)],
        "axis_deg": round(math.degrees(theta), 2),
    })

    # 벽 이격(2.7.3)은 따로 판정하지 않는다. 벽 위의 점까지 R 안이면 벽에서
    # 헤드까지의 수직거리는 그 벽을 따라 놓인 헤드 간격의 절반 이하가 되기
    # 때문이다(S₁²+S₂²=(2R)²). 정방형 상한(S/2)을 장방형 배치에 그대로 대면
    # 적법한 배치가 걸린다 — 그래서 이격은 metric 으로만 남기고, 판정은 피복이
    # 실제로 뚫린 경우에만 한다.
    holes = coverage_witnesses([(h.x, h.y) for h in layout.heads], polygon, radius)
    if holes:
        layout.metrics["coverage_hole_at"] = [round(holes[0][0], 1), round(holes[0][1], 1)]
        layout.flags.append({
            "code": "HEAD_COVERAGE_GAP", "room": room.id,
            "message": f"{room.id}: 수평거리 {constraints.horizontal_distance_m}m 밖으로 "
                       f"남는 곳이 있습니다 (예: {holes[0][0]:.0f}, {holes[0][1]:.0f}).",
        })
    return layout


def _clearance_variance(placed: list, polygon: list) -> float:
    """벽 이격의 분산. 헤드 수가 같으면 벽에서 고르게 떨어진 배치를 고른다(§8.1)."""
    if not placed:
        return 0.0
    gaps = [distance_to_boundary(p, polygon) for p in placed]
    mean = sum(gaps) / len(gaps)
    return sum((g - mean) ** 2 for g in gaps) / len(gaps)


# ────────────────────────────────────────────────────────────────────────────
# 살수장애 (§8.4, §8.5)
# ────────────────────────────────────────────────────────────────────────────

def _obstacle_point(obs) -> tuple[float, float] | None:
    if isinstance(obs, dict):
        if isinstance(obs.get("point"), (list, tuple)) and len(obs["point"]) >= 2:
            return (float(obs["point"][0]), float(obs["point"][1]))
        if obs.get("x") is not None and obs.get("y") is not None:
            return (float(obs["x"]), float(obs["y"]))
        if obs.get("polygon"):
            return representative_point(obs["polygon"])
    return None


def check_obstacles(layout: RoomLayout, room, constraints, obstacles) -> RoomLayout:
    """덕트·조명·배관은 60cm 룰, 보는 별도 표(§8.4).

    보를 60cm 룰에 넣지 마라 — 보는 0.75m 미만에서도 반사판을 보 하단보다 낮추는
    조건으로 허용된다. 즉 60cm 원 안에 보가 들어오는 것이 정상이고, 넣으면 보가
    있는 모든 실에서 대량 오탐이 나 배치가 무한 재실행된다.
    """
    if obstacles is None or obstacles.status is None or obstacles.status == "none":
        # §8.5 — 조용히 통과시키지 않는다. 검증하지 않은 것과 통과한 것은 다르다.
        layout.flags.append({
            "code": "OBSTACLE_UNVERIFIED", "room": room.id,
            "message": f"{room.id}: 장애물 정보 미확보 — 살수장애 검증을 수행하지 않았습니다.",
        })
        return layout

    polygon = [(float(p[0]), float(p[1])) for p in (room.polygon or [])]
    if len(polygon) < 3 or not layout.heads:
        return layout

    clearance = constraints.head_clearance_radius_m * _MM_PER_M
    wall_min = constraints.head_to_wall_clearance_m * _MM_PER_M
    blocking = [o for group in (obstacles.ducts, obstacles.lights)
                for o in group if _inside(o, polygon)]

    added = 0
    for obs in blocking:
        point = _obstacle_point(obs)
        near = [h for h in layout.heads
                if math.hypot(h.x - point[0], h.y - point[1]) < clearance]
        if not near:
            continue
        if distance_to_boundary(point, polygon) < wall_min:
            # 벽에 붙은 덕트 아래에는 헤드를 못 넣는다(2.7.7.2 벽 이격 10cm).
            # 여기서 조용히 넘어가면 살수장애가 남은 채 배치가 통과한다.
            layout.flags.append({
                "code": "OBSTRUCTION_UNRESOLVED", "room": room.id,
                "message": f"{room.id}: 살수장애물 아래에 헤드를 놓을 자리를 찾지 못했습니다"
                           " (NFTC 103 2.7.7.1).",
            })
            continue
        # 하부 헤드를 먼저 시도한다. 원래 헤드를 옮기면 그 자리의 피복이 깨진다.
        head = near[0]
        layout.heads.append(Head(
            id=f"{head.id}-U{added + 1}", room_id=room.id, x=point[0], y=point[1],
            row=head.row, col=head.col, provenance="under_obstacle"))
        added += 1

    if added:
        layout.metrics["heads_under_obstacles"] = added
        layout.metrics["head_count"] = len(layout.heads)

    beams = [o for o in obstacles.beams if _inside(o, polygon)]
    for obs in beams:
        point = _obstacle_point(obs)
        if point is None:
            continue
        for head in layout.heads:
            horizontal_m = math.hypot(head.x - point[0], head.y - point[1]) / _MM_PER_M
            requirement = beam_clearance(horizontal_m)
            layout.beam_requirements.append({
                "head_id": head.id, "horizontal_m": round(horizontal_m, 3),
                "requirement": requirement,
            })
    if layout.beam_requirements:
        # 반사판 표고는 C5 가 정한다. 여기서 통과/실패를 적으면 재 본 적 없는
        # 수직거리를 판정한 것이 된다.
        layout.metrics["beam_requirements"] = len(layout.beam_requirements)
    return layout


def _inside(obs, polygon: list) -> bool:
    point = _obstacle_point(obs)
    return bool(point and point_in_polygon(point, polygon))
