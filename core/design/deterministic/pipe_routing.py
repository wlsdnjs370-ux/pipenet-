# -*- coding: utf-8 -*-
"""지시서 §9.2 C510·C520·C530 + §9.3 — 가지배관 골격과 교차배관 경로.

여기서 정하는 것은 넷이다. 교차배관이 어느 방향으로 눕는가(C510), 헤드가 어느
가지배관에 실리고 분기점이 어디에 서는가(C520), 그 분기점들을 실제로 어떤 경로로
잇는가(C530), 그리고 그 결과가 토너먼트가 아닌가(§9.3).

8개 상한은 **분기점 기준 한쪽**이다(§9.4). 가지배관 전체로는 16개가 적법하므로
전체를 세어 자르면 배관이 두 배로 늘고, 늘어난 만큼 전부 자재비다.

좌표는 mm 다 — C4 가 mm 로 놓았고, 여기서 m 로 바꾸면 왕복하는 동안 반올림이
쌓인다. m 로 바꾸는 것은 방출(C560) 한 곳에서만 한다.
"""
from __future__ import annotations

import heapq
import itertools
import math
from dataclasses import dataclass, field
from typing import Any

from ..recognize.spatial import (
    NodeIndex, SegmentGrid, point_in_polygon, segments_cross)
from .zoning import JOIN_TOL_MM

_MM_PER_M = 1000.0

# 좌표 비교 공차. mm 단위 격자 위의 점이라 반올림 오차만 흡수하면 된다.
_EPS_MM = 1e-6

# §9.2 — OBB 장단변비가 이 값을 넘으면 복도형으로 보고 축 판정을 건너뛴다.
# 복도에서 가지배관이 장변을 따라가면 실 안으로 못 들어간다.
CORRIDOR_ASPECT = 4.0

# §9.2 C520 — 같은 가지배관으로 묶는 축 좌표 편차 상한 = 헤드 간격 / 4.
_LINE_TOL_DIVISOR = 4.0


@dataclass
class ZoneAxes:
    """구역의 배관 축. `theta` 는 **교차배관이 눕는 방향**(rad)이다.

    가지배관은 여기에 수직으로 뻗는다. 투영 좌표 `(a, b)` 에서 a 는 교차배관을
    따라가는 거리, b 는 교차배관에서 떨어진 거리다.
    """

    theta: float
    cross_span_mm: float
    branch_span_mm: float
    corridor: bool = False
    flipped: bool = False

    def project(self, x: float, y: float) -> tuple[float, float]:
        cos_t, sin_t = math.cos(self.theta), math.sin(self.theta)
        return (x * cos_t + y * sin_t, -x * sin_t + y * cos_t)

    def world(self, a: float, b: float) -> tuple[float, float]:
        cos_t, sin_t = math.cos(self.theta), math.sin(self.theta)
        return (a * cos_t - b * sin_t, a * sin_t + b * cos_t)

    def to_dict(self) -> dict[str, Any]:
        return {
            "cross_axis_deg": round(math.degrees(self.theta) % 180.0, 2),
            "cross_span_m": round(self.cross_span_mm / _MM_PER_M, 3),
            "branch_span_m": round(self.branch_span_mm / _MM_PER_M, 3),
            "corridor": self.corridor, "flipped": self.flipped,
        }


@dataclass
class CrossMain:
    """교차배관 하나. `b_mm` 는 이 배관이 놓인 선(가지배관 방향 좌표)이다.

    `path` 와 `spurs` 는 C530 이 채우는 세계 좌표 폴리라인이다. C530 을 돌리기
    전에는 비어 있다 — 빈 것과 "직선이라 꺾인 점이 없다" 는 다르므로, 경로가 있는
    교차배관은 항상 끝점 둘을 갖는다.
    """

    id: str
    b_mm: float
    a_span: tuple[float, float]
    min_dn: int = 0
    path: list = field(default_factory=list)
    spurs: list = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {"id": self.id, "b_mm": round(self.b_mm, 1),
                "a_span_mm": [round(self.a_span[0], 1), round(self.a_span[1], 1)],
                "min_dn": self.min_dn, "path": self.path, "spurs": self.spurs}


@dataclass
class Branch:
    """가지배관 하나. `left`/`right` 는 분기점 기준 양쪽이다 — 상한은 각각에 건다."""

    id: str
    cross_id: str
    a_mm: float
    tee: tuple[float, float]
    left: list[str] = field(default_factory=list)
    right: list[str] = field(default_factory=list)

    @property
    def heads(self) -> list[str]:
        return self.left + self.right

    @property
    def per_side_max(self) -> int:
        return max(len(self.left), len(self.right))

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id, "cross_id": self.cross_id,
            "tee": [round(self.tee[0], 1), round(self.tee[1], 1)],
            "left": list(self.left), "right": list(self.right),
            "heads_per_side_max": self.per_side_max,
        }


@dataclass
class ZonePlan:
    """`head_ab` 는 헤드 id → 구역 축 좌표. C530 이후 단계가 전부 이 프레임에서 돈다.

    세계 좌표에서 다시 투영하면 같은 각도를 두 번 계산하게 되고, 축이 바뀌었을 때
    한쪽만 따라간다.
    """

    zone_id: str
    axes: ZoneAxes
    crosses: list[CrossMain] = field(default_factory=list)
    branches: list[Branch] = field(default_factory=list)
    head_ab: dict = field(default_factory=dict)
    flags: list = field(default_factory=list)
    metrics: dict = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {
            "zone_id": self.zone_id, "axes": self.axes.to_dict(),
            "cross_mains": [c.to_dict() for c in self.crosses],
            "branches": [b.to_dict() for b in self.branches],
            "flags": self.flags, "metrics": self.metrics,
        }


# ────────────────────────────────────────────────────────────────────────────
# C510 — 축 결정
# ────────────────────────────────────────────────────────────────────────────

def zone_axes(rooms: list, valve_point=None) -> ZoneAxes:
    """C510 — 교차배관 방향을 정한다.

    [문서정합 §9.2] 명세는 "구역 헤드의 OBB 장변" 을 교차배관 축으로 쓰라고 적었다.
    그런데 헤드는 실마다 **제 실의 격자** 위에 놓여 있고(C4 `metrics.axis_deg`),
    그 격자는 실의 벽 방향을 따른다. 여러 실을 한데 묶은 점 뭉치의 OBB 각도를 새로
    재면 어느 실의 헤드 열과도 어긋난 각도가 나오고, 그 각도로 그은 가지배관은
    지나가야 할 헤드를 하나도 지나가지 못한다. 그래서 각도 후보는 실 격자 각도로
    한정하고(헤드 수가 가장 많은 격자를 고른다), OBB 는 **장단변 판정에만** 쓴다.
    """
    weight: dict[float, int] = {}
    for room in rooms:
        heads = room.get("heads") or []
        if not heads:
            continue
        deg = float((room.get("metrics") or {}).get("axis_deg") or 0.0) % 90.0
        weight[deg] = weight.get(deg, 0) + len(heads)
    if not weight:
        raise ValueError("헤드가 없는 구역에는 배관을 놓을 수 없습니다.")

    theta = math.radians(min(weight, key=lambda d: (-weight[d], d)))
    points = [(float(h["x"]), float(h["y"]))
              for room in rooms for h in (room.get("heads") or [])]
    cos_t, sin_t = math.cos(theta), math.sin(theta)
    us = [x * cos_t + y * sin_t for x, y in points]
    vs = [-x * sin_t + y * cos_t for x, y in points]
    span_u, span_v = max(us) - min(us), max(vs) - min(vs)

    # 장변을 교차배관으로. 가지배관은 짧고 많아야 한쪽 8개 제한에 여유가 생긴다.
    long_is_u = span_u >= span_v
    cross_theta = theta if long_is_u else theta + math.pi / 2
    span_long, span_short = ((span_u, span_v) if long_is_u else (span_v, span_u))
    corridor = span_short <= 0 or span_long / span_short >= CORRIDOR_ASPECT

    flipped = False
    if not corridor and valve_point is not None:
        # [문서정합 §9.2] 명세는 "밸브가 단변 쪽에 있으면 뒤집는다" 고만 적어 어느
        # 변을 단변 쪽으로 볼지가 갈린다. 뒤집는 이유("교차배관은 밸브에서 곧게
        # 뻗어야 주배관 우회가 없다")로 되돌려 판정한다 — 밸브에서 구역 중심을
        # 향하는 방향이 교차배관 방향과 어긋나 있으면 뒤집는다.
        vx, vy = float(valve_point[0]), float(valve_point[1])
        reach_u = abs((max(us) + min(us)) / 2.0 - (vx * cos_t + vy * sin_t))
        reach_v = abs((max(vs) + min(vs)) / 2.0 - (-vx * sin_t + vy * cos_t))
        along, across = (reach_u, reach_v) if long_is_u else (reach_v, reach_u)
        flipped = across > along
        if flipped:
            cross_theta += math.pi / 2
            span_long, span_short = span_short, span_long

    return ZoneAxes(cross_theta % math.pi, span_long, span_short,
                    corridor=corridor, flipped=flipped)


# ────────────────────────────────────────────────────────────────────────────
# C520 — 가지배관 생성
# ────────────────────────────────────────────────────────────────────────────

def _cluster(values: list[float], tol: float) -> list[list[int]]:
    """축 좌표가 tol 안에 든 것끼리 묶는다. 반환은 원본 인덱스 묶음."""
    order = sorted(range(len(values)), key=lambda i: values[i])
    groups: list[list[int]] = []
    for i in order:
        if groups and values[i] - values[groups[-1][0]] <= tol:
            groups[-1].append(i)
        else:
            groups.append([i])
    return groups


def _segment_cost(prefix: list[list[int]], lo: int, hi: int) -> tuple[int, int]:
    """레벨 [lo, hi) 를 교차배관 하나가 먹을 때 (한쪽 최대 헤드 수, 분할점).

    분할점 p 는 레벨 p-1 과 p 사이를 뜻한다. p == lo 면 교차배관이 구간 바깥에
    서고 헤드가 전부 한쪽에 실린다 — 벽을 따라 도는 교차배관이 그 모양이다.
    """
    live = [row for row in prefix if row[hi] > row[lo]]
    best: tuple[int, int, int] | None = None
    for p in range(lo, hi + 1):
        worst = 0
        for row in live:
            worst = max(worst, row[p] - row[lo], row[hi] - row[p])
        # 같은 값이면 구간 가운데에 가까운 분할점을 쓴다. 교차배관이 한쪽 끝에
        # 붙으면 가지배관이 한 방향으로만 길어져 다음 구간에서 또 걸린다.
        offset = abs((p - lo) - (hi - p))
        if best is None or (worst, offset) < (best[0], best[1]):
            best = (worst, offset, p)
    return best[0], best[2]


def _split_levels(prefix: list[list[int]], levels: int,
                  per_side_max: int) -> tuple[int, list[tuple[int, int, int]]]:
    """교차배관 수를 1 부터 늘리며 한쪽 상한을 맞춘다 (§9.2 분할 전략 1·2).

    전략 1(분기점 이동)은 교차배관 1개일 때 분할점 p 를 고르는 것이고, 전략
    2(교차배관 추가)는 구간을 늘리는 것이다. 상한을 맞추는 **최소 개수**를 찾으므로
    맞출 수 있으면 교차배관을 더 놓지 않는다.

    전략 3(가지배관 분리)은 여기서 쓸 일이 없다 — 구간을 레벨 하나까지 쪼갤 수
    있는 한 전략 2 가 항상 먼저 성립한다. 라우팅이 실제로 교차배관을 못 놓는
    경우는 C530 의 재배정 소관이다.
    """
    cost: dict[tuple[int, int], tuple[int, int]] = {}
    reach = {0: (0, [])}
    best: tuple[int, list] | None = None
    for _k in range(levels):
        nxt: dict[int, tuple[int, list]] = {}
        for end, (worst, segs) in reach.items():
            for hi in range(end + 1, levels + 1):
                if (end, hi) not in cost:
                    cost[(end, hi)] = _segment_cost(prefix, end, hi)
                w, p = cost[(end, hi)]
                cand = (max(worst, w), segs + [(end, hi, p)])
                if hi not in nxt or cand[0] < nxt[hi][0]:
                    nxt[hi] = cand
        reach = nxt
        done = reach.get(levels)
        if done and (best is None or done[0] < best[0]):
            best = done
        if best and best[0] <= per_side_max:
            break
    return best if best else (0, [])


def _split_sides(head_ab: dict, ids, b_line: float) -> tuple[list, list]:
    """분기점 기준 좌우. 상한은 합계가 아니라 각각에 건다(§9.4).

    분기점 위에 정확히 놓인 헤드는 한쪽으로 몰아 넣는다. 어느 쪽에도 안 넣으면
    그 헤드는 가지배관에서 조용히 사라지고, 사라진 헤드는 미방호 구역이 된다.
    """
    return ([h for h in ids if head_ab[h][1] < b_line],
            [h for h in ids if head_ab[h][1] >= b_line])


def plan_branches(zone_id: str, rooms: list, valve_point, limits) -> ZonePlan:
    """C520 — 헤드를 가지배관에 싣고 분기점을 세운다.

    `limits` 는 `branch_heads_per_side_max` 와 `head_spacing_square_m` 만 본다.
    """
    axes = zone_axes(rooms, valve_point)
    heads = [(h["id"], *axes.project(float(h["x"]), float(h["y"])))
             for room in rooms for h in (room.get("heads") or [])]
    plan = ZonePlan(zone_id=zone_id, axes=axes,
                    head_ab={h[0]: (h[1], h[2]) for h in heads})
    if not heads:
        return plan

    tol = float(limits.head_spacing_square_m) * _MM_PER_M / _LINE_TOL_DIVISOR
    lines = _cluster([h[1] for h in heads], tol)      # 가지배관 후보 (a 가 같다)
    levels = _cluster([h[2] for h in heads], tol)     # 교차배관 후보 선 (b)
    level_of = {i: n for n, group in enumerate(levels) for i in group}
    level_b = [sum(heads[i][2] for i in g) / len(g) for g in levels]

    # prefix[line][n] = 그 가지배관이 레벨 n 미만에 가진 헤드 수. 한쪽 헤드 수를
    # 분할점마다 다시 세지 않으려고 미리 쌓아 둔다.
    prefix = []
    for group in lines:
        row = [0] * (len(levels) + 1)
        for i in group:
            row[level_of[i] + 1] += 1
        for n in range(1, len(row)):
            row[n] += row[n - 1]
        prefix.append(row)

    per_side_max = int(limits.branch_heads_per_side_max)
    worst, segments = _split_levels(prefix, len(levels), per_side_max)

    for s, (lo, hi, p) in enumerate(segments, start=1):
        if p == lo:
            b_line = level_b[lo] - tol
        elif p == hi:
            b_line = level_b[hi - 1] + tol
        else:
            b_line = (level_b[p - 1] + level_b[p]) / 2.0
        members = [i for i in range(len(heads)) if lo <= level_of[i] < hi]
        cross = CrossMain(id=f"CM-{zone_id}-{s:02d}", b_mm=b_line,
                          a_span=(min(heads[i][1] for i in members),
                                  max(heads[i][1] for i in members)))
        plan.crosses.append(cross)

        for n, group in enumerate(lines, start=1):
            mine = sorted((i for i in group if lo <= level_of[i] < hi),
                          key=lambda i: heads[i][2])
            if not mine:
                continue
            a_mm = sum(heads[i][1] for i in mine) / len(mine)
            left, right = _split_sides(plan.head_ab, [heads[i][0] for i in mine],
                                       b_line)
            plan.branches.append(Branch(
                id=f"BR-{zone_id}-{s:02d}-{n:02d}", cross_id=cross.id, a_mm=a_mm,
                tee=axes.world(a_mm, b_line), left=left, right=right))

    plan.metrics = {
        "heads": len(heads), "cross_mains": len(plan.crosses),
        "branches": len(plan.branches), "heads_per_side_max": worst,
        "heads_per_side_limit": per_side_max, **axes.to_dict(),
    }
    if worst > per_side_max:
        plan.flags.append({
            "code": "BRANCH_SPLIT_FAILED", "zone_id": zone_id,
            "message": f"분기점 한쪽 헤드가 {worst}개로 상한 {per_side_max}개를 "
                       "넘습니다. 교차배관을 더 놓아도 줄지 않아 구역을 다시 "
                       "나눠야 합니다.",
        })
    return plan


# ────────────────────────────────────────────────────────────────────────────
# C530 — 교차배관 경로 (§9.2)
# ────────────────────────────────────────────────────────────────────────────

GRID_MM = 500.0
# 좁은 복도가 있으면 격자를 낮춘다. 더 낮추면 탐색 비용이 급증한다(§9.2).
GRID_MM_NARROW = 250.0
NARROW_WIDTH_MM = 2000.0
NARROW_ASPECT = 2.0

# 방향 1회 전환 = 3m 우회와 등가. 크게 잡아야 실무 도면처럼 곧은 경로가 나온다.
TURN_PENALTY_FACTOR = 6.0
# 배관은 슬리브로 벽을 관통한다. 벽을 절대 장애물로 두면 경로가 아예 안 나온다.
WALL_PIERCE_FACTOR = 20.0
ASTAR_NODE_LIMIT = 200_000

_STEPS = ((1, 0), (-1, 0), (0, 1), (0, -1))


def wall_segments(rooms: list, doors: list) -> list:
    """실 경계에서 문을 뺀 것이 벽이다.

    [문서정합 §9.2] 명세는 장애물을 "벽 중심선 buffer(thickness/2 + 50mm)" 로
    잡으라고 적었지만 building.json(§10.1)에는 벽이 없다 — C150 이 낸 중심선은
    C170 이 실 폴리곤을 뽑으면서 소비된다. 남는 것은 실 경계와 가상 간선(문)뿐이라
    경계 중 문이 아닌 변을 벽으로 본다. 두께가 없으므로 buffer 대신 관통으로 센다.
    """
    index = NodeIndex(JOIN_TOL_MM)
    edges: dict[tuple[int, int], tuple] = {}
    for room in rooms:
        poly = room.get("polygon") or []
        if len(poly) < 3:
            continue
        nodes = [index.add(float(p[0]), float(p[1])) for p in poly]
        for i, a in enumerate(nodes):
            b = nodes[(i + 1) % len(nodes)]
            if a != b:
                edges.setdefault((a, b) if a < b else (b, a),
                                 (index.points[a], index.points[b]))

    for door in doors:
        p1, p2 = door.get("p1"), door.get("p2")
        if not p1 or not p2:
            continue
        n1 = index.find(float(p1[0]), float(p1[1]))
        n2 = index.find(float(p2[0]), float(p2[1]))
        if n1 is not None and n2 is not None and n1 != n2:
            edges.pop((n1, n2) if n1 < n2 else (n2, n1), None)

    return [(p[0], p[1], q[0], q[1]) for p, q in edges.values()]


def _auto_grid_mm(rooms: list) -> float:
    """§9.2 — 폭 2m 미만의 복도가 있으면 250mm.

    폭만 보면 안 된다. 1.5m 짜리 창고까지 걸려 건물 전체의 격자가 내려가고, 그
    비용은 A* 탐색이 네 배로 문다. 길쭉한 실(장단변비 2 이상)만 복도로 본다.
    """
    for room in rooms:
        poly = room.get("polygon") or []
        if len(poly) < 3:
            continue
        xs = [float(p[0]) for p in poly]
        ys = [float(p[1]) for p in poly]
        short = min(max(xs) - min(xs), max(ys) - min(ys))
        long_ = max(max(xs) - min(xs), max(ys) - min(ys))
        if short < NARROW_WIDTH_MM and long_ >= short * NARROW_ASPECT:
            return GRID_MM_NARROW
    return GRID_MM


def _hits(grid: SegmentGrid, segments: list, p, q) -> int:
    if not segments:
        return 0
    seg = (p[0], p[1], q[0], q[1])
    return sum(1 for i in grid.lookup(grid.walk(seg))
               if segments_cross(seg, segments[i]))


def _simplify(points: list) -> list:
    """같은 점과 일직선 위의 중간 점을 지운다. 남는 꼭짓점이 곧 방향 전환이다."""
    out: list = []
    for p in points:
        if not out or abs(out[-1][0] - p[0]) > _EPS_MM or abs(out[-1][1] - p[1]) > _EPS_MM:
            out.append(p)
    if len(out) < 3:
        return out
    keep = [out[0]]
    for i in range(1, len(out) - 1):
        (ax, ay), (bx, by), (cx, cy) = keep[-1], out[i], out[i + 1]
        if abs((bx - ax) * (cy - by) - (by - ay) * (cx - bx)) > _EPS_MM:
            keep.append(out[i])
    keep.append(out[-1])
    return keep


def _axis_parallel(p, q) -> bool:
    return abs(p[0] - q[0]) < _EPS_MM or abs(p[1] - q[1]) < _EPS_MM


def _nearest_on(path: list, point) -> tuple[float, float]:
    """폴리라인 위에서 `point` 에 가장 가까운 점. 꼭짓점으로만 스냅하지 않는다."""
    best, best_d = path[0], None
    for i in range(len(path) - 1):
        (ax, ay), (bx, by) = path[i], path[i + 1]
        dx, dy = bx - ax, by - ay
        span = dx * dx + dy * dy
        t = 0.0 if span == 0 else max(0.0, min(
            1.0, ((point[0] - ax) * dx + (point[1] - ay) * dy) / span))
        cand = (ax + dx * t, ay + dy * t)
        d = math.hypot(cand[0] - point[0], cand[1] - point[1])
        if best_d is None or d < best_d:
            best, best_d = cand, d
    return best


class RouteField:
    """구역 축 프레임 위의 직교 격자. 벽은 비싸고 코어는 지날 수 없다(§9.2).

    [문서정합 §9.2] 명세는 "직교 격자 A*" 라고만 적었다. 세계 좌표 x·y 에 격자를
    깔면 축이 기운 건물에서 교차배관이 계단처럼 꺾인다. 직교의 기준은 구역 축이므로
    격자를 C510 이 정한 `(a, b)` 프레임 위에 깐다.

    `rooms` 는 그 층 전체를 받는다 — 구역 밖의 벽도 벽이고, 옆 구역의 코어도
    통과 불가다.
    """

    def __init__(self, axes: ZoneAxes, rooms: list, *,
                 cores=(), doors=(), grid_mm=None):
        self.axes = axes
        self.grid_mm = float(grid_mm) if grid_mm else _auto_grid_mm(rooms)
        project = axes.project

        self.walls = [(*project(s[0], s[1]), *project(s[2], s[3]))
                      for s in wall_segments(rooms, doors)]
        self._wall_grid = SegmentGrid(self.grid_mm)
        for i, seg in enumerate(self.walls):
            self._wall_grid.add(i, self._wall_grid.walk(seg))

        polygons = [c.get("polygon") or [] for c in cores]
        polygons += [r.get("polygon") or [] for r in rooms if r.get("head_exempt")]
        self.blocked = [[project(float(p[0]), float(p[1])) for p in poly]
                        for poly in polygons if len(poly) >= 3]
        self._block_edges = [(p[0], p[1], q[0], q[1]) for poly in self.blocked
                             for p, q in zip(poly, poly[1:] + poly[:1])]
        self._block_grid = SegmentGrid(self.grid_mm)
        for i, seg in enumerate(self._block_edges):
            self._block_grid.add(i, self._block_grid.walk(seg))

        points = [project(float(p[0]), float(p[1]))
                  for r in rooms for p in (r.get("polygon") or [])]
        pad = self.grid_mm * 4.0
        self.lo = self.hi = None
        if points:
            self.lo = (min(a for a, _ in points) - pad, min(b for _, b in points) - pad)
            self.hi = (max(a for a, _ in points) + pad, max(b for _, b in points) + pad)

    def inside(self, a: float, b: float) -> bool:
        """탐색 범위. 없으면 열린 평면에서 20만 노드를 다 쓰고서야 실패한다."""
        if self.lo is None:
            return True
        return (self.lo[0] <= a <= self.hi[0]) and (self.lo[1] <= b <= self.hi[1])

    def blocked_at(self, a: float, b: float) -> bool:
        return any(point_in_polygon((a, b), poly) for poly in self.blocked)

    def pierce_count(self, p, q) -> int:
        """선분이 관통하는 벽의 수. 비용이자 산출물이다 — 슬리브가 그만큼 든다."""
        return _hits(self._wall_grid, self.walls, p, q)

    def clear(self, p, q) -> bool:
        """절대 장애물을 지나지 않는가. 벽은 여기서 막지 않는다 — 비용으로 문다."""
        if self.blocked_at(*p) or self.blocked_at(*q):
            return False
        return _hits(self._block_grid, self._block_edges, p, q) == 0

    def route(self, start, goal, *, grid_mm=None, node_limit=ASTAR_NODE_LIMIT):
        """4-이웃 A*. 상태는 `(셀, 들어온 방향)` 이라야 전환 비용을 물 수 있다.

        격자 원점을 출발점에 맞춘다. 맞추지 않으면 첫 걸음이 셀 중심으로 비스듬히
        꺾여 축 평행 경로가 아니게 된다.
        """
        g = float(grid_mm or self.grid_mm)
        sa, sb = float(start[0]), float(start[1])
        turn, pierce = g * TURN_PENALTY_FACTOR, g * WALL_PIERCE_FACTOR
        target = (round((float(goal[0]) - sa) / g), round((float(goal[1]) - sb) / g))

        def point(cell):
            return (sa + cell[0] * g, sb + cell[1] * g)

        if self.blocked_at(sa, sb) or self.blocked_at(*point(target)):
            return None, "unreachable"

        tick = itertools.count()
        start_state = ((0, 0), None)
        heap = [(0.0, next(tick), 0.0, start_state)]
        came: dict = {start_state: None}
        best: dict = {start_state: 0.0}
        expanded = 0

        while heap:
            _f, _n, cost, state = heapq.heappop(heap)
            if cost > best.get(state, math.inf):
                continue
            cell, direction = state
            if cell == target:
                # 마지막 모서리를 못 넣으면 경로가 없는 것이다. 사유를 비워서
                # 돌려주면 호출부가 성공으로 읽는다.
                traced = self._trace(came, state, point, goal)
                return (traced, None) if traced else (None, "unreachable")
            expanded += 1
            if expanded > node_limit:
                return None, "timeout"

            here = point(cell)
            for d, (di, dj) in enumerate(_STEPS):
                nxt = (cell[0] + di, cell[1] + dj)
                there = point(nxt)
                if not self.inside(*there) or not self.clear(here, there):
                    continue
                step = g + (turn if direction is not None and d != direction else 0.0)
                step += pierce * self.pierce_count(here, there)
                nstate = (nxt, d)
                if cost + step < best.get(nstate, math.inf):
                    best[nstate] = cost + step
                    came[nstate] = state
                    heuristic = (abs(nxt[0] - target[0]) + abs(nxt[1] - target[1])) * g
                    heapq.heappush(
                        heap, (cost + step + heuristic, next(tick), cost + step, nstate))
        return None, "unreachable"

    def _trace(self, came: dict, state, point, goal):
        cells = []
        while state is not None:
            cells.append(state[0])
            state = came[state]
        path = _simplify([point(c) for c in reversed(cells)])

        # 마지막 셀 중심과 실제 목표점은 격자의 절반 안에서 어긋나 있다. 비스듬히
        # 이으면 직교 라우팅이 아니게 되므로 축 평행 모서리를 하나 넣는다.
        last, exact = path[-1], (float(goal[0]), float(goal[1]))
        if _axis_parallel(last, exact) and self.clear(last, exact):
            return _simplify(path + [exact])
        for corner in ((exact[0], last[1]), (last[0], exact[1])):
            if self.clear(last, corner) and self.clear(corner, exact):
                return _simplify(path + [corner, exact])
        return None

    def connect(self, start, goal):
        """곧게 갈 수 있으면 곧게 간다. 반환 `(폴리라인, 사유)` — 사유 None 이 성공.

        §9.2 는 탐색 노드 20만을 넘으면 격자를 2배로 올려 1회 재시도하라고 한다.
        """
        start = (float(start[0]), float(start[1]))
        goal = (float(goal[0]), float(goal[1]))
        if _axis_parallel(start, goal) and self.clear(start, goal):
            return [start, goal], None
        path, reason = self.route(start, goal)
        if path is None and reason == "timeout":
            path, reason = self.route(start, goal, grid_mm=self.grid_mm * 2.0)
        return path, reason


def route_cross_mains(plan: ZonePlan, field: RouteField, *,
                      min_dn: int = 0) -> ZonePlan:
    """C530 — 분기점들을 잇는다. 못 이으면 인접 교차배관으로 재배정한다.

    도달 못한 분기점을 빼고 그래프를 완성하면 수리계산은 통과하는데 실제로는
    미방호 구역이 생긴다(§9.2). 그래서 빼지 않고 플래그를 남긴다.
    """
    axes = plan.axes
    paths: dict[str, list] = {}
    spurs: list[list] = []
    detached: list[tuple[CrossMain, Branch, str]] = []

    for cross in plan.crosses:
        cross.min_dn = int(min_dn)
        mine = sorted((b for b in plan.branches if b.cross_id == cross.id),
                      key=lambda b: b.a_mm)
        if not mine:
            continue
        anchor = mine[0]
        path = [(anchor.a_mm, cross.b_mm)]
        for branch in mine[1:]:
            leg, reason = field.connect((anchor.a_mm, cross.b_mm),
                                        (branch.a_mm, cross.b_mm))
            if leg is None:
                detached.append((cross, branch, reason))
                continue
            path.extend(leg[1:])
            anchor = branch
        paths[cross.id] = _simplify(path)

    per_side_max = int(plan.metrics.get("heads_per_side_limit") or 0)
    reassigned = 0
    for cross, branch, reason in detached:
        # 재배정은 분기점을 인접 교차배관 위로 옮기는 것이다. 자리를 그대로 두고
        # 먹이는 배관만 바꾸면 아무것도 달라지지 않는다 — 그 자리에 닿느냐는 어느
        # 교차배관에서 출발하든 같은 물음이라, 도달 못한 분기점은 여전히 도달 못한다.
        moved = False
        for other in sorted((c for c in plan.crosses if c.id != cross.id
                             and len(paths.get(c.id) or ()) >= 2),
                            key=lambda c: abs(c.b_mm - cross.b_mm)):
            tee = (branch.a_mm, other.b_mm)
            left, right = _split_sides(plan.head_ab, branch.heads, other.b_mm)
            # 옮긴 자리에서 §9.4 한쪽 8개가 깨지면 그것은 해결이 아니다.
            if per_side_max and max(len(left), len(right)) > per_side_max:
                continue
            spur, why = field.connect(_nearest_on(paths[other.id], tee), tee)
            reason = why or reason
            if spur is None:
                continue
            branch.cross_id, branch.left, branch.right = other.id, left, right
            branch.tee = axes.world(*tee)
            other.spurs.append([list(axes.world(a, b)) for a, b in spur])
            spurs.append(spur)
            reassigned += 1
            moved = True
            break
        if not moved:
            plan.flags.append({
                "code": "ROUTING_TIMEOUT" if reason == "timeout" else "ROUTING_UNREACHABLE",
                "zone_id": plan.zone_id, "cross_id": cross.id,
                "branch_id": branch.id, "heads": list(branch.heads),
                "message": f"{branch.id} 분기점까지 배관을 놓을 경로가 없습니다. "
                           f"헤드 {len(branch.heads)}개가 미방호입니다.",
            })

    for cross in plan.crosses:
        cross.path = [list(axes.world(a, b)) for a, b in (paths.get(cross.id) or ())]

    length = 0.0
    turns = pierces = 0
    for line in list(paths.values()) + spurs:
        turns += max(0, len(line) - 2)
        for p, q in zip(line, line[1:]):
            length += math.hypot(q[0] - p[0], q[1] - p[1])
            pierces += field.pierce_count(p, q)

    plan.metrics.update({
        "grid_mm": field.grid_mm,
        "cross_length_m": round(length / _MM_PER_M, 2),
        "cross_turns": int(turns),
        "wall_pierces": int(pierces),
        "branches_reassigned": reassigned,
    })
    return plan


# ────────────────────────────────────────────────────────────────────────────
# §9.3 — 토너먼트 금지
# ────────────────────────────────────────────────────────────────────────────

def check_tournament(adjacency: dict, source) -> list[str]:
    """토너먼트(대칭 이분) 형상을 찾아 위반 목록을 낸다. 빈 목록이 통과다.

    [문서정합 §9.3] 명세의 예시 코드는 "급수원→헤드 경로에 차수 3 이상 분기가
    연속 2회" 를 위반으로 삼는다. 그대로 쓰면 **빗살이 전부 걸린다** — 교차배관
    위의 분기점들은 서로 이웃한 차수 3 노드라서 두 번째 분기점에서 바로 연속 2회가
    된다. 금지의 실제 취지는 NFTC 2.5.10 의 "가지배관 배열", 즉 **갈라져 나온
    가지가 또 갈라지지 않는다**이므로 그렇게 판정한다: 어떤 분기점의 하류 갈래
    둘 이상이 각각 또 분기를 품으면 위반이다.

    급수원 자신은 뺀다. 급수원에서 교차배관 여럿으로 갈라지는 것은 주배관 배열이지
    가지배관 배열이 아니다.
    """
    if source not in adjacency:
        return []

    parent = {source: None}
    order = [source]
    for node in order:
        for nxt in adjacency.get(node, ()):
            if nxt not in parent:
                parent[nxt] = node
                order.append(nxt)

    # 하류부터 되짚어 "이 노드 아래에 분기가 있는가" 를 쌓는다.
    children: dict[Any, list] = {n: [] for n in order}
    for node in order[1:]:
        children[parent[node]].append(node)
    has_split: dict[Any, bool] = {}
    for node in reversed(order):
        has_split[node] = (len(children[node]) >= 2
                           or any(has_split[c] for c in children[node]))

    violations = []
    for node in order:
        if node is source or len(children[node]) < 2:
            continue
        splitting = [c for c in children[node] if has_split[c]]
        if len(splitting) >= 2:
            violations.append(
                f"{node}: 갈라진 가지 {len(splitting)}개가 또 갈라집니다 — "
                "토너먼트 방식입니다.")
    return violations
