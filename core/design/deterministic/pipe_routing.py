# -*- coding: utf-8 -*-
"""지시서 §9.2 C510·C520 + §9.3 — 가지배관 골격.

여기서 정하는 것은 셋이다. 교차배관이 어느 방향으로 눕는가(C510), 헤드가 어느
가지배관에 실리고 분기점이 어디에 서는가(C520), 그리고 그 결과가 토너먼트가
아닌가(§9.3).

8개 상한은 **분기점 기준 한쪽**이다(§9.4). 가지배관 전체로는 16개가 적법하므로
전체를 세어 자르면 배관이 두 배로 늘고, 늘어난 만큼 전부 자재비다.

좌표는 mm 다 — C4 가 mm 로 놓았고, 여기서 m 로 바꾸면 왕복하는 동안 반올림이
쌓인다. m 로 바꾸는 것은 방출(C560) 한 곳에서만 한다.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from typing import Any

_MM_PER_M = 1000.0

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
    """교차배관 하나. `b_mm` 는 이 배관이 놓인 선(가지배관 방향 좌표)이다."""

    id: str
    b_mm: float
    a_span: tuple[float, float]

    def to_dict(self) -> dict[str, Any]:
        return {"id": self.id, "b_mm": round(self.b_mm, 1),
                "a_span_mm": [round(self.a_span[0], 1), round(self.a_span[1], 1)]}


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
    zone_id: str
    axes: ZoneAxes
    crosses: list[CrossMain] = field(default_factory=list)
    branches: list[Branch] = field(default_factory=list)
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


def plan_branches(zone_id: str, rooms: list, valve_point, limits) -> ZonePlan:
    """C520 — 헤드를 가지배관에 싣고 분기점을 세운다.

    `limits` 는 `branch_heads_per_side_max` 와 `head_spacing_square_m` 만 본다.
    """
    axes = zone_axes(rooms, valve_point)
    heads = [(h["id"], *axes.project(float(h["x"]), float(h["y"])))
             for room in rooms for h in (room.get("heads") or [])]
    plan = ZonePlan(zone_id=zone_id, axes=axes)
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
            plan.branches.append(Branch(
                id=f"BR-{zone_id}-{s:02d}-{n:02d}", cross_id=cross.id, a_mm=a_mm,
                tee=axes.world(a_mm, b_line),
                left=[heads[i][0] for i in mine if heads[i][2] < b_line],
                right=[heads[i][0] for i in mine if heads[i][2] > b_line]))

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
