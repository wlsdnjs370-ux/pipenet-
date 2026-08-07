# -*- coding: utf-8 -*-
"""부속(fitting) 계상 규칙 — 두 브랜치가 함께 쓰는 순수 함수.

현업 관행(2026-08-05 실무자 강의)에서 온 두 가지가 여기 들어 있다.

  1. 티는 흐름 방향이 바뀌면 분류티, 직진하면 직류티다. **직류티는 수리계산에
     입력하지 않는다** — 마찰손실 비가 대략 10:1(구경에 따라 5:1)이라
     직류티까지 분류티로 세면 손실이 통째로 부풀려진다.
  2. PIPENET 이 제공하는 엘보는 45°/90°(및 롱레디우스)뿐이다. 그 사이 각도는
     둘 중 가까운 쪽으로 보내되, 어느 쪽으로도 보낼 수 없으면 지어내지 말고
     "판정 불가" 로 센다.

이 모듈은 좌표와 각도만 다룬다. 테이블 구조도, DXF 도, 솔버도 모른다 —
`remote30_constants.py` 와 같은 위치에 두고 양 브랜치가 import 한다.

주의: 부속 등가길이를 내부 솔버 해석에 넣는 것은 여기 범위가 아니다
(`BLOCKED.md §7 질문2` — 미모델링 + safety=1.1 상쇄가 별도 과제).
"""
from __future__ import annotations

import math

# ── 직진/꺾임 임계 ────────────────────────────────────────────────────────
# remote30_prototype._classify_branch_edges 의 trunk 추적과 **같은 자**를 쓴다.
# 두 판정이 다른 임계를 쓰면 "가지로는 분류됐는데 티는 안 달린" 조합이 생긴다.
TRUNK_TURN_TOL_DEG = 45.0

# ── 엘보 각도 구간 ────────────────────────────────────────────────────────
# 입력 각도는 편향각(0° = 직선, 90° = 직각으로 꺾임)이다. 고를 수 있는 부속이
# 직선(0°)·45° 엘보·90° 엘보 셋뿐이므로 경계는 이웃 사이 중점으로 둔다.
#   0 ─┬─ 22.5 ─┬─ 67.5 ─┬─ 95
#      직선     45° 엘보  90° 엘보
# 22.5° 미만은 45° 엘보보다 직선에 가깝다. 그런데 이 각도가 기록됐다는 것은
# collinear merge 가 직선으로 흡수하기를 거부했다는 뜻이라(COLLINEAR_TOL 12°),
# 직선으로 단정할 수도 없다 — 그래서 부속을 달지 않고 "판정 불가" 로 센다.
ELBOW_STRAIGHT_MAX_DEG = 22.5
ELBOW_45_MAX_DEG = 67.5
# collinear merge 의 ELBOW_MERGE_TOL(95°) 상한. 그보다 큰 편향은 흡수되지 않아
# 애초에 여기로 오지 않는다. 넘어오면 형상이 깨진 것이므로 판정 불가.
ELBOW_MAX_DEG = 95.0

ELBOW_45 = "elbow-45"
ELBOW_90 = "elbow"
TEE = "tee"


def bearing_deg(origin: tuple[float, float], target: tuple[float, float]) -> float:
    """origin 에서 target 을 볼 때의 방위각(도, -180~180)."""
    return math.degrees(math.atan2(target[1] - origin[1], target[0] - origin[0]))


def deflection_deg(bearing_a: float, bearing_b: float) -> float:
    """두 방위각 사이 편향(0~180). 0 이면 같은 방향, 180 이면 되돌아감."""
    return abs((bearing_a - bearing_b + 180.0) % 360.0 - 180.0)


def classify_elbow(angle_deg: float) -> str | None:
    """편향각 → 엘보 종류. 어느 부속으로도 보낼 수 없으면 None."""
    if angle_deg <= ELBOW_STRAIGHT_MAX_DEG or angle_deg > ELBOW_MAX_DEG:
        return None
    return ELBOW_45 if angle_deg <= ELBOW_45_MAX_DEG else ELBOW_90


def elbow_fittings(angles_deg) -> tuple[list[str], int]:
    """편향각 목록 → (엘보 종류 목록, 판정 불가 건수)."""
    kinds: list[str] = []
    unresolved = 0
    for a in angles_deg:
        kind = classify_elbow(float(a))
        if kind is None:
            unresolved += 1
        else:
            kinds.append(kind)
    return kinds, unresolved


def tee_fittings(
    node: tuple[float, float],
    upstream: tuple[float, float] | None,
    downstream: list[tuple[str, tuple[float, float]]],
    *,
    turn_tol_deg: float = TRUNK_TURN_TOL_DEG,
) -> tuple[list[str], int]:
    """분기 노드에서 **분류티가 필요한** 하류 관로를 고른다.

    Args:
        node: 분기점 좌표.
        upstream: 상류(부모) 관로의 반대쪽 끝 좌표. 소스처럼 상류가 없으면 None.
        downstream: 이 노드에서 갈라져 나가는 [(관로 라벨, 반대쪽 끝 좌표)].
        turn_tol_deg: 이 각도 이내로 이어지면 직진 — 직류티라 부속 없음.

    Returns:
        (분류티를 달 관로 라벨, 판정 불가 건수)

        상류 방향을 모르면 어느 갈래가 직진인지 가릴 수 없다. 그때는 기존
        동작대로 전부 티로 두되(부속을 없애는 쪽이 비보수측이라) 판정 불가로
        세어 리포트에 드러낸다.
    """
    labels = [lab for lab, _pos in downstream]
    if len(downstream) < 2:
        # 분기가 아니다 — 관통이면 티가 아니라 (있다면) 엘보다.
        return [], 0
    if upstream is None:
        return labels, len(labels)

    inflow = bearing_deg(upstream, node)
    turning: list[str] = []
    straight = 0
    for lab, pos in downstream:
        if deflection_deg(bearing_deg(node, pos), inflow) <= turn_tol_deg:
            straight += 1
        else:
            turning.append(lab)
    # 직진 갈래가 없으면(전부 꺾임) 어느 것도 spine 이 아니라는 뜻이라 그대로
    # 전부 분류티다. 직진이 둘 이상이면 spine 이 갈라진 것이라 가릴 수 없다 —
    # 하나만 남기지 않고 전부 티로 두되 판정 불가로 센다.
    if straight > 1:
        return labels, straight
    return turning, 0
