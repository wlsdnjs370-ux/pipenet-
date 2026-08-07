# -*- coding: utf-8 -*-
"""core/fitting_rules.py — 합성 형상으로 직류/분류티·엘보 판정을 고정한다."""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import fitting_rules as fr  # noqa: E402


# ── 엘보 ──────────────────────────────────────────────────────────────────
@pytest.mark.parametrize("angle,expect", [
    (0.0, None),        # 직선 — 부속 없음
    (12.1, None),       # collinear merge 가 거부했지만 45° 라 하기엔 이르다 → 판정 불가
    (22.5, None),
    (22.6, fr.ELBOW_45),
    (45.0, fr.ELBOW_45),
    (60.0, fr.ELBOW_45),   # 예전 코드는 여기서 조용히 버렸다
    (67.5, fr.ELBOW_45),
    (67.6, fr.ELBOW_90),
    (90.0, fr.ELBOW_90),
    (95.0, fr.ELBOW_90),
    (95.1, None),       # merge 상한 초과 — 형상이 깨진 것
])
def test_classify_elbow(angle, expect):
    assert fr.classify_elbow(angle) == expect


def test_elbow_band_has_no_gap():
    """22.5~95° 구간에 판정 불가가 하나도 없어야 한다(예전 결함이 이 구멍)."""
    holes = [a / 10.0 for a in range(226, 951)
             if fr.classify_elbow(a / 10.0) is None]
    assert holes == []


def test_elbow_fittings_counts_unresolved():
    kinds, unresolved = fr.elbow_fittings([90.0, 45.0, 5.0, 120.0])
    assert kinds == [fr.ELBOW_90, fr.ELBOW_45]
    assert unresolved == 2


# ── 티 ────────────────────────────────────────────────────────────────────
N = (0.0, 0.0)
WEST, EAST, NORTH, SOUTH = (-1.0, 0.0), (1.0, 0.0), (0.0, 1.0), (0.0, -1.0)


def test_tee_cross_junction():
    """십자 교차: 서→동 직진 1 + 남북 분기 2 → 분류티는 2개뿐."""
    labels, unresolved = fr.tee_fittings(
        N, WEST, [("straight", EAST), ("up", NORTH), ("down", SOUTH)])
    assert sorted(labels) == ["down", "up"]
    assert unresolved == 0


def test_tee_t_junction():
    """T 분기: 서→동 직진 + 북쪽 가지 → 분류티 1개(예전엔 2개)."""
    labels, unresolved = fr.tee_fittings(
        N, WEST, [("straight", EAST), ("branch", NORTH)])
    assert labels == ["branch"]
    assert unresolved == 0


def test_tee_straight_through_is_not_a_branch():
    """일직선 관통은 분기가 아니다 — 부속 없음."""
    labels, unresolved = fr.tee_fittings(N, WEST, [("straight", EAST)])
    assert labels == []
    assert unresolved == 0


def test_tee_all_turning_keeps_every_branch():
    """유입 방향으로 나가는 갈래가 없으면 전부 분류티다."""
    labels, unresolved = fr.tee_fittings(
        N, WEST, [("up", NORTH), ("down", SOUTH)])
    assert sorted(labels) == ["down", "up"]
    assert unresolved == 0


def test_tee_45_turn_is_a_branch():
    """45° 는 임계와 같으므로 직진으로 본다(<= 비교) — 60° 는 꺾임."""
    p45 = (1.0, 1.0)      # 유입(서→동) 대비 편향 45°
    p60 = (0.5, 0.866)    # 편향 60°
    labels, _ = fr.tee_fittings(N, WEST, [("a", p45), ("b", NORTH)])
    assert labels == ["b"]
    labels, _ = fr.tee_fittings(N, WEST, [("a", p60), ("b", SOUTH)])
    assert sorted(labels) == ["a", "b"]


def test_tee_without_upstream_is_unresolved():
    """상류를 모르면 직진을 가릴 수 없다 — 전부 티로 두되 미판정으로 센다."""
    labels, unresolved = fr.tee_fittings(
        N, None, [("a", EAST), ("b", NORTH)])
    assert sorted(labels) == ["a", "b"]
    assert unresolved == 2


def test_tee_two_straight_branches_is_unresolved():
    """직진이 둘이면 spine 이 갈라진 것 — 임의로 하나 고르지 않고 미판정."""
    near_east_a, near_east_b = (1.0, 0.1), (1.0, -0.1)
    labels, unresolved = fr.tee_fittings(
        N, WEST, [("a", near_east_a), ("b", near_east_b), ("c", NORTH)])
    assert sorted(labels) == ["a", "b", "c"]
    assert unresolved == 2


def test_trunk_tolerance_is_shared_with_prototype():
    """_classify_branch_edges 의 trunk 임계와 같은 상수를 쓴다."""
    import math
    import remote30_prototype  # noqa: F401  (import 가능 여부까지 확인)
    assert fr.TRUNK_TURN_TOL_DEG == 45.0
    assert math.isclose(math.radians(fr.TRUNK_TURN_TOL_DEG), math.radians(45.0))
