# -*- coding: utf-8 -*-
"""core/hb_rules.py — 한백 설계지침서 §2.4 룰 엔진 테스트.

NFTC 보다 엄격한 값(체절 120%)과 NFTC 가 침묵하는 영역(유속 6/10 m/s,
Case 1~5 위상)이 이 모듈의 존재 이유다. 완화 방향 회귀를 잡는 것이 목적이라
경계값은 양쪽(통과/불통과)을 모두 고정한다.

실행::

    python -m pytest tests/test_hb_rules.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import hb_rules as H  # noqa: E402
from nftc_rules import Verdict  # noqa: E402


# ────────────────────────────────────────────────────────────────────────────
# §2.4.1 시스템 선정
# ────────────────────────────────────────────────────────────────────────────

def test_default_system_is_wet():
    d = H.decide_system_type()
    assert d.value == H.SystemType.WET.value
    assert d.verdict is Verdict.PASS


def test_open_heads_beats_freezing_risk():
    """개방형 헤드 요구는 동결 우려보다 우선한다 (일제살수식 확정)."""
    d = H.decide_system_type(needs_open_heads=True, has_freezing_risk=True,
                             detector_priority=True, room_use="무대부")
    assert d.value == H.SystemType.DELUGE.value
    assert d.rule_id == "HB-241-DELUGE"


def test_freezing_plus_detector_is_double_interlock():
    d = H.decide_system_type(has_freezing_risk=True, detector_priority=True)
    assert d.value == H.SystemType.PREACTION_DOUBLE.value


def test_freezing_alone_is_dry():
    assert H.decide_system_type(has_freezing_risk=True).value == H.SystemType.DRY.value


def test_detector_alone_is_single_interlock():
    assert H.decide_system_type(detector_priority=True).value == \
        H.SystemType.PREACTION_SINGLE.value


# ────────────────────────────────────────────────────────────────────────────
# §2.4.16 Case 1~5 위상
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("height, expected", [
    (30.0, H.HBCase.CASE_1),
    (60.0, H.HBCase.CASE_1),     # 경계 포함
    (60.1, H.HBCase.CASE_2A),
    (80.0, H.HBCase.CASE_2A),    # 경계 포함
    (80.1, H.HBCase.CASE_3A),
    (120.0, H.HBCase.CASE_3A),   # 경계 포함
])
def test_case_bands_by_height(height, expected):
    assert H.decide_hb_case(building_height_m=height).case is expected


def test_case1_needs_no_prv():
    d = H.decide_hb_case(building_height_m=50.0)
    assert d.pump_location == "basement"
    assert d.prv_required is False
    assert (d.rated_head_max_m, d.churn_head_max_m) == (100.0, 120.0)


def test_rooftop_infeasible_switches_to_basement_prv():
    """옥상 수조 불가 → 지하 펌프 + 감압. 감압 없이 지하로만 내려보내면 안 된다."""
    d = H.decide_hb_case(building_height_m=70.0, rooftop_tank_feasible=False)
    assert d.case is H.HBCase.CASE_2C
    assert (d.pump_location, d.prv_required) == ("basement", True)
    d3 = H.decide_hb_case(building_height_m=100.0, rooftop_tank_feasible=False)
    assert d3.case is H.HBCase.CASE_3B
    assert d3.prv_required is True


def test_rooftop_case_pressurizes_only_40m():
    """옥상안의 가압구간은 건물 높이가 아니라 40 m 고정 — 나머지는 자연낙차."""
    d = H.decide_hb_case(building_height_m=100.0)
    assert d.pressurized_zone_m == 40.0
    assert d.natural_drop_zone_m == pytest.approx(60.0)
    assert d.rated_head_max_m == pytest.approx(140.0)
    assert d.churn_head_max_m == pytest.approx(168.0)


def test_material_change_boundary_is_rated_minus_120m():
    d = H.decide_hb_case(building_height_m=100.0)
    assert d.pipe_material_change_at_m == pytest.approx(20.0)


def test_supertall_refuge_interval_selects_case4a_or_5a():
    over = H.decide_hb_case(building_height_m=200.0, refuge_floor_interval_m=60.0)
    assert over.case is H.HBCase.CASE_4A
    wide = H.decide_hb_case(building_height_m=200.0, refuge_floor_interval_m=100.0)
    assert wide.case is H.HBCase.CASE_5A
    assert wide.natural_drop_zone_m == 120.0   # 120 m마다 감압


def test_supertall_without_refuge_interval_defaults_to_case4a():
    """피난안전층 간격 미상 → 80 m 로 가정(원칙안). None 이 5a 로 새면 안 된다."""
    assert H.decide_hb_case(building_height_m=200.0).case is H.HBCase.CASE_4A


def test_prv_secondary_is_always_4bar():
    for h in (50.0, 70.0, 100.0, 200.0):
        assert H.decide_hb_case(building_height_m=h).prv_secondary_bar == 4.0


# ────────────────────────────────────────────────────────────────────────────
# §2.4.5 배관 재질 · 내경
# ────────────────────────────────────────────────────────────────────────────

def test_material_boundary_1_2mpa_is_inclusive_for_3507():
    assert H.decide_pipe_material(1.2).value == "KSD 3507"
    assert H.decide_pipe_material(1.201).value == "KSD 3562"


def test_inner_diameter_3562_is_never_larger_than_3507():
    """고압관(3562)은 두께가 두꺼워 내경이 작다. 뒤집히면 마찰손실이 과소평가된다."""
    for nominal in ("25A", "50A", "100A", "150A", "300A"):
        d3507 = H.get_inner_diameter_mm(nominal, "KSD 3507")
        d3562 = H.get_inner_diameter_mm(nominal, "KSD 3562")
        assert d3562 < d3507


def test_inner_diameter_accepts_lowercase_and_whitespace():
    assert H.get_inner_diameter_mm(" 25a ", "KSD 3507") == 27.5


def test_unknown_nominal_returns_none_not_zero():
    """미등록 호칭에 0 을 돌려주면 유속이 무한대로 튀어 조용히 합격한다."""
    assert H.get_inner_diameter_mm("999A", "KSD 3507") is None


def test_unknown_material_falls_back_to_3507():
    assert H.get_inner_diameter_mm("50A", "") == H.get_inner_diameter_mm("50A", "KSD 3507")


# ────────────────────────────────────────────────────────────────────────────
# §2.4.5 유속 한계 (NFTC 2.2.1.10 침묵 영역)
# ────────────────────────────────────────────────────────────────────────────

def test_velocity_limits_are_6_and_10():
    assert (H.BRANCH_PIPE_V_LIMIT, H.MAIN_PIPE_V_LIMIT) == (6.0, 10.0)


@pytest.mark.parametrize("role, v, expected", [
    ("branch", 5.99, Verdict.PASS),
    ("branch", 6.0, Verdict.PASS),     # 경계 포함
    ("branch", 6.01, Verdict.FAIL),
    ("other", 6.01, Verdict.PASS),     # 가지 한계가 그 밖 배관에 새면 안 된다
    ("other", 10.0, Verdict.PASS),
    ("other", 10.01, Verdict.FAIL),
])
def test_velocity_verdicts(role, v, expected):
    d = H.validate_velocity(pipe_role=role, velocity_mps=v)
    assert d.verdict is expected
    assert d.value["ok"] is (expected is Verdict.PASS)


def test_unknown_pipe_role_uses_main_limit_not_skip():
    d = H.validate_velocity(pipe_role="riser", velocity_mps=12.0)
    assert d.verdict is Verdict.FAIL
    assert d.value["limit_mps"] == H.MAIN_PIPE_V_LIMIT


# ────────────────────────────────────────────────────────────────────────────
# §2.4.16 체절압 (HB 120% / NFTC 140%)
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("rated, churn, expected", [
    (100.0, 110.0, Verdict.PASS),
    (100.0, 120.0, Verdict.PASS),     # HB 경계 포함
    (100.0, 120.1, Verdict.REVIEW),   # HB 초과 · NFTC 이내 → 인간 검토
    (100.0, 140.0, Verdict.REVIEW),   # NFTC 경계 포함
    (100.0, 140.1, Verdict.FAIL),
])
def test_churn_three_way_verdict(rated, churn, expected):
    assert H.validate_churn_pressure(rated, churn).verdict is expected


def test_hb_churn_limit_is_stricter_than_nftc():
    assert H.CHURN_PRESSURE_MAX_RATIO_HB < H.CHURN_PRESSURE_MAX_RATIO_NFTC


def test_zero_rated_head_fails_instead_of_dividing_by_zero():
    """정격양정 0(미입력)은 조용히 통과시키지 않고 FAIL 로 드러낸다."""
    d = H.validate_churn_pressure(0.0, 100.0)
    assert d.verdict is Verdict.FAIL
    assert d.value["ratio"] == float("inf")


# ────────────────────────────────────────────────────────────────────────────
# §2.4.2 방호구역 분할
# ────────────────────────────────────────────────────────────────────────────

def test_single_zone_at_3000m2_boundary():
    zones = H.decide_zone_partition(
        floor_area_m2=3000.0, estimated_head_count=90, floor_label="F3")
    assert len(zones) == 1
    assert zones[0].zone_id == "Z-F3-1"
    assert zones[0].area_m2 == 3000.0


def test_area_over_limit_splits_evenly():
    zones = H.decide_zone_partition(
        floor_area_m2=7000.0, estimated_head_count=90, floor_label="F3")
    assert len(zones) == 3
    assert sum(z.area_m2 for z in zones) == pytest.approx(7000.0)
    assert [z.zone_id for z in zones] == ["Z-F3-1", "Z-F3-2", "Z-F3-3"]
    assert all(z.head_count_estimate == 30 for z in zones)


def test_grid_layout_raises_area_cap_to_3700():
    kw = dict(floor_area_m2=3500.0, estimated_head_count=90, floor_label="F3")
    assert len(H.decide_zone_partition(**kw)) == 2
    assert len(H.decide_zone_partition(is_grid_layout=True, **kw)) == 1


def test_multi_floor_grouping_only_when_small_or_loft():
    big = H.decide_zone_partition(
        floor_area_m2=1000.0, estimated_head_count=40, floor_label="F3")
    assert big[0].multi_floor_grouping is False
    few = H.decide_zone_partition(
        floor_area_m2=1000.0, estimated_head_count=10, floor_label="F3")
    assert few[0].multi_floor_grouping is True
    loft = H.decide_zone_partition(
        floor_area_m2=1000.0, estimated_head_count=40, floor_label="F3",
        is_apartment_loft=True)
    assert loft[0].multi_floor_grouping is True


def test_split_zones_drop_multi_floor_grouping():
    """분할된 구역을 다시 여러 층으로 묶으면 면적 한계가 무효가 된다."""
    zones = H.decide_zone_partition(
        floor_area_m2=9000.0, estimated_head_count=5, floor_label="B1")
    assert len(zones) > 1
    assert all(z.multi_floor_grouping is False for z in zones)


def test_dry_pipe_volume_constants():
    assert H.DRY_PIPE_VOLUME_MAX_L == 2840
    assert H.DRY_PIPE_VOLUME_FAST_OPEN_THRESHOLD_L == 1890
    assert H.DRY_PIPE_VOLUME_FAST_OPEN_THRESHOLD_L < H.DRY_PIPE_VOLUME_MAX_L


# ────────────────────────────────────────────────────────────────────────────
# §2.4.7 행거
# ────────────────────────────────────────────────────────────────────────────

def test_branch_hanger_between_each_head_pair():
    pos = H.hanger_positions_along_pipe(
        pipe_length_m=6.0, pipe_role="branch", head_positions_m=[0.0, 3.0, 6.0])
    assert pos == [1.5, 4.5]


def test_branch_gap_over_3_5m_gets_subdivided():
    pos = H.hanger_positions_along_pipe(
        pipe_length_m=8.0, pipe_role="branch", head_positions_m=[0.0, 8.0])
    assert pos == [3.5, 4.0, 7.0]


def test_hanger_too_close_to_head_is_dropped():
    """헤드↔행거 8 cm 미만이면 배치하지 않는다 (NFTC 2.5.10 단서)."""
    pos = H.hanger_positions_along_pipe(
        pipe_length_m=1.0, pipe_role="branch", head_positions_m=[0.0, 0.1])
    assert pos == []


def test_cross_main_uses_uniform_4_5m_spacing():
    pos = H.hanger_positions_along_pipe(pipe_length_m=10.0, pipe_role="cross_main")
    assert pos == [3.333, 6.667]
    assert max(pos) <= 10.0


def test_short_pipe_still_gets_one_hanger():
    assert len(H.hanger_positions_along_pipe(pipe_length_m=1.0, pipe_role="main")) == 1


def test_branch_without_head_positions_falls_back_to_uniform():
    pos = H.hanger_positions_along_pipe(pipe_length_m=10.0, pipe_role="branch")
    assert len(pos) == 2      # 3.5 m 간격 → int(10/3.5) = 2


# ────────────────────────────────────────────────────────────────────────────
# 공개 API
# ────────────────────────────────────────────────────────────────────────────

def test_all_exports_exist():
    for name in H.__all__:
        assert hasattr(H, name), name


def test_unit_conversion_constants_are_distinct():
    """m수두→bar 계수를 kgf/cm² 자리에 쓰면 2% 어긋난다 — 두 상수가 섞이면 안 된다."""
    assert H.KGFCM2_TO_BAR == pytest.approx(0.980665)
    assert H.M_H2O_PER_KGFCM2 == 10.0
