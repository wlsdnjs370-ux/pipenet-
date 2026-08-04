# -*- coding: utf-8 -*-
"""core/nftc_rules.py — NFTC 룰 엔진 단위 테스트.

f4496b1 에서 삭제된 옛 테스트(구 모듈 경로 + 스칼라 반환 API)를 현행
RuleDecision 기반 API 기준으로 복원한 것이다. 표를 옳게 읽는지만 검사하므로
도면이 바뀌어도 깨지지 않아야 한다.

기준개수 표(§2.1.1)는 **행 순서가 곧 규칙**이라 순서 회귀를 따로 고정한다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import nftc_rules as N  # noqa: E402


# ────────────────────────────────────────────────────────────────────────────
# §2.1.1 기준개수
# ────────────────────────────────────────────────────────────────────────────

def _meta(**kw):
    base = {"floors_total": 5, "use": "other_low", "has_special_combustible": False,
            "head_attach_h_m": 4.0}
    base.update(kw)
    return base


@pytest.mark.parametrize("meta,count,rule_id", [
    (_meta(use="apartment", floors_total=25), 10, "NFTC-211-H"),
    (_meta(use="apartment", floors_total=25, connected_to_basement_parking=True),
     30, "NFTC-211-I"),
    (_meta(use="other_low", floors_total=11), 30, "NFTC-211-G"),
    (_meta(use="underground_station", floors_total=3), 30, "NFTC-211-G"),
    (_meta(use="factory", has_special_combustible=True), 30, "NFTC-211-A"),
    (_meta(use="factory"), 20, "NFTC-211-B"),
    (_meta(use="warehouse"), 20, "NFTC-211-B"),
    (_meta(use="retail"), 30, "NFTC-211-C"),
    (_meta(use="neighborhood"), 20, "NFTC-211-C2"),
    (_meta(use="transit"), 20, "NFTC-211-C2"),
    (_meta(use="other_high", head_attach_h_m=8.0), 20, "NFTC-211-E"),
    (_meta(use="other_low", head_attach_h_m=7.9), 10, "NFTC-211-F"),
])
def test_reference_count_table(meta, count, rule_id):
    d = N.decide_reference_count(meta)
    assert (d.value, d.rule_id) == (count, rule_id)
    assert d.verdict is N.Verdict.PASS


def test_apartment_row_precedes_high_rise_row():
    """행 순서 회귀 — 아파트 행이 '11층 이상' 뒤로 밀리면 10 이 30 으로 뒤집힌다.

    국내 아파트 대부분이 11층 이상이라, 이 한 줄 순서가 기준개수를 3배로
    바꾸고 수원·펌프 용량 전체를 따라 부풀린다.
    """
    assert N.decide_reference_count(_meta(use="apartment", floors_total=50)).value == 10


def test_basement_parking_row_precedes_apartment_row():
    d = N.decide_reference_count(
        _meta(use="apartment", floors_total=50, connected_to_basement_parking=True))
    assert (d.value, d.rule_id) == (30, "NFTC-211-I")


def test_other_use_8m_boundary_is_inclusive_upward():
    """8 m 는 '이상' 쪽(20개)에 붙는다. 경계를 반대로 잡으면 과소설계다."""
    assert N.decide_reference_count(_meta(use="other_low", head_attach_h_m=8.0)).value == 20
    assert N.decide_reference_count(_meta(use="other_low", head_attach_h_m=7.999)).value == 10


def test_unknown_use_is_review_not_silent_minimum():
    """미상 용도를 '그 밖의 것 10개'로 쓸어담으면 조용한 과소설계가 된다."""
    d = N.decide_reference_count({"use": "spaceport", "floors_total": 3})
    assert d.verdict is N.Verdict.REVIEW
    assert d.value == 20
    assert d.rule_id == "NFTC-211-FALLBACK"


def test_other_use_without_attach_height_falls_to_review():
    """부착높이가 없으면 F 행(0 < h < 8)에 걸리지 않고 REVIEW 로 빠져야 한다."""
    d = N.decide_reference_count({"use": "other_low", "floors_total": 3})
    assert d.verdict is N.Verdict.REVIEW


def test_reference_count_decision_is_serializable():
    d = N.decide_reference_count(_meta(use="retail"))
    obj = d.to_dict()
    assert obj["verdict"] == "PASS"
    assert obj["value"] == 30
    assert obj["trace"]["NFTC"].startswith("NFTC 103")


# ────────────────────────────────────────────────────────────────────────────
# §2.7.6 표시온도
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("ambient,lo,hi", [
    (0.0, -1e9, 79.0),
    (38.999, -1e9, 79.0),
    (39.0, 79.0, 121.0),
    (63.9, 79.0, 121.0),
    (64.0, 121.0, 162.0),
    (105.9, 121.0, 162.0),
    (106.0, 162.0, 1e9),
])
def test_temperature_rating_bands(ambient, lo, hi):
    d = N.decide_temperature_rating(ambient)
    assert d.rule_id == "NFTC-276-MAIN"
    assert (d.value["min_c"], d.value["max_c"]) == (lo, hi)


@pytest.mark.parametrize("flag", ["is_factory_4m_high", "is_warehouse_4m_high",
                                  "is_rack_storage"])
def test_temperature_rating_4m_proviso_overrides_ambient(flag):
    d = N.decide_temperature_rating(20.0, **{flag: True})
    assert d.rule_id == "NFTC-276-PROVISO"
    assert d.value["min_c"] == 121.0
    assert d.value["max_c"] is None


def test_hb_temperature_formula_is_auxiliary_only():
    assert N.hb_temperature_formula(72.0) == pytest.approx(0.9 * 72.0 - 27.3)


# ────────────────────────────────────────────────────────────────────────────
# §2.7.3 수평거리 R
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("kwargs,r_m,rule_id", [
    (dict(room_use="office"), 2.1, "NFTC-273-4"),
    (dict(room_use="office", structure="fire_resistant"), 2.3, "NFTC-273-4-FR"),
    (dict(room_use="stage"), 1.7, "NFTC-273-1"),
    (dict(room_use="office", has_special_combustible=True), 1.7, "NFTC-273-1"),
    (dict(room_use="warehouse", is_rack_storage=True), 2.5, "NFTC-273-2"),
    (dict(room_use="apartment_living"), 3.2, "NFTC-273-3"),
])
def test_horizontal_distance(kwargs, r_m, rule_id):
    d = N.decide_horizontal_distance(**kwargs)
    assert (d.value, d.rule_id) == (r_m, rule_id)


def test_special_combustible_beats_rack_storage():
    """특수가연물 랙크식은 2.5 가 아니라 1.7 이다 — 더 촘촘한 쪽이 이겨야 한다."""
    d = N.decide_horizontal_distance(
        room_use="warehouse", is_rack_storage=True, has_special_combustible=True)
    assert d.value == 1.7


def test_apartment_living_beats_fire_resistant_structure():
    d = N.decide_horizontal_distance(
        room_use="apartment_living", structure="fire_resistant")
    assert d.value == 3.2


# ────────────────────────────────────────────────────────────────────────────
# §2.7.5.5 조기반응형 5종
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("room_use", [
    "apartment_living", "welfare_living", "officetel_bedroom",
    "hotel_bedroom", "hospital_ward",
])
def test_fast_response_mandated_locations(room_use):
    d = N.is_fast_response_required(room_use)
    assert d.value is True
    assert d.verdict is N.Verdict.PASS


def test_fast_response_not_mandated_is_na_not_fail():
    """의무 장소가 아닌 것은 '위반'이 아니라 '해당 없음'이다."""
    d = N.is_fast_response_required("office")
    assert d.value is False
    assert d.verdict is N.Verdict.NA


def test_fast_response_mandate_covers_exactly_five_locations():
    assert len(N._FAST_RESPONSE_LOCATIONS) == 5


# ────────────────────────────────────────────────────────────────────────────
# §2.7.7.1 살수공간 (60 cm / 벽 10 cm)
# ────────────────────────────────────────────────────────────────────────────

def _box(cx, cy, half=0.2, **kw):
    d = {"id": kw.pop("id", "o1"),
         "polygon": [(cx - half, cy - half), (cx + half, cy - half),
                     (cx + half, cy + half), (cx - half, cy + half)]}
    d.update(kw)
    return d


def test_clearance_passes_for_distant_obstacle():
    d = N.validate_head_clearance(head_xy=(0.0, 0.0), obstacles=[_box(2.0, 2.0)])
    assert d.verdict is N.Verdict.PASS


def test_clearance_fails_for_close_obstacle():
    d = N.validate_head_clearance(head_xy=(0.0, 0.0), obstacles=[_box(0.5, 0.0)])
    assert d.verdict is N.Verdict.FAIL
    assert d.value["violating_obstacle"] == "o1"


def test_head_inside_obstacle_is_zero_distance_not_pass():
    """장애물 한가운데 놓인 헤드 — 변까지 거리만 재면 통과로 나온다."""
    d = N.validate_head_clearance(head_xy=(0.0, 0.0), obstacles=[_box(0.0, 0.0, half=1.0)])
    assert d.verdict is N.Verdict.FAIL
    assert d.value["distance_m"] == 0.0


def test_wall_uses_10cm_exception_not_skipped():
    """is_wall 은 검사 면제가 아니라 10 cm 기준으로의 이동이다."""
    near_wall = _box(0.35, 0.0, id="w1", is_wall=True)
    assert N.validate_head_clearance(
        head_xy=(0.0, 0.0), obstacles=[near_wall]).verdict is N.Verdict.PASS
    touching_wall = _box(0.05, 0.0, half=0.01, id="w2", is_wall=True)
    d = N.validate_head_clearance(head_xy=(0.0, 0.0), obstacles=[touching_wall])
    assert d.verdict is N.Verdict.FAIL
    assert d.value["violating_wall"] == "w2"


def test_obstacle_without_geometry_is_review_not_pass():
    d = N.validate_head_clearance(
        head_xy=(0.0, 0.0), obstacles=[{"id": "tray", "polygon": []}])
    assert d.verdict is N.Verdict.REVIEW
    assert d.value["obstacles_without_footprint"] == ["tray"]


def test_clearance_accepts_bbox_instead_of_polygon():
    d = N.validate_head_clearance(
        head_xy=(0.0, 0.0), obstacles=[{"id": "b", "bbox": (0.3, -0.1, 0.5, 0.1)}])
    assert d.verdict is N.Verdict.FAIL


# ────────────────────────────────────────────────────────────────────────────
# §2.7.7.2 헤드↔천장
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("head_z,ceiling_z,verdict", [
    (2.75, 3.0, N.Verdict.PASS),   # 25 cm
    (2.70, 3.0, N.Verdict.PASS),   # 30 cm 경계 포함
    (2.69, 3.0, N.Verdict.FAIL),   # 31 cm
])
def test_head_to_ceiling_default_limit(head_z, ceiling_z, verdict):
    assert N.validate_head_to_ceiling(
        head_z_m=head_z, ceiling_z_m=ceiling_z).verdict is verdict


def test_head_to_ceiling_beam_exception_extends_to_55cm():
    kw = dict(head_z_m=2.5, ceiling_z_m=3.0)
    assert N.validate_head_to_ceiling(**kw).verdict is N.Verdict.FAIL
    assert N.validate_head_to_ceiling(
        **kw, has_beam=True, beam_clear_m=0.55).verdict is N.Verdict.PASS


def test_beam_exception_requires_qualifying_clearance():
    """보가 있어도 천장~보 하단이 55 cm 미만이면 30 cm 기준 그대로다."""
    d = N.validate_head_to_ceiling(head_z_m=2.5, ceiling_z_m=3.0,
                                   has_beam=True, beam_clear_m=0.4)
    assert d.verdict is N.Verdict.FAIL
    assert d.value["limit_m"] == 0.30


def test_head_above_ceiling_is_fail():
    d = N.validate_head_to_ceiling(head_z_m=3.2, ceiling_z_m=3.0)
    assert d.verdict is N.Verdict.FAIL
    assert d.rule_id == "NFTC-2772-INVALID"


# ────────────────────────────────────────────────────────────────────────────
# §2.13 겸용 수원
# ────────────────────────────────────────────────────────────────────────────

def test_combined_water_supply_sums_and_takes_max_head():
    res = N.decide_combined_water_supply({
        "sprinkler": {"v_m3": 32.0, "q_lpm": 1600.0, "h_m": 60.0},
        "indoor_hydrant": {"v_m3": 5.2, "q_lpm": 260.0, "h_m": 50.0},
    })
    assert res.tank_total_m3 == pytest.approx(37.2)
    assert res.pump_total_lpm == pytest.approx(1860.0)
    assert res.pump_required_head_m == pytest.approx(60.0)
    assert res.tank_min_total_m3 == pytest.approx(37.2 / 0.8)
    assert set(res.systems) == {"sprinkler", "indoor_hydrant"}


def test_combined_water_supply_can_disable_sharing():
    res = N.decide_combined_water_supply(
        {"sprinkler": {"v_m3": 32.0, "q_lpm": 1600.0, "h_m": 60.0}},
        use_combined_tank=False, use_combined_pump=False)
    assert res.tank_total_m3 == 0.0
    assert res.pump_total_lpm == 0.0


def test_hose_connection_count_is_clamped():
    small = N.decide_combined_water_supply(
        {"sprinkler": {"v_m3": 1.0, "zone_area_m2": 100.0}})
    huge = N.decide_combined_water_supply(
        {"sprinkler": {"v_m3": 1.0, "zone_area_m2": 900000.0}})
    assert small.hose_connection_count == 1
    assert huge.hose_connection_count == 5


# ────────────────────────────────────────────────────────────────────────────
# NFTC 103B — ESFR
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("ceiling,k,in_rack,review", [
    (9.1, 200, False, False),
    (12.0, 320, False, False),
    (13.7, 360, False, True),
    (14.0, 360, True, True),
])
def test_esfr_k_factor_bands(ceiling, k, in_rack, review):
    d = N.decide_esfr_branch(room_use="rack_storage", ceiling_h_m=ceiling)
    assert d.activated is True
    assert d.k_lpm_bar05 == k
    assert d.in_rack_required is in_rack
    assert d.human_review_required is review


def test_esfr_warehouse_activation_threshold():
    assert N.decide_esfr_branch(room_use="warehouse", ceiling_h_m=9.0).activated is False
    assert N.decide_esfr_branch(room_use="warehouse", ceiling_h_m=9.1).activated is True


def test_esfr_not_activated_for_ordinary_use():
    d = N.decide_esfr_branch(room_use="office", ceiling_h_m=20.0)
    assert d.activated is False
    assert d.k_lpm_bar05 is None


def test_ev_charging_takes_k115_branch_without_esfr():
    d = N.decide_esfr_branch(room_use="EV_charging", ceiling_h_m=4.0)
    assert d.activated is False
    assert d.k_lpm_bar05 == 115


# ────────────────────────────────────────────────────────────────────────────
# §2.6.1.6 경보 캐스케이드
# ────────────────────────────────────────────────────────────────────────────

def test_alarm_cascade_below_threshold_is_all_floors():
    d = N.decide_alarm_cascade(fire_floor=3, total_floors=10)
    assert d.verdict is N.Verdict.NA
    assert d.value["all_floors"] is True


def test_apartment_threshold_is_16_not_11():
    assert N.decide_alarm_cascade(
        fire_floor=3, total_floors=15, is_apartment=True).verdict is N.Verdict.NA
    assert N.decide_alarm_cascade(
        fire_floor=3, total_floors=16, is_apartment=True).verdict is N.Verdict.PASS


def test_alarm_cascade_upper_four_floors():
    d = N.decide_alarm_cascade(fire_floor=3, total_floors=20)
    assert d.value["alarm_floors"] == [3, 4, 5, 6, 7]


def test_alarm_cascade_clips_at_top_floor():
    d = N.decide_alarm_cascade(fire_floor=19, total_floors=20)
    assert d.value["alarm_floors"] == [19, 20]


def test_alarm_cascade_first_floor_includes_basement():
    d = N.decide_alarm_cascade(fire_floor=1, total_floors=20)
    assert "B" in d.value["alarm_floors"]


def test_alarm_cascade_basement_includes_all_basement():
    d = N.decide_alarm_cascade(fire_floor=-1, total_floors=20)
    assert "all_basement" in d.value["alarm_floors"]


# ────────────────────────────────────────────────────────────────────────────
# §2.5.10 행거 · 최소 압력/유량 상수
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("role,spacing", [
    ("branch", 3.5), ("cross_main", 4.5), ("main", 4.5), ("riser", 4.5),
])
def test_hanger_max_spacing(role, spacing):
    assert N.hanger_max_spacing_m(role) == spacing


def test_scalar_minimums():
    assert N.head_to_hanger_min_m() == 0.08
    assert N.head_pressure_min_mpa() == 0.1
    assert N.head_pressure_max_mpa() == 1.2
    assert N.head_flow_min_lpm() == 80.0
    assert N.emergency_power_min_minutes() == 20.0


# ────────────────────────────────────────────────────────────────────────────
# 요약 집계
# ────────────────────────────────────────────────────────────────────────────

def test_summarize_counts_and_overall():
    decisions = [
        N.decide_reference_count(_meta(use="retail")),                 # PASS
        N.is_fast_response_required("office"),                         # NA
        N.decide_reference_count({"use": "spaceport"}),                # REVIEW
        N.validate_head_to_ceiling(head_z_m=2.0, ceiling_z_m=3.0),     # FAIL
    ]
    s = N.summarize_nftc_decisions(decisions)
    assert s["counts"] == {"PASS": 1, "FAIL": 1, "REVIEW": 1, "NA": 1}
    assert s["overall"] == "FAIL"
    assert len(s["fails"]) == 1 and len(s["reviews"]) == 1


def test_summarize_review_alone_does_not_fail_overall():
    s = N.summarize_nftc_decisions([N.decide_reference_count({"use": "spaceport"})])
    assert s["overall"] == "PASS"
    assert len(s["reviews"]) == 1


def test_public_api_is_exported():
    for name in N.__all__:
        assert hasattr(N, name), f"__all__ 에 있으나 모듈에 없음: {name}"
