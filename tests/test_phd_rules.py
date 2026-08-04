# -*- coding: utf-8 -*-
"""core/phd_rules.py — 박사논문 압력존 · 공백변수 · 3대 불균형 지표 테스트.

이 모듈은 "같은 건물 · 같은 법령인데 설계자마다 결과가 다른" 원인을 공백변수
5종으로 좁혀 자동 표준화한다. 따라서 검증 근거가 없을 때 지표를 0.0/True 로
채워 조용히 합격시키지 않는지(= None 유지)가 핵심 회귀 방어선이다.

실행::

    python -m pytest tests/test_phd_rules.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import phd_rules as P  # noqa: E402
from hb_rules import HBCase, decide_hb_case  # noqa: E402


def _floors(*pairs):
    return [{"label": lbl, "z_m": z} for lbl, z in pairs]


# ────────────────────────────────────────────────────────────────────────────
# 1. 압력존 분류 (HSP / MSP / LSP / LLSP)
# ────────────────────────────────────────────────────────────────────────────

def test_basement_pump_case_splits_only_by_ground_level():
    """지하펌프안(Case 1)은 지상 전체가 HSP, 지하만 LLSP."""
    case1 = decide_hb_case(building_height_m=50.0)
    zones = P.classify_pressure_zones(
        floors=_floors(("F10", 30.0), ("F1", 0.0), ("B1", -3.5)),
        hb_case=case1, elevated_tank_z_m=50.0)
    assert [z.zone for z in zones] == [
        P.PressureZone.HSP, P.PressureZone.HSP, P.PressureZone.LLSP]


def test_rooftop_case_produces_hsp_msp_llsp_bands():
    case2a = decide_hb_case(building_height_m=70.0)
    zones = P.classify_pressure_zones(
        floors=_floors(("F20", 70.0), ("F12", 35.0), ("F8", 25.0), ("B1", -3.5)),
        hb_case=case2a, elevated_tank_z_m=70.0)
    assert [z.zone for z in zones] == [
        P.PressureZone.HSP,   # 수조 직하 40 m 이내
        P.PressureZone.HSP,
        P.PressureZone.MSP,   # 자연낙차만, 12 bar 이내
        P.PressureZone.LLSP,
    ]


def test_overpressured_ground_floor_becomes_lsp_with_prv():
    """자연낙차가 12 bar 를 넘는 지상층은 감압존 — PRV 없이 두면 헤드가 터진다."""
    case3a = decide_hb_case(building_height_m=120.0)
    zones = P.classify_pressure_zones(
        floors=_floors(("F1", 0.0),), hb_case=case3a, elevated_tank_z_m=130.0)
    assert zones[0].zone is P.PressureZone.LSP
    assert zones[0].requires_prv is True
    assert zones[0].natural_drop_pressure_bar > 12.0


def test_hsp_band_is_40m_below_pump():
    case2a = decide_hb_case(building_height_m=70.0)
    zones = P.classify_pressure_zones(
        floors=_floors(("just_in", 30.0), ("just_out", 29.9)),
        hb_case=case2a, elevated_tank_z_m=70.0, pump_z_m=70.0)
    assert zones[0].zone is P.PressureZone.HSP
    assert zones[1].zone is not P.PressureZone.HSP


def test_floor_above_tank_has_zero_not_negative_pressure():
    case2a = decide_hb_case(building_height_m=70.0)
    zones = P.classify_pressure_zones(
        floors=_floors(("PH", 80.0),), hb_case=case2a, elevated_tank_z_m=70.0)
    assert zones[0].natural_drop_pressure_bar == 0.0


def test_zone_vertical_span_recommendation():
    assert P.ZONE_VERTICAL_SPAN_M == (40.0, 50.0)


# ────────────────────────────────────────────────────────────────────────────
# 2. 기준구역 생성 (공백변수 ①)
# ────────────────────────────────────────────────────────────────────────────

def _sample_floors():
    case2a = decide_hb_case(building_height_m=70.0)
    return P.classify_pressure_zones(
        floors=_floors(("F20", 70.0), ("F16", 55.0), ("F12", 35.0),
                       ("F8", 25.0), ("F2", 5.0), ("B1", -3.5)),
        hb_case=case2a, elevated_tank_z_m=70.0), case2a


def test_reference_zones_have_top_and_bottom_per_zone():
    floors, _ = _sample_floors()
    refs = P.generate_reference_zones(floors)
    by_zone = {}
    for r in refs:
        by_zone.setdefault(r["zone"], []).append(r["position"])
    assert by_zone["hsp"] == ["top", "bottom"]
    assert by_zone["msp"] == ["top", "bottom"]


def test_single_floor_zone_emits_top_only():
    """층이 하나뿐인 존에 top/bottom 두 시나리오를 만들면 중복 계산이 된다."""
    floors, _ = _sample_floors()
    refs = P.generate_reference_zones(floors)
    llsp = [r for r in refs if r["zone"] == "llsp"]
    assert len(llsp) == 1
    assert llsp[0]["position"] == "top"


def test_reference_zone_count_never_exceeds_eight():
    floors, _ = _sample_floors()
    assert len(P.generate_reference_zones(floors)) <= 8


def test_llsp_top_is_conditional_priority():
    floors, _ = _sample_floors()
    refs = P.generate_reference_zones(floors)
    llsp = next(r for r in refs if r["zone"] == "llsp")
    assert llsp["priority"] == "조건부"


def test_reference_zone_top_is_highest_floor():
    floors, _ = _sample_floors()
    refs = P.generate_reference_zones(floors)
    hsp_top = next(r for r in refs if r["zone"] == "hsp" and r["position"] == "top")
    assert hsp_top["floor"] == "F20"


def test_no_floors_yields_no_reference_zones():
    assert P.generate_reference_zones([]) == []


# ────────────────────────────────────────────────────────────────────────────
# 3. 공백변수 5종
# ────────────────────────────────────────────────────────────────────────────

def _discretionary(**kw):
    floors, case = _sample_floors()
    params = dict(floors=floors, hb_case=case, elevated_tank_z_m=70.0,
                  pump_rated_q_lpm=2400.0, pump_rated_h_m=110.0,
                  pump_churn_h_m=130.0)
    params.update(kw)
    return P.decide_discretionary_variables(**params)


def test_pipenet_verified_equivalent_lengths():
    """FX/AV/PV 등가길이는 PIPENET 검증 표준값 — 임의로 흔들리면 수리계산이 어긋난다."""
    dv = _discretionary()
    assert dv.fx_equivalent_length_m == 0.6
    assert dv.fx_inner_diameter_mm == 21.6
    assert dv.fx_c_value == 120
    assert dv.av_equivalent_length_m == 12.9
    assert dv.pv_equivalent_length_m == 10.1


def test_natural_drop_start_is_highest_msp_floor():
    dv = _discretionary()
    assert dv.natural_drop_start_floor == "F8"   # F12(35 m)까지는 HSP 대역


def test_no_msp_floor_gives_none_not_first_floor():
    case1 = decide_hb_case(building_height_m=50.0)
    floors = P.classify_pressure_zones(
        floors=_floors(("F10", 30.0), ("B1", -3.5)),
        hb_case=case1, elevated_tank_z_m=50.0)
    dv = _discretionary(floors=floors, hb_case=case1, elevated_tank_z_m=50.0)
    assert dv.natural_drop_start_floor is None


def test_prv_settings_generated_only_for_prv_floors():
    case3a = decide_hb_case(building_height_m=120.0)
    floors = P.classify_pressure_zones(
        floors=_floors(("F30", 125.0), ("F1", 0.0)),
        hb_case=case3a, elevated_tank_z_m=130.0)
    dv = _discretionary(floors=floors, hb_case=case3a, elevated_tank_z_m=130.0)
    assert [p["floor"] for p in dv.prv_settings] == ["F1"]
    prv = dv.prv_settings[0]
    assert prv["p2_bar"] == 4.0
    assert prv["delta_p_bar"] == pytest.approx(prv["p1_bar"] - 4.0, abs=0.01)


def test_q150_check_stays_none_without_pump_curve():
    """펌프 성능곡선이 없으면 150% 유량점은 판정 불가 — True/False 로 채우면 거짓 합격."""
    assert _discretionary().pump_check_q150_validated is None


def test_churn_check_is_none_when_rated_head_missing():
    assert _discretionary(pump_rated_h_m=0.0).pump_check_churn_le_120pct is None


@pytest.mark.parametrize("rated, churn, expected", [
    (100.0, 120.0, True),    # 경계 포함
    (100.0, 120.1, False),
])
def test_churn_120pct_boundary(rated, churn, expected):
    dv = _discretionary(pump_rated_h_m=rated, pump_churn_h_m=churn)
    assert dv.pump_check_churn_le_120pct is expected


# ────────────────────────────────────────────────────────────────────────────
# 4. 계산 시나리오 생성
# ────────────────────────────────────────────────────────────────────────────

def test_scenarios_mirror_reference_zones_plus_maxq():
    dv = _discretionary()
    scs = P.generate_calculation_scenarios(discretionary=dv)
    assert len(scs) == len(dv.reference_zones) + 1
    assert scs[-1].scenario_id == "S-MAX-Q"


def test_k115_scenario_carries_head_k_override():
    dv = _discretionary()
    scs = P.generate_calculation_scenarios(discretionary=dv, has_k115_zones=True)
    k115 = next(s for s in scs if s.scenario_id == "S-K115")
    assert k115.config_overrides == {"head_k_factor": 115}
    assert k115.floor == "EV_charging"


def test_maxq_can_be_disabled():
    dv = _discretionary()
    scs = P.generate_calculation_scenarios(discretionary=dv, has_max_q_zone=False)
    assert all(s.scenario_id != "S-MAX-Q" for s in scs)


def test_scenario_total_never_exceeds_twelve():
    dv = _discretionary()
    scs = P.generate_calculation_scenarios(
        discretionary=dv, has_k115_zones=True, has_max_q_zone=True)
    assert len(scs) <= 12


def test_scenario_ids_are_unique():
    dv = _discretionary()
    scs = P.generate_calculation_scenarios(discretionary=dv, has_k115_zones=True)
    ids = [s.scenario_id for s in scs]
    assert len(ids) == len(set(ids))


# ────────────────────────────────────────────────────────────────────────────
# 5. 3대 불균형 지표
# ────────────────────────────────────────────────────────────────────────────

def test_delta_p_is_per_zone_span():
    dp = P.calc_pressure_imbalance({"hsp": [0.35, 0.9, 0.5], "lsp": [0.2, 0.25]})
    assert dp == {"hsp": 0.55, "lsp": 0.05}


def test_empty_zone_pressure_is_zero():
    assert P.calc_pressure_imbalance({"hsp": []}) == {"hsp": 0.0}


def test_cv_is_zero_for_identical_flows():
    assert P.calc_flow_cv([80.0] * 30) == 0.0


def test_cv_grows_with_spread():
    assert P.calc_flow_cv([60.0, 100.0]) > P.calc_flow_cv([79.0, 81.0])


def test_water_duration_uses_80pct_effective_volume():
    assert P.calc_water_duration(
        tank_total_volume_m3=30.0, total_actual_flow_lpm=1200.0) == 20.0


def test_zero_flow_gives_infinite_duration_not_zero():
    assert P.calc_water_duration(
        tank_total_volume_m3=30.0, total_actual_flow_lpm=0.0) == float("inf")


@pytest.mark.parametrize("dp, expected", [
    (0.6, "auto_pass"),          # 경계 포함
    (0.61, "human_review"),
    (0.9, "human_review"),       # 경계 포함
    (0.91, "redesign_required"),
])
def test_delta_p_tier_boundaries(dp, expected):
    tier, _ = P.evaluate_imbalance_tier(
        delta_p_max_mpa=dp, cv_flow=0.0, tau_water_minutes=30.0)
    assert tier == expected


@pytest.mark.parametrize("cv, expected", [
    (0.10, "auto_pass"),
    (0.11, "human_review"),
    (0.20, "human_review"),
    (0.21, "redesign_required"),
])
def test_cv_tier_boundaries(cv, expected):
    tier, _ = P.evaluate_imbalance_tier(
        delta_p_max_mpa=0.0, cv_flow=cv, tau_water_minutes=30.0)
    assert tier == expected


@pytest.mark.parametrize("tau, expected", [
    (22.0, "auto_pass"),          # legal × 1.10 경계 포함
    (21.9, "human_review"),
    (20.0, "human_review"),
    (19.9, "redesign_required"),
])
def test_tau_tier_boundaries(tau, expected):
    tier, _ = P.evaluate_imbalance_tier(
        delta_p_max_mpa=0.0, cv_flow=0.0, tau_water_minutes=tau)
    assert tier == expected


def test_worst_of_three_metrics_wins():
    """둘이 통과라도 하나가 재설계면 전체는 재설계 — 평균으로 물타면 안 된다."""
    tier, _ = P.evaluate_imbalance_tier(
        delta_p_max_mpa=0.0, cv_flow=0.0, tau_water_minutes=10.0)
    assert tier == "redesign_required"


def test_all_pass_yields_single_pass_message():
    _, msgs = P.evaluate_imbalance_tier(
        delta_p_max_mpa=0.1, cv_flow=0.01, tau_water_minutes=30.0)
    assert len(msgs) == 1
    assert msgs[0].startswith("PASS")


def test_each_failing_metric_adds_its_own_diagnosis():
    _, msgs = P.evaluate_imbalance_tier(
        delta_p_max_mpa=1.2, cv_flow=0.3, tau_water_minutes=10.0)
    assert len(msgs) == 3


def test_evaluate_imbalance_bundles_metrics():
    m = P.evaluate_imbalance(
        head_flows_lpm=[80.0] * 30,
        zone_pressures={"hsp": [0.35, 0.40]},
        tank_total_volume_m3=80.0)
    assert m.tier == "auto_pass"
    assert m.cv_flow == 0.0
    assert m.legal_duration_minutes == 20.0
    assert m.duration_reduction_pct < 0     # 법정보다 오래 버티면 감소율은 음수


def test_evaluate_imbalance_flags_overpressure_spread():
    m = P.evaluate_imbalance(
        head_flows_lpm=[80.0] * 30,
        zone_pressures={"lsp": [0.2, 1.5]},
        tank_total_volume_m3=80.0)
    assert m.tier == "redesign_required"
    assert m.delta_p_max_mpa_per_zone["lsp"] == 1.3


def test_empty_measurement_is_not_silently_auto_pass():
    """근거 0건은 '편차 없음'이 아니다.

    현재 evaluate_imbalance 는 빈 입력에서 CV 0.0 · τ ∞ 로 auto_pass 를 낸다.
    이 조용한 합격은 호출부(pipeline_orchestrator._insufficient_evidence_metrics)
    에서 차단하고 있어 여기서는 '모듈 단독 호출 시 방어가 없다'는 현재 계약을
    고정만 해 둔다. 모듈 안으로 가드를 옮기면 이 테스트를 뒤집어야 한다.
    """
    m = P.evaluate_imbalance(
        head_flows_lpm=[], zone_pressures={}, tank_total_volume_m3=0.0)
    assert m.tier == "auto_pass"
    assert m.tau_water_minutes == float("inf")


# ────────────────────────────────────────────────────────────────────────────
# 6. 대안 생성 · 순위
# ────────────────────────────────────────────────────────────────────────────

def _diagnosis(spread=1.5):
    return P.evaluate_imbalance(
        head_flows_lpm=[80.0] * 30,
        zone_pressures={"lsp": [0.2, spread]},
        tank_total_volume_m3=80.0)


def test_five_alternatives_are_generated():
    alts = P.generate_alternative_scenarios(
        diagnosis=_diagnosis(), hb_case=decide_hb_case(building_height_m=200.0))
    assert [a.alt_id for a in alts] == [
        "ALT-1-PRV", "ALT-2-LOOP", "ALT-3-MID-TANK",
        "ALT-4-BASEMENT-TANK", "ALT-5-FLOW-CONTROL"]


def test_structural_alternatives_require_human_review():
    """중간수조·지하수조는 건축 협의가 필요하다 — 자동 확정 대상이 아니다."""
    alts = P.generate_alternative_scenarios(
        diagnosis=_diagnosis(), hb_case=decide_hb_case(building_height_m=200.0))
    review = {a.alt_id for a in alts if a.requires_human_review}
    assert review == {"ALT-3-MID-TANK", "ALT-4-BASEMENT-TANK"}


def test_cheaper_alternatives_have_no_space_impact():
    alts = P.generate_alternative_scenarios(
        diagnosis=_diagnosis(), hb_case=decide_hb_case(building_height_m=200.0))
    for a in alts:
        if a.estimated_material_cost_pct <= 8.0:
            assert a.space_impact == "없음"


def test_unsimulated_alternatives_are_no_sim_not_pass():
    """시뮬레이션을 돌리지 않은 대안이 점수 0.0 으로 1위가 되면 안 된다."""
    alts = P.generate_alternative_scenarios(
        diagnosis=_diagnosis(), hb_case=decide_hb_case(building_height_m=200.0))
    rows = P.rank_alternatives(alts, simulation_results={})
    assert {r["verdict"] for r in rows} == {"NO-SIM"}
    assert all("recommendation" not in r for r in rows)


def test_simulated_pass_outranks_unsimulated():
    alts = P.generate_alternative_scenarios(
        diagnosis=_diagnosis(), hb_case=decide_hb_case(building_height_m=200.0))
    rows = P.rank_alternatives(alts, simulation_results={"ALT-2-LOOP": _diagnosis(0.25)})
    assert rows[0]["alt_id"] == "ALT-2-LOOP"
    assert rows[0]["verdict"] == "PASS"
    assert rows[0]["recommendation"] == "★ 자동 추천"


def test_ranking_prefers_lower_score_within_same_tier():
    alts = P.generate_alternative_scenarios(
        diagnosis=_diagnosis(), hb_case=decide_hb_case(building_height_m=200.0))
    sims = {"ALT-1-PRV": _diagnosis(0.25), "ALT-4-BASEMENT-TANK": _diagnosis(0.25)}
    rows = P.rank_alternatives(alts, simulation_results=sims)
    assert rows[0]["alt_id"] == "ALT-1-PRV"      # 동일 tier → 저비용 우선
    assert rows[0]["score"] < rows[1]["score"]


def test_failing_simulation_sinks_below_review():
    alts = P.generate_alternative_scenarios(
        diagnosis=_diagnosis(), hb_case=decide_hb_case(building_height_m=200.0))
    sims = {"ALT-1-PRV": _diagnosis(1.5), "ALT-2-LOOP": _diagnosis(0.85)}
    rows = P.rank_alternatives(alts, simulation_results=sims)
    verdicts = [r["verdict"] for r in rows if r["verdict"] != "NO-SIM"]
    assert verdicts == ["REVIEW", "FAIL"]


# ────────────────────────────────────────────────────────────────────────────
# 공개 API
# ────────────────────────────────────────────────────────────────────────────

def test_all_exports_exist():
    for name in P.__all__:
        assert hasattr(P, name), name
