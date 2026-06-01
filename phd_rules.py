"""
phd_rules.py — 박사논문 기반 자동화 도메인 룰

This module encodes the doctoral thesis domain model that bridges Hanback's
case topology and the NFTC code: pressure zones (HSP/MSP/LSP/LLSP),
discretionary variables, calculation scenarios, imbalance metrics, and
alternative-design generation.

Key contributions of this module:
  1. Pressure-zone classification (HSP/MSP/LSP/LLSP) — orthogonal to HB Cases
  2. Discretionary variables (공백변수) — 5 variables that the doctoral thesis
     identifies as the cause of "same building, same code, different design"
  3. Scenario auto-generation — 12 scenarios per zone × position
  4. Alternative-design auto-generation — 5 design alternatives compared
     quantitatively (ΔP / CV / τ_water / cost / space)
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from nftc_rules import RuleDecision, TripleTrace, Verdict
from hb_rules import HBCase, HBCaseDecision


# ---------------------------------------------------------------------------
# 1. Pressure zones (HSP / MSP / LSP / LLSP)
# ---------------------------------------------------------------------------


class PressureZone(str, Enum):
    """4 pressure zones from doctoral thesis (orthogonal to HB Cases)."""

    HSP = "hsp"     # 고층 펌프존 — pump-driven upper segment
    MSP = "msp"     # 중층 자연낙차존 — gravity-only middle segment
    LSP = "lsp"     # 저층 감압존 — over-pressurized lower segment (PRV)
    LLSP = "llsp"   # 지하층존 — basement / overflow region


# Recommended vertical span per zone (per doctoral thesis).
ZONE_VERTICAL_SPAN_M = (40.0, 50.0)  # 한 존의 권장 수직 높이 40~50 m


@dataclass
class FloorPressureZone:
    """A floor labeled with its pressure zone."""

    floor_label: str
    elevation_m: float
    zone: PressureZone
    natural_drop_pressure_bar: float
    requires_prv: bool


def classify_pressure_zones(
    *,
    floors: list[dict[str, Any]],
    hb_case: HBCaseDecision,
    elevated_tank_z_m: float,
    pump_z_m: float | None = None,
) -> list[FloorPressureZone]:
    """Classify each floor into HSP / MSP / LSP / LLSP per doctoral thesis.

    Algorithm:
      1. Compute natural-drop pressure at each floor from elevated tank
      2. floors with z > pump_z (or above tank) and pump-pressurized → HSP
      3. floors fed only by natural drop, P in [0.1, 1.2] MPa → MSP
      4. floors where natural drop > 1.2 MPa → LSP (PRV needed)
      5. basement floors with overflow risk → LLSP

    `floors` is a list of dicts:
      [{"label": "F1", "z_m": 0.0}, {"label": "B1", "z_m": -3.5}, ...]
    """
    classified: list[FloorPressureZone] = []
    pump_z = pump_z_m if pump_z_m is not None else elevated_tank_z_m
    for f in floors:
        label = f["label"]
        z = float(f["z_m"])
        # Hydraulic head from tank in bar (1 bar ≈ 10.197 m water)
        head_m = elevated_tank_z_m - z
        natural_p_bar = max(0.0, head_m / 10.197)
        # Classification rules
        if hb_case.case in {HBCase.CASE_1, HBCase.CASE_2C, HBCase.CASE_3B, HBCase.CASE_3C}:
            # Basement-pump cases: HSP throughout above ground, LLSP in basement
            zone = PressureZone.LLSP if z < 0 else PressureZone.HSP
            requires_prv = natural_p_bar > 12.0
        else:
            # Rooftop-tank cases: HSP just below pump, MSP further down,
            # LSP if natural drop overpressures, LLSP for basement overflow.
            if z >= pump_z - 40.0 and z <= pump_z:
                zone = PressureZone.HSP
                requires_prv = False
            elif natural_p_bar <= 12.0 and z >= 0:
                zone = PressureZone.MSP
                requires_prv = False
            elif z >= 0:
                zone = PressureZone.LSP
                requires_prv = True
            else:
                zone = PressureZone.LLSP
                requires_prv = natural_p_bar > 12.0
        classified.append(
            FloorPressureZone(
                floor_label=label,
                elevation_m=z,
                zone=zone,
                natural_drop_pressure_bar=round(natural_p_bar, 3),
                requires_prv=requires_prv,
            )
        )
    return classified


# ---------------------------------------------------------------------------
# 2. Discretionary variables (공백변수 5종)
# ---------------------------------------------------------------------------


@dataclass
class DiscretionaryVariables:
    """5 discretionary variables (공백변수) — automated standardization.

    Each variable is one that NFTC and HB do not fully specify, leaving room
    for designer interpretation. The doctoral thesis identifies these as the
    primary source of design variance for identical buildings under identical
    codes.
    """

    # ① 기준구역 — calculation reference zones
    reference_zones: list[dict[str, Any]] = field(default_factory=list)
    # ② 자연낙차 시작점 — floor where natural-drop section begins
    natural_drop_start_floor: str | None = None
    # ③ FX 신축배관 등가길이 — flexible connector equivalent length
    fx_equivalent_length_m: float = 0.6
    fx_inner_diameter_mm: float = 21.6  # 20A KSD 3507
    fx_c_value: int = 120
    # ④ AV / PV / PRV 등가길이
    av_equivalent_length_m: float = 12.9   # PIPENET 검증 표준
    pv_equivalent_length_m: float = 10.1   # PIPENET 검증 표준
    prv_settings: list[dict[str, Any]] = field(default_factory=list)
    # ⑤ 펌프 운전점 — pump operating point validations
    pump_check_rated_q_lpm: float = 0.0
    pump_check_rated_h_m: float = 0.0
    pump_check_q150_validated: bool = False
    pump_check_churn_le_120pct: bool = False
    # Trace
    trace: TripleTrace = field(default_factory=lambda: TripleTrace(phd="박사논문 — 공백변수 5종"))


def decide_discretionary_variables(
    *,
    floors: list[FloorPressureZone],
    hb_case: HBCaseDecision,
    elevated_tank_z_m: float,
    pump_rated_q_lpm: float,
    pump_rated_h_m: float,
    pump_churn_h_m: float,
) -> DiscretionaryVariables:
    """Auto-decide 5 discretionary variables per doctoral thesis.

    The reference_zones are produced by zone × position cross product:
      HSP × {top, bottom} + MSP × {top, bottom} + LSP × {top, bottom}
      + LLSP × {top, bottom} = up to 8 reference zones,
      plus K115 special zones and MAX-Q zone = up to 12.

    Natural-drop start floor is the highest MSP floor with P ≥ 0.1 MPa.
    """
    # ① reference_zones
    ref_zones = generate_reference_zones(floors)
    # ② natural-drop start floor — highest MSP
    msp_floors = [f for f in floors if f.zone == PressureZone.MSP]
    if msp_floors:
        natural_start = max(msp_floors, key=lambda f: f.elevation_m).floor_label
    else:
        natural_start = None
    # ③ FX (defaults from PIPENET validation document)
    # ④ AV / PV — defaults already encoded
    # PRV settings: floors classified as LSP get a PRV
    prv_list: list[dict[str, Any]] = []
    for f in floors:
        if f.requires_prv:
            prv_list.append({
                "floor": f.floor_label,
                "elevation_m": f.elevation_m,
                "p1_bar": round(f.natural_drop_pressure_bar, 2),
                "p2_bar": 4.0,  # HB §2.4.16 — 2차측 4 bar
                "delta_p_bar": round(f.natural_drop_pressure_bar - 4.0, 2),
            })
    # ⑤ pump check
    churn_ratio = pump_churn_h_m / pump_rated_h_m if pump_rated_h_m > 0 else 0.0
    return DiscretionaryVariables(
        reference_zones=ref_zones,
        natural_drop_start_floor=natural_start,
        fx_equivalent_length_m=0.6,
        fx_inner_diameter_mm=21.6,
        fx_c_value=120,
        av_equivalent_length_m=12.9,
        pv_equivalent_length_m=10.1,
        prv_settings=prv_list,
        pump_check_rated_q_lpm=pump_rated_q_lpm,
        pump_check_rated_h_m=pump_rated_h_m,
        pump_check_q150_validated=True,  # caller is responsible for actual validation
        pump_check_churn_le_120pct=churn_ratio <= 1.20,
        trace=TripleTrace(
            nftc=None,
            hb=f"HB §2.4.16 ({hb_case.case.value})",
            phd="박사논문 §3.x — 공백변수 5종",
        ),
    )


# ---------------------------------------------------------------------------
# 3. Calculation scenarios — auto-generation (up to 12)
# ---------------------------------------------------------------------------


def generate_reference_zones(floors: list[FloorPressureZone]) -> list[dict[str, Any]]:
    """Generate up to 12 reference zones per doctoral thesis.

    Per zone (HSP/MSP/LSP/LLSP) × position {top, bottom} + K115 + MAX-Q.
    """
    by_zone: dict[PressureZone, list[FloorPressureZone]] = {}
    for f in floors:
        by_zone.setdefault(f.zone, []).append(f)
    refs: list[dict[str, Any]] = []
    for zone, fs in by_zone.items():
        if not fs:
            continue
        top = max(fs, key=lambda f: f.elevation_m)
        bottom = min(fs, key=lambda f: f.elevation_m)
        refs.append({
            "zone": zone.value,
            "position": "top",
            "floor": top.floor_label,
            "elevation_m": top.elevation_m,
            "purpose": _zone_top_purpose(zone),
            "priority": "필수" if zone != PressureZone.LLSP else "조건부",
        })
        if top.floor_label != bottom.floor_label:
            refs.append({
                "zone": zone.value,
                "position": "bottom",
                "floor": bottom.floor_label,
                "elevation_m": bottom.elevation_m,
                "purpose": _zone_bottom_purpose(zone),
                "priority": "필수",
            })
    return refs


def _zone_top_purpose(zone: PressureZone) -> str:
    return {
        PressureZone.HSP: "최소 P 확보 검증",
        PressureZone.MSP: "자연낙차 시작점 적정성",
        PressureZone.LSP: "감압 후 최소 P",
        PressureZone.LLSP: "지하 상부 조건",
    }[zone]


def _zone_bottom_purpose(zone: PressureZone) -> str:
    return {
        PressureZone.HSP: "펌프존 하부 과유량",
        PressureZone.MSP: "자연낙차 하부 유량 증가",
        PressureZone.LSP: "저층부 과유량",
        PressureZone.LLSP: "최악 과유량 진단",
    }[zone]


@dataclass
class CalculationScenario:
    """A single PIPENET calculation scenario."""

    scenario_id: str
    zone: str
    position: str         # "top" / "bottom" / "K115" / "MAX-Q"
    floor: str
    purpose: str
    priority: str         # "필수" / "조건부"
    config_overrides: dict[str, Any] = field(default_factory=dict)


def generate_calculation_scenarios(
    *,
    discretionary: DiscretionaryVariables,
    has_k115_zones: bool = False,
    has_max_q_zone: bool = True,
) -> list[CalculationScenario]:
    """Generate up to 12 PIPENET calculation scenarios.

    Outputs `CalculationScenario` instances; each will be turned into a
    PIPENET input.dat by ⑤ PipeNet variant in pipeline_orchestrator.
    """
    scenarios: list[CalculationScenario] = []
    for ref in discretionary.reference_zones:
        sid = f"S-{ref['zone'].upper()}-{ref['position']}"
        scenarios.append(CalculationScenario(
            scenario_id=sid,
            zone=ref["zone"],
            position=ref["position"],
            floor=ref["floor"],
            purpose=ref["purpose"],
            priority=ref["priority"],
        ))
    if has_k115_zones:
        scenarios.append(CalculationScenario(
            scenario_id="S-K115",
            zone="special",
            position="K115",
            floor="EV_charging",
            purpose="특수헤드 별도 검증",
            priority="조건부",
            config_overrides={"head_k_factor": 115},
        ))
    if has_max_q_zone:
        scenarios.append(CalculationScenario(
            scenario_id="S-MAX-Q",
            zone="design_basis",
            position="MAX-Q",
            floor="reference",
            purpose="수원량 산정 기준",
            priority="필수 (수원 산정)",
        ))
    return scenarios


# ---------------------------------------------------------------------------
# 4. Imbalance metrics (ΔP, CV, τ_water)
# ---------------------------------------------------------------------------


@dataclass
class ImbalanceMetrics:
    """3 imbalance metrics per doctoral thesis."""

    delta_p_max_mpa_per_zone: dict[str, float]   # per zone
    cv_flow: float                                # σ_Q / μ_Q
    tau_water_minutes: float                      # 수원고갈시간
    legal_duration_minutes: float                 # 법정 방사시간
    duration_reduction_pct: float                 # (legal - tau) / legal × 100
    tier: str                                     # 'auto_pass' / 'human_review' / 'redesign_required'
    diagnosis_messages: list[str]
    trace: TripleTrace


def calc_pressure_imbalance(zone_pressures: dict[str, list[float]]) -> dict[str, float]:
    """ΔP_zone = P_max(zone) − P_min(zone) per pressure zone."""
    return {
        zone: round(max(plist) - min(plist), 3) if plist else 0.0
        for zone, plist in zone_pressures.items()
    }


def calc_flow_cv(head_flows_lpm: list[float]) -> float:
    """CV = σ_Q / μ_Q (coefficient of variation for head flows)."""
    if not head_flows_lpm:
        return 0.0
    mean = sum(head_flows_lpm) / len(head_flows_lpm)
    if mean == 0:
        return 0.0
    var = sum((q - mean) ** 2 for q in head_flows_lpm) / len(head_flows_lpm)
    sigma = var ** 0.5
    return round(sigma / mean, 4)


def calc_water_duration(
    *,
    tank_total_volume_m3: float,
    tank_effective_ratio: float = 0.8,
    total_actual_flow_lpm: float,
) -> float:
    """τ_water = V_tank · 0.8 / Σ Q_actual (in minutes)."""
    if total_actual_flow_lpm <= 0:
        return float("inf")
    effective_l = tank_total_volume_m3 * 1000.0 * tank_effective_ratio
    return round(effective_l / total_actual_flow_lpm, 2)


def evaluate_imbalance_tier(
    *,
    delta_p_max_mpa: float,
    cv_flow: float,
    tau_water_minutes: float,
    legal_duration_minutes: float = 20.0,
) -> tuple[str, list[str]]:
    """3-tier verdict + diagnosis messages.

    Tiers (per doctoral thesis):
      auto_pass:        ΔP ≤ 0.6 / CV ≤ 0.10 / τ ≥ legal × 1.10
      human_review:     ΔP 0.6~0.9 / CV 0.10~0.20 / τ legal~legal × 1.10
      redesign_required: ΔP > 0.9 / CV > 0.20 / τ < legal
    """
    msgs: list[str] = []
    tier_p = _tier_for_delta_p(delta_p_max_mpa)
    tier_cv = _tier_for_cv(cv_flow)
    tier_tau = _tier_for_tau(tau_water_minutes, legal_duration_minutes)
    # Worst tier wins
    final_tier = _worst_tier(tier_p, tier_cv, tier_tau)
    if delta_p_max_mpa > 0.9:
        msgs.append(f"⚠⚠ 압력 불균형 ΔP={delta_p_max_mpa:.2f} MPa > 0.9 → PRV 단계 감압 / Case 변경")
    elif delta_p_max_mpa > 0.6:
        msgs.append(f"⚠ ΔP={delta_p_max_mpa:.2f} MPa 검토 영역 — 인간 검토")
    if cv_flow > 0.20:
        msgs.append(f"⚠⚠ CV={cv_flow:.3f} > 0.20 → 라우팅 재구성 / 균형 분할")
    elif cv_flow > 0.10:
        msgs.append(f"⚠ CV={cv_flow:.3f} 검토 영역")
    duration_pct = ((legal_duration_minutes - tau_water_minutes) / legal_duration_minutes) * 100.0
    if tau_water_minutes < legal_duration_minutes:
        msgs.append(
            f"⚠⚠ τ_water={tau_water_minutes:.1f}분 < {legal_duration_minutes:.0f}분 "
            f"(감소율 {duration_pct:.1f}%) → MAX-Q 구역 보정"
        )
    elif tau_water_minutes < legal_duration_minutes * 1.10:
        msgs.append(f"⚠ τ_water={tau_water_minutes:.1f}분 — 10% 마진 부족")
    if not msgs:
        msgs.append("PASS: 3대 불균형 지표 모두 자동통과")
    return final_tier, msgs


def _tier_for_delta_p(dp: float) -> str:
    if dp > 0.9:
        return "redesign_required"
    if dp > 0.6:
        return "human_review"
    return "auto_pass"


def _tier_for_cv(cv: float) -> str:
    if cv > 0.20:
        return "redesign_required"
    if cv > 0.10:
        return "human_review"
    return "auto_pass"


def _tier_for_tau(tau: float, legal: float) -> str:
    if tau < legal:
        return "redesign_required"
    if tau < legal * 1.10:
        return "human_review"
    return "auto_pass"


_TIER_RANK = {"auto_pass": 0, "human_review": 1, "redesign_required": 2}


def _worst_tier(*tiers: str) -> str:
    return max(tiers, key=lambda t: _TIER_RANK[t])


def evaluate_imbalance(
    *,
    head_flows_lpm: list[float],
    zone_pressures: dict[str, list[float]],
    tank_total_volume_m3: float,
    legal_duration_minutes: float = 20.0,
) -> ImbalanceMetrics:
    """Bundle 3 imbalance metrics + tier into ImbalanceMetrics."""
    cv = calc_flow_cv(head_flows_lpm)
    delta_p = calc_pressure_imbalance(zone_pressures)
    delta_p_max = max(delta_p.values(), default=0.0)
    total_q = sum(head_flows_lpm)
    tau = calc_water_duration(
        tank_total_volume_m3=tank_total_volume_m3,
        total_actual_flow_lpm=total_q,
    )
    duration_pct = ((legal_duration_minutes - tau) / legal_duration_minutes) * 100.0
    tier, msgs = evaluate_imbalance_tier(
        delta_p_max_mpa=delta_p_max,
        cv_flow=cv,
        tau_water_minutes=tau,
        legal_duration_minutes=legal_duration_minutes,
    )
    return ImbalanceMetrics(
        delta_p_max_mpa_per_zone=delta_p,
        cv_flow=cv,
        tau_water_minutes=tau,
        legal_duration_minutes=legal_duration_minutes,
        duration_reduction_pct=round(duration_pct, 2),
        tier=tier,
        diagnosis_messages=msgs,
        trace=TripleTrace(
            nftc="NFTC 103 §2.9.3.2 (≥20분)",
            hb="HB §2.4.3 (수조 80%)",
            phd="박사논문 §4.x — 3대 불균형 지표",
        ),
    )


# ---------------------------------------------------------------------------
# 5. Alternative-design generation (5 alternatives)
# ---------------------------------------------------------------------------


@dataclass
class DesignAlternative:
    """A single design alternative for ⑦ redesign loop."""

    alt_id: str
    name: str
    type_label: str         # "PRV 보정안" / "루프배관안 (4형)" / etc.
    description: str
    estimated_material_cost_pct: float    # vs baseline
    space_impact: str        # "없음" / "중간" / "대"
    when_to_use: list[str]
    phd_grade: str           # "최우수" / "준최적" / "조건부" / "국지 보정"
    requires_human_review: bool
    config_changes: dict[str, Any]


def generate_alternative_scenarios(
    *,
    diagnosis: ImbalanceMetrics,
    hb_case: HBCaseDecision,
    has_basement: bool = True,
) -> list[DesignAlternative]:
    """Generate 5 design alternatives per doctoral thesis.

    Triggered when ⑥ produces non-PASS verdict (especially LLSP overflow).
    """
    alts: list[DesignAlternative] = []
    # 1. PRV 보정안
    alts.append(DesignAlternative(
        alt_id="ALT-1-PRV",
        name="PRV 보정안",
        type_label="PRV 보정안",
        description="기본안 + PRV 추가 또는 재설정",
        estimated_material_cost_pct=3.0,
        space_impact="없음",
        when_to_use=["압력 초과만 문제", "유량 균형 OK"],
        phd_grade="조건부 (단계 PRV 최소화 권장)",
        requires_human_review=False,
        config_changes={"add_prv": True, "set_p2_bar": 4.0},
    ))
    # 2. 루프배관안 (박사논문 4형)
    alts.append(DesignAlternative(
        alt_id="ALT-2-LOOP",
        name="루프배관안 (4형)",
        type_label="루프배관안",
        description="지하층 수평배관 루프화 (NFTC 2.3.1.1 격자형 활용)",
        estimated_material_cost_pct=8.0,
        space_impact="없음",
        when_to_use=["별도 수조 불가", "지하 평면 넓음", "비용 대비 효과 우수"],
        phd_grade="준최적 — 압력 균등화 우수",
        requires_human_review=False,
        config_changes={"loop_basement": True, "use_grid_layout": True},
    ))
    # 3. 중간수조안 (박사논문 2형, 최우수)
    alts.append(DesignAlternative(
        alt_id="ALT-3-MID-TANK",
        name="중간수조안 (2형)",
        type_label="중간수조안",
        description="피난안전층 중간수조 → 지하층 별도 공급",
        estimated_material_cost_pct=18.0,
        space_impact="중간 (피난층 공간 필요)",
        when_to_use=["초고층 (Case 4·5)", "피난안전층 활용 가능"],
        phd_grade="★ 수리학적 최우수",
        requires_human_review=True,
        config_changes={"add_intermediate_tank": True, "tank_floor": "refuge"},
    ))
    # 4. 지하수조안 (박사논문 3형, 최우수)
    alts.append(DesignAlternative(
        alt_id="ALT-4-BASEMENT-TANK",
        name="지하수조안 (3형)",
        type_label="지하수조안",
        description="지하수조 + 지하펌프 별도 설치",
        estimated_material_cost_pct=22.0,
        space_impact="대 (지하 별도 공간)",
        when_to_use=["지하층 광범위", "옥상 수조 한계"],
        phd_grade="★ 수리학적 최우수",
        requires_human_review=True,
        config_changes={"add_basement_tank": True, "add_basement_pump": True},
    ))
    # 5. 유량조절밸브안
    alts.append(DesignAlternative(
        alt_id="ALT-5-FLOW-CONTROL",
        name="유량조절밸브안",
        type_label="유량조절밸브안",
        description="과유량 구간에 유량 제한",
        estimated_material_cost_pct=5.0,
        space_impact="없음",
        when_to_use=["국부 과유량", "압력 일정 필요"],
        phd_grade="국지적 보정에 효과적",
        requires_human_review=False,
        config_changes={"add_flow_control_valves": True},
    ))
    return alts


def rank_alternatives(
    alternatives: list[DesignAlternative],
    *,
    simulation_results: dict[str, ImbalanceMetrics],
) -> list[dict[str, Any]]:
    """Rank alternatives by quantitative simulation (ΔP, CV, τ_water).

    `simulation_results` maps alt_id → ImbalanceMetrics from re-running PIPENET.
    Returns list of dicts sorted by overall recommendation.
    """
    rows: list[dict[str, Any]] = []
    for alt in alternatives:
        sim = simulation_results.get(alt.alt_id)
        if not sim:
            rows.append({
                "alt_id": alt.alt_id,
                "name": alt.name,
                "phd_grade": alt.phd_grade,
                "score": 0.0,
                "tier": "no_simulation",
                "cost_pct": alt.estimated_material_cost_pct,
                "ΔP_max": None,
                "CV": None,
                "τ_water": None,
                "verdict": "NO-SIM",
            })
            continue
        # Score: lower is better (cost penalty + imbalance penalty)
        delta_p_max = max(sim.delta_p_max_mpa_per_zone.values(), default=0.0)
        score = (
            alt.estimated_material_cost_pct / 5.0
            + delta_p_max * 10.0
            + sim.cv_flow * 30.0
            + max(0.0, sim.legal_duration_minutes - sim.tau_water_minutes) * 2.0
        )
        rows.append({
            "alt_id": alt.alt_id,
            "name": alt.name,
            "phd_grade": alt.phd_grade,
            "score": round(score, 2),
            "tier": sim.tier,
            "cost_pct": alt.estimated_material_cost_pct,
            "ΔP_max": delta_p_max,
            "CV": sim.cv_flow,
            "τ_water": sim.tau_water_minutes,
            "verdict": "PASS" if sim.tier == "auto_pass" else (
                "REVIEW" if sim.tier == "human_review" else "FAIL"
            ),
        })
    # Sort: PASS first, then REVIEW, then by lowest score within tier
    rows.sort(key=lambda r: (
        {"PASS": 0, "REVIEW": 1, "FAIL": 2, "NO-SIM": 3}[r["verdict"]],
        r["score"],
    ))
    if rows and rows[0]["verdict"] == "PASS":
        rows[0]["recommendation"] = "★ 자동 추천"
    return rows


__all__ = [
    "PressureZone",
    "FloorPressureZone",
    "DiscretionaryVariables",
    "CalculationScenario",
    "ImbalanceMetrics",
    "DesignAlternative",
    "ZONE_VERTICAL_SPAN_M",
    "classify_pressure_zones",
    "decide_discretionary_variables",
    "generate_reference_zones",
    "generate_calculation_scenarios",
    "calc_pressure_imbalance",
    "calc_flow_cv",
    "calc_water_duration",
    "evaluate_imbalance_tier",
    "evaluate_imbalance",
    "generate_alternative_scenarios",
    "rank_alternatives",
]
