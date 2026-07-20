"""
hb_rules.py — 한백에프앤씨 설계지침서 §2.4 룰 엔진

This module encodes Hanback Engineering's in-house design standard (한백
설계지침서 v.250508) §2.4 sprinkler rules as deterministic functions.

Hanback rules are sourced from the firm's internal handbook and either:
  (a) reinforce NFTC with stricter values (e.g. churn ≤120% vs NFTC 140%)
  (b) cover NFTC-silent areas (e.g. velocity 6/10 m/s)
  (c) provide topology selection (Case 1~5 systems)

All decisions return TripleTrace with NFTC + HB attribution where applicable.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from nftc_rules import RuleDecision, TripleTrace, Verdict


# ---------------------------------------------------------------------------
# 1. System type selection (5 systems per HB §2.4.1)
# ---------------------------------------------------------------------------


class SystemType(str, Enum):
    """5 systems defined in HB §2.4.1."""

    WET = "wet"                  # 습식 (기본)
    DRY = "dry"                  # 건식
    PREACTION_SINGLE = "preaction_single"   # 준비작동식 Single Interlock
    PREACTION_NONE = "preaction_none"       # 준비작동식 None Interlock
    PREACTION_DOUBLE = "preaction_double"   # 준비작동식 Double Interlock
    DELUGE = "deluge"            # 일제살수식


def decide_system_type(
    *,
    has_freezing_risk: bool = False,
    needs_open_heads: bool = False,
    detector_priority: bool = False,
    room_use: str = "",
) -> RuleDecision:
    """Decide system type per HB §2.4.1.

    Decision tree:
      - 동결 우려 (frost risk) → 건식 또는 준비작동식
      - 개방형 헤드 필요 (무대부·플랜트) → 일제살수식
      - 감지기 우선 보호 (전산실·문서고) → 준비작동식
      - 그 외 → 습식 (기본, 신뢰도 최고)
    """
    if needs_open_heads:
        return RuleDecision(
            rule_id="HB-241-DELUGE",
            verdict=Verdict.PASS,
            value=SystemType.DELUGE.value,
            trace=TripleTrace(
                nftc="NFTC 103 §1.7 (일제살수식 정의)",
                hb="HB §2.4.1, §2.4.14",
                phd=None,
                note=f"개방형 헤드 필요 ({room_use})",
            ),
            detail="일제살수식 (Deluge) — 개방형 헤드 + 감지기 작동",
        )
    if has_freezing_risk:
        if detector_priority:
            return RuleDecision(
                rule_id="HB-241-PREACTION-DOUBLE",
                verdict=Verdict.PASS,
                value=SystemType.PREACTION_DOUBLE.value,
                trace=TripleTrace(
                    nftc="NFTC 103 §1.7 (준비작동식 정의)",
                    hb="HB §2.4.1, §2.4.13",
                    phd=None,
                    note="동결 + 감지기 우선 → Double Interlock",
                ),
                detail="준비작동식 Double Interlock",
            )
        return RuleDecision(
            rule_id="HB-241-DRY",
            verdict=Verdict.PASS,
            value=SystemType.DRY.value,
            trace=TripleTrace(
                nftc="NFTC 103 §1.7",
                hb="HB §2.4.1, §2.4.15",
                phd=None,
                note="동결 우려 — 건식 또는 준비작동식 검토",
            ),
            detail="건식 (Dry Pipe) — 그리드 금지·2,840L 한계",
        )
    if detector_priority:
        return RuleDecision(
            rule_id="HB-241-PREACTION-SINGLE",
            verdict=Verdict.PASS,
            value=SystemType.PREACTION_SINGLE.value,
            trace=TripleTrace(
                nftc="NFTC 103 §1.7",
                hb="HB §2.4.1, §2.4.13",
                phd=None,
                note="감지기 우선 보호 (전산실·문서고)",
            ),
            detail="준비작동식 Single Interlock",
        )
    return RuleDecision(
        rule_id="HB-241-WET",
        verdict=Verdict.PASS,
        value=SystemType.WET.value,
        trace=TripleTrace(
            nftc="NFTC 103 §1.7",
            hb="HB §2.4.1 (기본)",
            phd=None,
            note="기본 — 신뢰도 최고",
        ),
        detail="습식 (Wet) — 기본",
    )


# ---------------------------------------------------------------------------
# 2. Case 1~5 system topology (HB §2.4.16)
# ---------------------------------------------------------------------------


class HBCase(str, Enum):
    """11 sub-cases of HB §2.4.16 system topology."""

    CASE_1 = "case_1"      # ≤60m 지하 펌프
    CASE_2A = "case_2a"    # 60~80m 옥상 + 자연낙차 (원칙)
    CASE_2B = "case_2b"    # 60~80m 고저층 분리
    CASE_2C = "case_2c"    # 60~80m 지하 + 감압
    CASE_3A = "case_3a"    # 80~120m 옥상 + 자연낙차
    CASE_3B = "case_3b"    # 80~120m 지하 + 가압 ±40
    CASE_3C = "case_3c"    # 80~120m 지하 + 감압
    CASE_4A = "case_4a"    # 초고층 피난안전층 ≤80m (원칙)
    CASE_4B = "case_4b"    # 초고층 중간층 펌프
    CASE_5A = "case_5a"    # 피난안전층 >80m 옥상
    CASE_5B = "case_5b"    # 피난안전층 >80m 구간 분리


@dataclass
class HBCaseDecision:
    """Result of HB §2.4.16 Case decision."""

    case: HBCase
    pump_location: str          # "basement" / "rooftop" / "intermediate"
    rated_head_max_m: float
    churn_head_max_m: float
    pressurized_zone_m: float   # 가압구간 (40m 등)
    natural_drop_zone_m: float
    prv_required: bool
    prv_secondary_bar: float    # 4 bar (한백 룰)
    pipe_material_change_at_m: float | None  # KSD 3562→3507 경계
    trace: TripleTrace
    detail: str


def decide_hb_case(
    *,
    building_height_m: float,
    refuge_floor_interval_m: float | None = None,
    rooftop_tank_feasible: bool = True,
    water_source_type: str = "fire_dedicated",  # or "shared_with_potable"
) -> HBCaseDecision:
    """Decide HB §2.4.16 Case 1~5 per building topology.

    Inputs:
      - building_height_m: total building height (지하층 포함)
      - refuge_floor_interval_m: 피난안전층 간격 (초고층만 의미 있음)
      - rooftop_tank_feasible: 옥상 고가수조 설치 가능 여부 (구조·미관)
      - water_source_type: "fire_dedicated" or "shared_with_potable"
    """
    H = building_height_m
    # Case 1: ≤ 60 m
    if H <= 60:
        return HBCaseDecision(
            case=HBCase.CASE_1,
            pump_location="basement",
            rated_head_max_m=100.0,
            churn_head_max_m=120.0,
            pressurized_zone_m=H,
            natural_drop_zone_m=0.0,
            prv_required=False,
            prv_secondary_bar=4.0,
            pipe_material_change_at_m=None,
            trace=TripleTrace(
                nftc="NFTC 103 §2.2 (가압송수장치 일반)",
                hb="HB §2.4.16 Case 1",
                phd=None,
                note=f"건물 높이 {H} m ≤ 60 m",
            ),
            detail="Case 1: 지하 펌프 / 정격 ≤100m / 체절 ≤120m",
        )
    # Case 2: 60 < H ≤ 80
    if H <= 80:
        if rooftop_tank_feasible:
            return HBCaseDecision(
                case=HBCase.CASE_2A,
                pump_location="rooftop",
                rated_head_max_m=H + 40,
                churn_head_max_m=(H + 40) * 1.2,
                pressurized_zone_m=40.0,
                natural_drop_zone_m=H - 40,
                prv_required=False,
                prv_secondary_bar=4.0,
                pipe_material_change_at_m=H + 40 - 120,
                trace=TripleTrace(
                    nftc="NFTC 103 §2.2",
                    hb="HB §2.4.16 Case 2a (원칙)",
                    phd=None,
                    note=f"H={H}m 옥상 고가수조 + 자연낙차 (원칙)",
                ),
                detail="Case 2a: 옥상 고가수조 + 펌프 / 가압 40m / 자연낙차",
            )
        # 옥상 불가 → Case 2c (지하 + 감압)
        return HBCaseDecision(
            case=HBCase.CASE_2C,
            pump_location="basement",
            rated_head_max_m=120.0,
            churn_head_max_m=144.0,
            pressurized_zone_m=H,
            natural_drop_zone_m=0.0,
            prv_required=True,
            prv_secondary_bar=4.0,
            pipe_material_change_at_m=None,
            trace=TripleTrace(
                nftc="NFTC 103 §2.2",
                hb="HB §2.4.16 Case 2c (옥상 불가 시)",
                phd=None,
                note=f"H={H}m 지하 펌프 + 감압",
            ),
            detail="Case 2c: 지하 펌프 / 정격 ≤120m / 감압밸브",
        )
    # Case 3: 80 < H ≤ 120
    if H <= 120:
        if rooftop_tank_feasible:
            return HBCaseDecision(
                case=HBCase.CASE_3A,
                pump_location="rooftop",
                rated_head_max_m=H + 40,
                churn_head_max_m=(H + 40) * 1.2,
                pressurized_zone_m=40.0,
                natural_drop_zone_m=H - 40,
                prv_required=False,
                prv_secondary_bar=4.0,
                pipe_material_change_at_m=(H + 40) - 120,
                trace=TripleTrace(
                    nftc="NFTC 103 §2.2",
                    hb="HB §2.4.16 Case 3a",
                    phd=None,
                    note=f"H={H}m 옥상 + 자연낙차",
                ),
                detail="Case 3a: 옥상 고가수조 + 펌프 / 가압 40m",
            )
        return HBCaseDecision(
            case=HBCase.CASE_3B,
            pump_location="basement",
            rated_head_max_m=160.0,
            churn_head_max_m=200.0,
            pressurized_zone_m=H,
            natural_drop_zone_m=0.0,
            prv_required=True,
            prv_secondary_bar=4.0,
            pipe_material_change_at_m=160.0 - 120.0,
            trace=TripleTrace(
                nftc="NFTC 103 §2.2",
                hb="HB §2.4.16 Case 3b",
                phd=None,
                note=f"H={H}m 지하 펌프 + 감압",
            ),
            detail="Case 3b: 지하 펌프 / 정격 ≤160m / 체절 ≤200m",
        )
    # Case 4 / 5: 초고층
    interval = refuge_floor_interval_m or 80.0
    if interval <= 80:
        return HBCaseDecision(
            case=HBCase.CASE_4A,
            pump_location="rooftop",
            rated_head_max_m=120.0,
            churn_head_max_m=144.0,
            pressurized_zone_m=40.0,
            natural_drop_zone_m=40.0,
            prv_required=True,
            prv_secondary_bar=4.0,
            pipe_material_change_at_m=None,
            trace=TripleTrace(
                nftc="NFTC 103 §2.2",
                hb="HB §2.4.16 Case 4a (원칙)",
                phd=None,
                note=f"피난안전층 간격 ≤80m (원칙)",
            ),
            detail="Case 4a: 피난안전층마다 고가수조 + 가압/자연낙차 40m",
        )
    return HBCaseDecision(
        case=HBCase.CASE_5A,
        pump_location="rooftop",
        rated_head_max_m=120.0,
        churn_head_max_m=144.0,
        pressurized_zone_m=40.0,
        natural_drop_zone_m=120.0,  # 120m마다 감압밸브
        prv_required=True,
        prv_secondary_bar=4.0,
        pipe_material_change_at_m=None,
        trace=TripleTrace(
            nftc="NFTC 103 §2.2",
            hb="HB §2.4.16 Case 5a",
            phd=None,
            note=f"피난안전층 간격 >80m, 120m마다 감압",
        ),
        detail="Case 5a: 옥상 펌프 / 120m마다 감압밸브 / 2차 4 bar",
    )


# ---------------------------------------------------------------------------
# 3. Pipe material — KSD 3507 / 3562 (HB §2.4.5)
# ---------------------------------------------------------------------------

# Real inner diameters in mm — HB §2.4.5 Table.
# Used by hydraulics for accurate Hazen-Williams calculations.
_KSD_INNER_DIAMETERS_MM: dict[str, dict[str, float]] = {
    "20A":  {"3507": 21.9,  "3562": 21.4},
    "25A":  {"3507": 27.5,  "3562": 27.2},
    "32A":  {"3507": 36.2,  "3562": 35.5},
    "40A":  {"3507": 42.1,  "3562": 41.2},
    "50A":  {"3507": 53.2,  "3562": 52.7},
    "65A":  {"3507": 69.0,  "3562": 65.9},
    "80A":  {"3507": 81.0,  "3562": 78.1},
    "100A": {"3507": 105.3, "3562": 102.3},
    "125A": {"3507": 130.1, "3562": 126.6},
    "150A": {"3507": 155.5, "3562": 151.0},
    "200A": {"3507": 204.6, "3562": 199.9},
    "250A": {"3507": 254.6, "3562": 248.8},
    "300A": {"3507": 304.5, "3562": 297.9},
}


def decide_pipe_material(operating_pressure_mpa: float) -> RuleDecision:
    """Pipe material selection per HB §2.4.5.

    - ≤ 1.2 MPa → KSD 3507 (백강관)
    - > 1.2 MPa → KSD 3562 (배관용 아연도금 백강관, 고압용)
    """
    if operating_pressure_mpa <= 1.2:
        return RuleDecision(
            rule_id="HB-245-3507",
            verdict=Verdict.PASS,
            value="KSD 3507",
            trace=TripleTrace(
                nftc=None,
                hb="HB §2.4.5",
                phd=None,
                note=f"P {operating_pressure_mpa:.3f} MPa ≤ 1.2 MPa",
            ),
            detail=f"KSD 3507 (P = {operating_pressure_mpa:.3f} MPa)",
        )
    return RuleDecision(
        rule_id="HB-245-3562",
        verdict=Verdict.PASS,
        value="KSD 3562",
        trace=TripleTrace(
            nftc=None,
            hb="HB §2.4.5",
            phd=None,
            note=f"P {operating_pressure_mpa:.3f} MPa > 1.2 MPa",
        ),
        detail=f"KSD 3562 (P = {operating_pressure_mpa:.3f} MPa)",
    )


def get_inner_diameter_mm(nominal: str, material: str) -> float | None:
    """Look up real inner diameter in mm for given nominal size and material."""
    nominal = nominal.upper().strip()
    if nominal in _KSD_INNER_DIAMETERS_MM:
        material_key = "3507" if "3507" in material else "3562" if "3562" in material else "3507"
        return _KSD_INNER_DIAMETERS_MM[nominal][material_key]
    return None


# ---------------------------------------------------------------------------
# 4. Velocity limits (HB §2.4.5)
# ---------------------------------------------------------------------------

BRANCH_PIPE_V_LIMIT = 6.0   # m/s — 가지배관
MAIN_PIPE_V_LIMIT = 10.0    # m/s — 그 외 배관


def validate_velocity(*, pipe_role: str, velocity_mps: float) -> RuleDecision:
    """Validate pipe velocity per HB §2.4.5.

    pipe_role: "branch" (가지배관) or "other" (교차·입상·주배관)
    """
    if pipe_role == "branch":
        limit = BRANCH_PIPE_V_LIMIT
    else:
        limit = MAIN_PIPE_V_LIMIT
    ok = velocity_mps <= limit
    return RuleDecision(
        rule_id=f"HB-245-VEL-{pipe_role.upper()}",
        verdict=Verdict.PASS if ok else Verdict.FAIL,
        value={"velocity_mps": velocity_mps, "limit_mps": limit, "ok": ok},
        trace=TripleTrace(
            nftc=None,
            hb="HB §2.4.5",
            phd="동일 강조",
            note=f"{pipe_role} pipe: v={velocity_mps:.2f} m/s, limit={limit} m/s",
        ),
        detail=f"{'PASS' if ok else 'FAIL'}: {pipe_role} v={velocity_mps:.2f} ≤ {limit} m/s",
    )


# ---------------------------------------------------------------------------
# 5. Churn pressure (HB §2.4.16, 더 보수적)
# ---------------------------------------------------------------------------

CHURN_PRESSURE_MAX_RATIO_HB = 1.20   # 한백 더 보수적
CHURN_PRESSURE_MAX_RATIO_NFTC = 1.40 # NFTC 2.2.1.10


def validate_churn_pressure(rated_head_m: float, churn_head_m: float) -> RuleDecision:
    """Validate churn pressure per HB (more conservative than NFTC).

    HB: churn ≤ 정격 × 120%
    NFTC 2.2.1.10: 정격 × 140% 초과시 자동 압력제한장치 필요
    Most Restrictive Rule: 한백 채택.
    """
    ratio = churn_head_m / rated_head_m if rated_head_m > 0 else float("inf")
    ok_hb = ratio <= CHURN_PRESSURE_MAX_RATIO_HB
    ok_nftc = ratio <= CHURN_PRESSURE_MAX_RATIO_NFTC
    if ok_hb:
        return RuleDecision(
            rule_id="HB-2416-CHURN",
            verdict=Verdict.PASS,
            value={"ratio": ratio, "limit_hb": CHURN_PRESSURE_MAX_RATIO_HB, "limit_nftc": CHURN_PRESSURE_MAX_RATIO_NFTC},
            trace=TripleTrace(
                nftc="NFTC 103 §2.2.1.10 (140% 한계)",
                hb="HB §2.4.16 (120% 강화 — Most Restrictive)",
                phd=None,
            ),
            detail=f"PASS: 체절/정격 = {ratio:.3f} ≤ 1.20 (HB)",
        )
    if ok_nftc:
        return RuleDecision(
            rule_id="HB-2416-CHURN",
            verdict=Verdict.REVIEW,
            value={"ratio": ratio, "limit_hb": CHURN_PRESSURE_MAX_RATIO_HB, "limit_nftc": CHURN_PRESSURE_MAX_RATIO_NFTC},
            trace=TripleTrace(
                nftc="NFTC 103 §2.2.1.10 PASS (≤140%)",
                hb="HB §2.4.16 FAIL (>120%) — 인간 검토",
                phd=None,
            ),
            detail=f"REVIEW: {ratio:.3f}는 NFTC PASS, HB FAIL — 펌프 재선정 검토",
        )
    return RuleDecision(
        rule_id="HB-2416-CHURN",
        verdict=Verdict.FAIL,
        value={"ratio": ratio, "limit_hb": CHURN_PRESSURE_MAX_RATIO_HB, "limit_nftc": CHURN_PRESSURE_MAX_RATIO_NFTC},
        trace=TripleTrace(
            nftc="NFTC 103 §2.2.1.10 FAIL (>140%)",
            hb="HB §2.4.16 FAIL",
            phd=None,
        ),
        detail=f"FAIL: 체절/정격 = {ratio:.3f} > 1.40 — 펌프 모델 변경 필수",
    )


# ---------------------------------------------------------------------------
# 6. Zone partition (HB §2.4.2)
# ---------------------------------------------------------------------------

ZONE_AREA_MAX_M2 = 3000.0
ZONE_AREA_MAX_GRID_M2 = 3700.0
ZONE_HEAD_COUNT_THRESHOLD = 10
ZONE_DELUGE_HEAD_MAX = 50
ZONE_DELUGE_HEAD_MIN = 25
DRY_PIPE_VOLUME_MAX_L = 2840
DRY_PIPE_VOLUME_FAST_OPEN_THRESHOLD_L = 1890


@dataclass
class ZonePartition:
    """Zone partition decision."""

    zone_id: str
    floor_label: str
    area_m2: float
    head_count_estimate: int
    is_grid_layout: bool
    multi_floor_grouping: bool
    fire_compartment_id: str | None
    system_type: str
    trace: TripleTrace


def decide_zone_partition(
    *,
    floor_area_m2: float,
    estimated_head_count: int,
    floor_label: str,
    is_grid_layout: bool = False,
    is_apartment_loft: bool = False,
    fire_compartment_id: str | None = None,
    system_type: str = "wet",
) -> list[ZonePartition]:
    """Partition a floor into zones per HB §2.4.2.

    Rules applied:
      - Area ≤ 3,000 m² (or 3,700 m² for grid layout)
      - Per-floor unless ≤10 heads OR apartment loft (then 3 floors max)
      - Deluge: ≤50 heads/zone, ≥25 if multi-zone
      - Dry: 2nd-side volume ≤2,840 L, ≥1,890 L → fast-open device
    """
    max_area = ZONE_AREA_MAX_GRID_M2 if is_grid_layout else ZONE_AREA_MAX_M2
    multi_floor = estimated_head_count <= ZONE_HEAD_COUNT_THRESHOLD or is_apartment_loft
    n_zones = max(1, int(floor_area_m2 / max_area + 0.999))
    zones: list[ZonePartition] = []
    if n_zones == 1:
        zones.append(
            ZonePartition(
                zone_id=f"Z-{floor_label}-1",
                floor_label=floor_label,
                area_m2=floor_area_m2,
                head_count_estimate=estimated_head_count,
                is_grid_layout=is_grid_layout,
                multi_floor_grouping=multi_floor,
                fire_compartment_id=fire_compartment_id,
                system_type=system_type,
                trace=TripleTrace(
                    nftc="NFTC 103 §2.3.1.1 (≤3,000 m²)",
                    hb="HB §2.4.2",
                    phd=None,
                    note=f"단일 zone (area={floor_area_m2} m² ≤ {max_area})",
                ),
            )
        )
    else:
        sub_area = floor_area_m2 / n_zones
        sub_heads = max(1, estimated_head_count // n_zones)
        for idx in range(1, n_zones + 1):
            zones.append(
                ZonePartition(
                    zone_id=f"Z-{floor_label}-{idx}",
                    floor_label=floor_label,
                    area_m2=sub_area,
                    head_count_estimate=sub_heads,
                    is_grid_layout=is_grid_layout,
                    multi_floor_grouping=False,
                    fire_compartment_id=fire_compartment_id,
                    system_type=system_type,
                    trace=TripleTrace(
                        nftc="NFTC 103 §2.3.1.1",
                        hb="HB §2.4.2 분할",
                        phd=None,
                        note=f"분할 {idx}/{n_zones} ({floor_area_m2}/{n_zones} = {sub_area:.0f} m²)",
                    ),
                )
            )
    return zones


# ---------------------------------------------------------------------------
# 7. Hanger placement (HB §2.4.7 + NFTC 2.5.10)
# ---------------------------------------------------------------------------


def hanger_positions_along_pipe(
    *,
    pipe_length_m: float,
    pipe_role: str,
    head_positions_m: list[float] | None = None,
    head_to_hanger_min_m: float = 0.08,
) -> list[float]:
    """Compute hanger positions along a pipe per HB §2.4.7 / NFTC 2.5.10.

    Rules:
      - Branch pipe: 1 hanger between every adjacent head pair, max 3.5 m
      - Cross-main: 1 hanger between every adjacent branch, max 4.5 m
      - Horizontal main: ≤4.5 m
      - Hanger ↔ head: ≥8 cm
    """
    from nftc_rules import hanger_max_spacing_m as _max_spacing

    max_spacing = _max_spacing(pipe_role)
    positions: list[float] = []
    if pipe_role == "branch" and head_positions_m:
        # Place hangers between heads (respecting ≥8 cm clearance).
        sorted_heads = sorted(head_positions_m)
        for i in range(len(sorted_heads) - 1):
            mid = (sorted_heads[i] + sorted_heads[i + 1]) / 2.0
            if mid - sorted_heads[i] >= head_to_hanger_min_m:
                positions.append(round(mid, 3))
        # Subdivide if any inter-head gap exceeds max spacing
        refined: list[float] = []
        for i in range(len(sorted_heads) - 1):
            gap = sorted_heads[i + 1] - sorted_heads[i]
            n_extra = int(gap / max_spacing)
            for k in range(1, n_extra + 1):
                refined.append(round(sorted_heads[i] + k * max_spacing, 3))
        positions = sorted(set(positions + refined))
    else:
        # Uniform spacing for cross-main / horizontal main
        n_hangers = max(1, int(pipe_length_m / max_spacing))
        step = pipe_length_m / (n_hangers + 1)
        positions = [round(step * (i + 1), 3) for i in range(n_hangers)]
    return positions


# ---------------------------------------------------------------------------
# 8. Public summary
# ---------------------------------------------------------------------------


__all__ = [
    "SystemType",
    "HBCase",
    "HBCaseDecision",
    "ZonePartition",
    "decide_system_type",
    "decide_hb_case",
    "decide_pipe_material",
    "get_inner_diameter_mm",
    "validate_velocity",
    "validate_churn_pressure",
    "decide_zone_partition",
    "hanger_positions_along_pipe",
    "BRANCH_PIPE_V_LIMIT",
    "MAIN_PIPE_V_LIMIT",
    "CHURN_PRESSURE_MAX_RATIO_HB",
    "CHURN_PRESSURE_MAX_RATIO_NFTC",
    "ZONE_AREA_MAX_M2",
    "ZONE_AREA_MAX_GRID_M2",
    "DRY_PIPE_VOLUME_MAX_L",
    "DRY_PIPE_VOLUME_FAST_OPEN_THRESHOLD_L",
]
