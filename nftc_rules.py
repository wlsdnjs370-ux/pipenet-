"""
nftc_rules.py — NFTC 101~109 화재안전기술기준 룰 엔진

This module encodes Korean fire safety technical codes (NFTC) as deterministic
decision functions. All thresholds are sourced from the relevant NFTC clauses
(2024.07.01 revision). Each function returns the decision plus a 3-tuple trace
so that callers can record exactly which clause produced which value.

The module is intentionally pure Python and has no I/O or framework dependencies.
It is consumed by:
  - auto_design.py (zone partition / head spec decisions)
  - pipeline_orchestrator.py (pipeline orchestration + 3-tier traces)
  - pipenet_validator.py (post-validation cross-checks)

Coverage map:
  - NFTC 2.1.1   reference head count table (9 rules)
  - NFTC 2.2.1   pressure / flow minimums (≥0.1 MPa, ≥80 LPM)
  - NFTC 2.3.1   protection zone area (≤3,000 m² / grid 3,700 m²)
  - NFTC 2.5.10  hanger spacing (3.5 / 4.5 / 8 cm)
  - NFTC 2.6.1.6 high-rise alarm cascade (11F+ / apt 16F+)
  - NFTC 2.7.3   horizontal distance R (5 categories)
  - NFTC 2.7.5.5 fast-response head 5 mandated locations
  - NFTC 2.7.6   temperature rating table (4 tiers + 4m factory clause)
  - NFTC 2.7.7.1 head clearance (60 cm radius, 10 cm wall exception)
  - NFTC 2.7.7.2 head-to-ceiling (≤30 cm)
  - NFTC 2.9.3.2 emergency power (≥20 minutes)
  - NFTC 2.13    combined water-supply / pump / connection
  - NFTC 103B    ESFR / early-suppression sprinkler branch
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from typing import Any


# ---------------------------------------------------------------------------
# 0. Common types
# ---------------------------------------------------------------------------

class Verdict(str, Enum):
    """Outcome of a single rule evaluation."""

    PASS = "PASS"
    FAIL = "FAIL"
    REVIEW = "REVIEW"
    NA = "NA"


@dataclass(frozen=True)
class TripleTrace:
    """3-source citation for every automated decision.

    The pipeline guarantees that every rule decision carries this triple so
    audit / review / sign-off can answer "which clause decided this?" instantly.
    """

    nftc: str | None = None
    hb: str | None = None
    phd: str | None = None
    note: str | None = None

    def to_dict(self) -> dict[str, str | None]:
        return {"NFTC": self.nftc, "HB": self.hb, "PhD": self.phd, "note": self.note}


@dataclass(frozen=True)
class RuleDecision:
    """Decision + value + 3-source trace, used by all NFTC rule functions."""

    rule_id: str
    verdict: Verdict
    value: Any
    trace: TripleTrace
    detail: str = ""

    def to_dict(self) -> dict[str, Any]:
        return {
            "rule_id": self.rule_id,
            "verdict": self.verdict.value,
            "value": self.value,
            "trace": self.trace.to_dict(),
            "detail": self.detail,
        }


# ---------------------------------------------------------------------------
# 1. NFTC 2.1.1 — Reference head count table
# ---------------------------------------------------------------------------

# Hierarchical decision tree. Each row is matched in order; the first match wins.
# Source: NFTC 103, Table 2.1.1.1 (closed-head sprinkler reference head count).
_REFERENCE_COUNT_TABLE: list[dict[str, Any]] = [
    # 10층 이하
    {
        "rule_id": "NFTC-211-A",
        "predicate": lambda m: (
            m.get("floors_total", 0) <= 10
            and m.get("use") in {"factory", "warehouse"}
            and m.get("has_special_combustible") is True
        ),
        "count": 30,
        "label": "10층 이하 공장·창고로 특수가연물 저장·취급",
    },
    {
        "rule_id": "NFTC-211-B",
        "predicate": lambda m: (
            m.get("floors_total", 0) <= 10
            and m.get("use") in {"factory", "warehouse"}
        ),
        "count": 20,
        "label": "10층 이하 그 외 공장·창고",
    },
    {
        "rule_id": "NFTC-211-C",
        "predicate": lambda m: (
            m.get("floors_total", 0) <= 10
            and m.get("use") in {"neighborhood", "retail", "transit", "complex"}
        ),
        "count": 30,
        "label": "10층 이하 근린·판매·운수·복합건축물",
    },
    {
        "rule_id": "NFTC-211-D",
        "predicate": lambda m: (
            m.get("floors_total", 0) <= 10
            and m.get("use") not in {"apartment"}
            and m.get("head_attach_h_m", 0) <= 8
        ),
        "count": 20,
        "label": "10층 이하 그 외 (헤드 부착높이 8m 이하 일반)",
    },
    # Note: head attach height rule applies to "기타" buildings.
    {
        "rule_id": "NFTC-211-E",
        "predicate": lambda m: (
            m.get("floors_total", 0) <= 10
            and m.get("head_attach_h_m", 0) > 8
        ),
        "count": 20,
        "label": "10층 이하 기타 (헤드 부착높이 8m 이상)",
    },
    {
        "rule_id": "NFTC-211-F",
        "predicate": lambda m: (
            m.get("floors_total", 0) <= 10
            and m.get("head_attach_h_m", 0) <= 8
            and m.get("use") == "other_low"
        ),
        "count": 10,
        "label": "10층 이하 기타 (헤드 부착높이 8m 미만)",
    },
    # 11층 이상 / 지하역사 / 지하가
    {
        "rule_id": "NFTC-211-G",
        "predicate": lambda m: (
            m.get("floors_total", 0) >= 11
            or m.get("use") in {"underground_station", "underground_arcade"}
        ),
        "count": 30,
        "label": "11층 이상 특정소방대상물 / 지하역사 / 지하가",
    },
    # 아파트
    {
        "rule_id": "NFTC-211-I",
        "predicate": lambda m: (
            m.get("use") == "apartment"
            and m.get("connected_to_basement_parking") is True
        ),
        "count": 30,
        "label": "공동주택 + 지하주차장 연결",
    },
    {
        "rule_id": "NFTC-211-H",
        "predicate": lambda m: m.get("use") == "apartment",
        "count": 10,
        "label": "공동주택(아파트) 일반 세대",
    },
]


def decide_reference_count(building_meta: dict[str, Any]) -> RuleDecision:
    """Decide reference head count [NFTC 2.1.1].

    The reference count drives water-source volume and pump flow sizing.
    Required keys in `building_meta`:
      - floors_total: int (total above-ground floors)
      - use: str — one of {factory, warehouse, neighborhood, retail, transit,
        complex, apartment, underground_station, underground_arcade, other_low,
        other_high}
      - has_special_combustible: bool (NFTC 위험물 시행령 별표 2 해당 여부)
      - head_attach_h_m: float (헤드 부착높이; 평균 천장고)
      - connected_to_basement_parking: bool (apartment only)

    Returns RuleDecision with .value = int (head count).
    """
    for row in _REFERENCE_COUNT_TABLE:
        if row["predicate"](building_meta):
            return RuleDecision(
                rule_id=row["rule_id"],
                verdict=Verdict.PASS,
                value=row["count"],
                trace=TripleTrace(
                    nftc="NFTC 103 §2.1.1",
                    hb="HB §2.4.2 (인용)",
                    phd=None,
                    note=row["label"],
                ),
                detail=f"{row['label']} → 기준개수 {row['count']}",
            )
    # Fallback when nothing matches.
    return RuleDecision(
        rule_id="NFTC-211-FALLBACK",
        verdict=Verdict.REVIEW,
        value=20,
        trace=TripleTrace(
            nftc="NFTC 103 §2.1.1 (해당 없음, 보수적 20개 적용)",
            hb=None,
            phd=None,
            note="building_meta가 표 9개 룰 어디에도 매칭되지 않음 — 인간 검토 필요",
        ),
        detail="9개 룰 중 매칭 없음 — 보수적으로 20개 가정",
    )


# ---------------------------------------------------------------------------
# 2. NFTC 2.7.6 — Temperature rating table
# ---------------------------------------------------------------------------

# 4-tier step function. Source: NFTC 103, Table 2.7.6 (revised 2024.01.01).
_TEMPERATURE_RATING_TABLE: list[tuple[float, float, float, float]] = [
    # (ambient_min, ambient_max, rating_min, rating_max)
    (-1e9, 39.0, -1e9, 79.0),       # 39 ℃ 미만 → 79 ℃ 미만
    (39.0, 64.0, 79.0, 121.0),      # 39 ~ 64 ℃ → 79 ~ 121 ℃
    (64.0, 106.0, 121.0, 162.0),    # 64 ~ 106 ℃ → 121 ~ 162 ℃
    (106.0, 1e9, 162.0, 1e9),       # 106 ℃ 이상 → 162 ℃ 이상
]


def decide_temperature_rating(
    ambient_temp_c: float,
    *,
    is_factory_4m_high: bool = False,
    is_warehouse_4m_high: bool = False,
    is_rack_storage: bool = False,
) -> RuleDecision:
    """Decide head temperature rating [NFTC 2.7.6].

    Returns the temperature *band* the head's marking temperature must fall
    into. The 4 m factory/warehouse/rack proviso permits ≥121 ℃ regardless
    of ambient (to stabilize against false-trip from process heat).

    The Hanback formula `T_max_ambient = 0.9·T_marking − 27.3` is treated
    as an auxiliary continuous interpolation but the table wins on conflict.
    """
    # Proviso: NFTC 2.7.6 단서 (개정 2024.1.1) — 4m 이상 공장·창고·랙크식
    if is_factory_4m_high or is_warehouse_4m_high or is_rack_storage:
        return RuleDecision(
            rule_id="NFTC-276-PROVISO",
            verdict=Verdict.PASS,
            value={"min_c": 121.0, "max_c": None, "auxiliary_band": "≥121 ℃"},
            trace=TripleTrace(
                nftc="NFTC 103 §2.7.6 단서 (개정 2024.1.1)",
                hb="HB §2.4.9 — '4m 이상 공장' 동일 (단 v3는 창고·랙크식 누락)",
                phd=None,
                note="높이 4 m 이상 공장 · 창고 · 랙크식 → 주위온도 무관 ≥121 ℃ 적용 가능",
            ),
            detail=f"주위온도 {ambient_temp_c} ℃ 무관 — 4m↑ 단서로 ≥121 ℃ 가능",
        )
    # Main table (4-tier step).
    for amb_lo, amb_hi, rate_lo, rate_hi in _TEMPERATURE_RATING_TABLE:
        if amb_lo <= ambient_temp_c < amb_hi:
            return RuleDecision(
                rule_id="NFTC-276-MAIN",
                verdict=Verdict.PASS,
                value={"min_c": rate_lo, "max_c": rate_hi},
                trace=TripleTrace(
                    nftc="NFTC 103 §2.7.6 (Table 2.7.6)",
                    hb="HB §2.4.9 (공식 0.9T-27.3은 보조)",
                    phd=None,
                    note=f"주위 {ambient_temp_c} ℃ → 표시온도 {rate_lo}~{rate_hi} ℃",
                ),
                detail=f"주위 {ambient_temp_c} ℃ → 표 매칭 → 표시온도 {rate_lo}~{rate_hi} ℃",
            )
    # Should not happen given table covers (-inf, +inf).
    return RuleDecision(
        rule_id="NFTC-276-FALLBACK",
        verdict=Verdict.REVIEW,
        value={"min_c": 79.0, "max_c": 121.0},
        trace=TripleTrace(nftc="NFTC 103 §2.7.6", note="표 매칭 실패 — 인간 검토"),
        detail="표 매칭 실패",
    )


def hb_temperature_formula(t_marking_c: float) -> float:
    """한백 §2.4.9 보조 공식: T_max_ambient = 0.9·T_marking − 27.3.

    표 4단계의 점간 보간 검증용. 표와 충돌 시 표가 우선이다.
    """
    return 0.9 * t_marking_c - 27.3


# ---------------------------------------------------------------------------
# 3. NFTC 2.7.3 — Horizontal distance R (5 categories)
# ---------------------------------------------------------------------------


def decide_horizontal_distance(
    *,
    room_use: str,
    structure: str = "non_fire_resistant",
    has_special_combustible: bool = False,
    is_rack_storage: bool = False,
) -> RuleDecision:
    """Decide horizontal distance R for sprinkler heads [NFTC 2.7.3].

    Decision tree (top wins):
      1. 무대부 OR 특수가연물        → R = 1.7 m  (2.7.3.1)
      2. 랙크식 창고                  → R = 2.5 m  (2.7.3.2)
         단, 특수가연물 랙크식        → R = 1.7 m  (2.7.3.2 단서)
      3. 공동주택 세대 거실          → R = 3.2 m  (2.7.3.3)
      4. 내화구조                     → R = 2.3 m  (2.7.3.4)
      5. 그 외 (비내화)              → R = 2.1 m  (2.7.3.4)
    """
    if room_use == "stage" or has_special_combustible:
        return RuleDecision(
            rule_id="NFTC-273-1",
            verdict=Verdict.PASS,
            value=1.7,
            trace=TripleTrace(
                nftc="NFTC 103 §2.7.3.1",
                hb="HB §2.4.8 동일",
                phd=None,
                note="무대부 또는 특수가연물 저장·취급",
            ),
            detail="R = 1.7 m (무대부·특수가연물)",
        )
    if is_rack_storage:
        if has_special_combustible:
            r_value = 1.7
            label = "특수가연물 랙크식 → 1.7 m"
        else:
            r_value = 2.5
            label = "랙크식 창고 → 2.5 m"
        return RuleDecision(
            rule_id="NFTC-273-2",
            verdict=Verdict.PASS,
            value=r_value,
            trace=TripleTrace(
                nftc="NFTC 103 §2.7.3.2",
                hb=None,  # HB does not specify
                phd=None,
                note=label,
            ),
            detail=f"R = {r_value} m ({label})",
        )
    if room_use == "apartment_living":
        return RuleDecision(
            rule_id="NFTC-273-3",
            verdict=Verdict.PASS,
            value=3.2,
            trace=TripleTrace(
                nftc="NFTC 103 §2.7.3.3",
                hb=None,
                phd=None,
                note="공동주택 세대 내 거실 (헤드 형식승인 유효반경 적용)",
            ),
            detail="R = 3.2 m (공동주택 세대 거실)",
        )
    if structure == "fire_resistant":
        return RuleDecision(
            rule_id="NFTC-273-4-FR",
            verdict=Verdict.PASS,
            value=2.3,
            trace=TripleTrace(
                nftc="NFTC 103 §2.7.3.4 단서",
                hb="HB §2.4.8 동일",
                phd=None,
                note="내화구조",
            ),
            detail="R = 2.3 m (내화구조)",
        )
    return RuleDecision(
        rule_id="NFTC-273-4",
        verdict=Verdict.PASS,
        value=2.1,
        trace=TripleTrace(
            nftc="NFTC 103 §2.7.3.4",
            hb="HB §2.4.8 동일",
            phd=None,
            note="비내화구조 (기본)",
        ),
        detail="R = 2.1 m (비내화)",
    )


# ---------------------------------------------------------------------------
# 4. NFTC 2.7.5.5 — Fast-response head mandate (5 locations)
# ---------------------------------------------------------------------------

_FAST_RESPONSE_LOCATIONS: set[str] = {
    "apartment_living",      # 공동주택 거실
    "welfare_living",        # 노유자시설 거실
    "officetel_bedroom",     # 오피스텔 침실
    "hotel_bedroom",         # 숙박시설 침실
    "hospital_ward",         # 병원·의원 입원실
}


def is_fast_response_required(room_use: str) -> RuleDecision:
    """Determine whether RTI ≤ 50 fast-response heads are mandated [NFTC 2.7.5.5].

    Mandatory locations (5):
      ① apartment_living    공동주택 거실
      ② welfare_living      노유자시설 거실
      ③ officetel_bedroom   오피스텔 침실
      ④ hotel_bedroom       숙박시설 침실
      ⑤ hospital_ward       병원·의원 입원실
    """
    if room_use in _FAST_RESPONSE_LOCATIONS:
        return RuleDecision(
            rule_id="NFTC-275-5",
            verdict=Verdict.PASS,
            value=True,
            trace=TripleTrace(
                nftc="NFTC 103 §2.7.5.5",
                hb="HB §2.4.10 — '사람 상주' 권고와 일치",
                phd=None,
                note=f"5종 의무 장소 매칭: {room_use}",
            ),
            detail=f"{room_use} → RTI ≤ 50 강제",
        )
    return RuleDecision(
        rule_id="NFTC-275-5-NA",
        verdict=Verdict.NA,
        value=False,
        trace=TripleTrace(
            nftc="NFTC 103 §2.7.5.5",
            hb="HB §2.4.10 — '사람 상주' 권고는 적용 가능",
            phd=None,
            note="NFTC 5종 의무 장소 미매칭",
        ),
        detail=f"{room_use}는 5종 미매칭 — 한백 권고 적용 여지",
    )


# ---------------------------------------------------------------------------
# 5. NFTC 2.7.7.1 — Head clearance (60 cm radius, 10 cm wall exception)
# ---------------------------------------------------------------------------


def validate_head_clearance(
    *,
    head_xy: tuple[float, float],
    obstacles: list[dict[str, Any]],
    walls: list[dict[str, Any]] | None = None,
    radius_m: float = 0.6,
    wall_min_m: float = 0.1,
) -> RuleDecision:
    """Validate that a head has 60 cm clearance to non-wall obstacles.

    `obstacles` is a list of dicts with at least:
      - "polygon" or "bbox": list of (x, y) representing the obstacle footprint
      - "is_wall": bool (walls have the 10 cm exception)
    """
    walls = walls or []
    for obs in obstacles:
        if obs.get("is_wall"):
            continue
        d = _min_distance_point_to_polygon(head_xy, obs.get("polygon") or _bbox_to_polygon(obs.get("bbox")))
        if d < radius_m:
            return RuleDecision(
                rule_id="NFTC-2771-OBSTACLE",
                verdict=Verdict.FAIL,
                value={"violating_obstacle": obs.get("id"), "distance_m": d},
                trace=TripleTrace(
                    nftc="NFTC 103 §2.7.7.1",
                    hb=None,
                    phd=None,
                    note=f"헤드↔장애물 거리 {d:.3f} m < 0.6 m",
                ),
                detail=f"FAIL: 헤드 {head_xy} → 장애물 {obs.get('id')} 거리 {d:.3f} m",
            )
    for wall in walls:
        d = _min_distance_point_to_polygon(head_xy, wall.get("polygon") or _bbox_to_polygon(wall.get("bbox")))
        if d < wall_min_m:
            return RuleDecision(
                rule_id="NFTC-2771-WALL",
                verdict=Verdict.FAIL,
                value={"violating_wall": wall.get("id"), "distance_m": d},
                trace=TripleTrace(
                    nftc="NFTC 103 §2.7.7.1 단서",
                    hb="HB §2.4.11 동일",
                    phd=None,
                    note=f"헤드↔벽 거리 {d:.3f} m < 0.1 m",
                ),
                detail=f"FAIL: 헤드 {head_xy} → 벽 {wall.get('id')} 거리 {d:.3f} m",
            )
    return RuleDecision(
        rule_id="NFTC-2771",
        verdict=Verdict.PASS,
        value={"clearance_ok": True},
        trace=TripleTrace(
            nftc="NFTC 103 §2.7.7.1",
            hb="HB §2.4.11 동일",
            phd=None,
            note="60 cm 살수공간 + 10 cm 벽 거리 모두 충족",
        ),
        detail="PASS: 살수공간·벽 거리 모두 충족",
    )


def _bbox_to_polygon(bbox: tuple[float, float, float, float] | None) -> list[tuple[float, float]]:
    if not bbox:
        return []
    x1, y1, x2, y2 = bbox
    return [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]


def _min_distance_point_to_polygon(point: tuple[float, float], polygon: list[tuple[float, float]]) -> float:
    """Minimum euclidean distance from a point to a polygon's edges."""
    if not polygon:
        return float("inf")
    best = float("inf")
    n = len(polygon)
    for i in range(n):
        a = polygon[i]
        b = polygon[(i + 1) % n]
        d = _distance_point_to_segment(point, a, b)
        if d < best:
            best = d
    return best


def _distance_point_to_segment(p: tuple[float, float], a: tuple[float, float], b: tuple[float, float]) -> float:
    px, py = p
    ax, ay = a
    bx, by = b
    dx = bx - ax
    dy = by - ay
    if dx == 0 and dy == 0:
        return ((px - ax) ** 2 + (py - ay) ** 2) ** 0.5
    t = ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)
    t = max(0.0, min(1.0, t))
    qx = ax + t * dx
    qy = ay + t * dy
    return ((px - qx) ** 2 + (py - qy) ** 2) ** 0.5


# ---------------------------------------------------------------------------
# 6. NFTC 2.13 — Combined water-supply / pump / connection
# ---------------------------------------------------------------------------


@dataclass
class CombinedWaterSupplyResult:
    """Result of NFTC 2.13 combined water supply / pump / hose connection."""

    combined: bool
    systems: list[str]
    tank_total_m3: float
    tank_min_total_m3: float
    pump_total_lpm: float
    pump_required_head_m: float
    hose_connection_count: int
    trace: TripleTrace
    breakdown: dict[str, Any] = field(default_factory=dict)


def decide_combined_water_supply(
    systems: dict[str, dict[str, float]],
    *,
    use_combined_tank: bool = True,
    use_combined_pump: bool = True,
) -> CombinedWaterSupplyResult:
    """Compute combined water supply / pump per NFTC 2.13.

    `systems` example:
        {
          "sprinkler":      {"v_m3": 42.0, "q_lpm": 2100, "h_m": 60.0},
          "indoor_hydrant": {"v_m3": 5.2,  "q_lpm": 260,  "h_m": 50.0},
        }

    Returns combined volume, pump flow, required head per NFTC 2.13.1~2.13.4.
    """
    keys = list(systems.keys())
    v_total = sum(d.get("v_m3", 0.0) for d in systems.values()) if use_combined_tank else 0.0
    # NFTC 2.13.1: 저수량 합산
    v_min_total = v_total / 0.8  # HB §2.4.3 — 유효용량 80% 룰
    # NFTC 2.13.2: 펌프 토출량 합산 (동시작동 가정)
    q_total = sum(d.get("q_lpm", 0.0) for d in systems.values()) if use_combined_pump else 0.0
    h_required = max((d.get("h_m", 0.0) for d in systems.values()), default=0.0)
    # NFTC 2.13.4: 송수구 (SP 기준 따름)
    hc_count = max(1, int(systems.get("sprinkler", {}).get("zone_area_m2", 0.0) / 3000.0))
    hc_count = min(5, hc_count)
    return CombinedWaterSupplyResult(
        combined=True,
        systems=keys,
        tank_total_m3=round(v_total, 3),
        tank_min_total_m3=round(v_min_total, 3),
        pump_total_lpm=round(q_total, 1),
        pump_required_head_m=round(h_required, 1),
        hose_connection_count=hc_count,
        trace=TripleTrace(
            nftc="NFTC 103 §2.13.1~2.13.4",
            hb="HB 통합 사무소 표준 (SP+옥내소화전)",
            phd=None,
            note=f"겸용 시스템: {', '.join(keys)}",
        ),
        breakdown={k: dict(v) for k, v in systems.items()},
    )


# ---------------------------------------------------------------------------
# 7. NFTC 103B — ESFR / Early-Suppression Fast-Response branch
# ---------------------------------------------------------------------------

# K-factor lookup table for ESFR / CMSA / Large Drop heads.
_ESFR_K_TABLE: list[tuple[float, int, str, float | None]] = [
    # (K_gpm_psi05, K_lpm_bar05, role, ceiling_max_m)
    (5.6, 80, "표준 (NFTC 103 본)", None),
    (8.0, 115, "EV 충전 · 중급위험", None),
    (11.2, 160, "CMSA 표준", None),
    (14.0, 200, "ESFR 일반", 9.1),
    (16.8, 240, "CMSA / Large Drop", None),
    (22.4, 320, "ESFR 고천장", 12.0),
    (25.2, 360, "ESFR 12m↑ 플라스틱", 13.7),
]


@dataclass
class ESFRDecision:
    """Result of NFTC 103B activation + K-factor decision."""

    activated: bool
    k_lpm_bar05: int | None
    role: str | None
    ceiling_max_m: float | None
    in_rack_required: bool
    human_review_required: bool
    trace: TripleTrace
    detail: str


def decide_esfr_branch(
    *,
    room_use: str,
    ceiling_h_m: float,
) -> ESFRDecision:
    """Decide whether NFTC 103B ESFR branch activates and which K-factor to use.

    Activation:
      - room_use == "rack_storage" → activated, K depends on ceiling_h
      - room_use == "warehouse" AND ceiling_h ≥ 9.1 m → activated
      - room_use == "EV_charging" → K115 + preaction system (NFTC 103 본 트랙)
      - else → not activated (NFTC 103 본)

    Ceiling height bands (when activated):
      - ≤ 9.1 m  → K = 200 (K14)   ESFR 일반
      - ≤ 12 m   → K = 320 (K22.4) ESFR 고천장
      - ≤ 13.7 m → K = 360 (K25.2) ESFR 12m↑ 플라스틱
      - > 13.7 m → in-rack 보강 + 인간 검토 강제
    """
    # EV charging: not strictly NFTC 103B, but K115 special branch.
    if room_use == "EV_charging":
        return ESFRDecision(
            activated=False,
            k_lpm_bar05=115,
            role="EV 충전 K115 (NFTC 103 본 + 특수 조건)",
            ceiling_max_m=None,
            in_rack_required=False,
            human_review_required=False,
            trace=TripleTrace(
                nftc="NFTC 103 본 + EV 특수 조건",
                hb="HB EV 충전 가이드라인",
                phd=None,
                note="K115 + preaction 시스템",
            ),
            detail="EV 충전 → NFTC 103 본 트랙 K115",
        )
    # Activation gate.
    activate = (
        room_use == "rack_storage"
        or (room_use == "warehouse" and ceiling_h_m >= 9.1)
    )
    if not activate:
        return ESFRDecision(
            activated=False,
            k_lpm_bar05=None,
            role=None,
            ceiling_max_m=None,
            in_rack_required=False,
            human_review_required=False,
            trace=TripleTrace(
                nftc="NFTC 103 본 트랙 유지",
                hb=None,
                phd=None,
                note=f"활성 조건 미충족 (use={room_use}, ceiling_h={ceiling_h_m})",
            ),
            detail="NFTC 103 본 트랙 (ESFR 분기 비활성)",
        )
    # K-factor by ceiling height.
    if ceiling_h_m <= 9.1:
        return ESFRDecision(
            activated=True,
            k_lpm_bar05=200,
            role="ESFR 일반 (K14)",
            ceiling_max_m=9.1,
            in_rack_required=False,
            human_review_required=False,
            trace=TripleTrace(
                nftc="NFTC 103B (ESFR 적용)",
                hb=None,
                phd=None,
                note="천장 ≤9.1 m → K=200",
            ),
            detail=f"NFTC 103B 활성, K=200 (천장 {ceiling_h_m} m)",
        )
    if ceiling_h_m <= 12.0:
        return ESFRDecision(
            activated=True,
            k_lpm_bar05=320,
            role="ESFR 고천장 (K22.4)",
            ceiling_max_m=12.0,
            in_rack_required=False,
            human_review_required=False,
            trace=TripleTrace(
                nftc="NFTC 103B",
                hb=None,
                phd=None,
                note="천장 9.1~12 m → K=320",
            ),
            detail=f"NFTC 103B 활성, K=320 (천장 {ceiling_h_m} m)",
        )
    if ceiling_h_m <= 13.7:
        return ESFRDecision(
            activated=True,
            k_lpm_bar05=360,
            role="ESFR 12m↑ 플라스틱 (K25.2)",
            ceiling_max_m=13.7,
            in_rack_required=False,
            human_review_required=True,
            trace=TripleTrace(
                nftc="NFTC 103B",
                hb=None,
                phd=None,
                note="천장 12~13.7 m → K=360 + 저장물 인간 검토",
            ),
            detail=f"NFTC 103B 활성, K=360 (천장 {ceiling_h_m} m) — 저장물 검토 필수",
        )
    # > 13.7 m: in-rack required + human review.
    return ESFRDecision(
        activated=True,
        k_lpm_bar05=360,
        role="ESFR + in-rack 보강",
        ceiling_max_m=None,
        in_rack_required=True,
        human_review_required=True,
        trace=TripleTrace(
            nftc="NFTC 103B + NFPC 2.7.2 (랙크식 헤드 4/6m)",
            hb=None,
            phd=None,
            note="천장 > 13.7 m → 천장 ESFR + in-rack 보강 + 인간 검토",
        ),
        detail=f"NFTC 103B 활성 + in-rack 필수 (천장 {ceiling_h_m} m)",
    )


# ---------------------------------------------------------------------------
# 8. NFTC 2.6.1.6 — High-rise alarm cascade
# ---------------------------------------------------------------------------


def decide_alarm_cascade(*, fire_floor: int, total_floors: int, is_apartment: bool = False) -> RuleDecision:
    """Decide which floors should sound the alarm in 11F+ buildings [NFTC 2.6.1.6].

    Threshold: total_floors ≥ 11 (apartment 16F+).
    Cascade:
      - 2F+ 발화 → 발화층 + 직상 4개층
      - 1F 발화 → 발화층 + 직상 4개층 + 지하층
      - B1F 발화 → 발화층 + 직상층 + 기타 지하층
    """
    threshold = 16 if is_apartment else 11
    if total_floors < threshold:
        return RuleDecision(
            rule_id="NFTC-2616-NA",
            verdict=Verdict.NA,
            value={"all_floors": True},
            trace=TripleTrace(
                nftc="NFTC 103 §2.6.1.6",
                hb=None,
                phd=None,
                note=f"전체 {total_floors}층 < {threshold}층 — 일제 경보",
            ),
            detail="전 층 일제 경보",
        )
    # Compute cascade
    if fire_floor >= 2:
        floors = [fire_floor] + [fire_floor + i for i in range(1, 5) if fire_floor + i <= total_floors]
        clause = "2.6.1.6.1"
    elif fire_floor == 1:
        floors = [1, 2, 3, 4, 5, "B"]
        clause = "2.6.1.6.2"
    else:  # basement
        floors = [fire_floor, fire_floor + 1, "all_basement"]
        clause = "2.6.1.6.3"
    return RuleDecision(
        rule_id=f"NFTC-{clause.replace('.', '')}",
        verdict=Verdict.PASS,
        value={"alarm_floors": floors},
        trace=TripleTrace(
            nftc=f"NFTC 103 §{clause}",
            hb="HB §2.4.17 동일",
            phd=None,
            note=f"발화층 {fire_floor} → 경보층 {floors}",
        ),
        detail=f"발화층 {fire_floor} → {floors}",
    )


# ---------------------------------------------------------------------------
# 9. NFTC 2.5.10 — Hanger spacing
# ---------------------------------------------------------------------------


def hanger_max_spacing_m(pipe_role: str) -> float:
    """Max hanger spacing per NFTC 2.5.10."""
    if pipe_role == "branch":
        return 3.5  # NFTC 2.5.10.1
    if pipe_role in {"cross_main", "main"}:
        return 4.5  # NFTC 2.5.10.2
    return 4.5  # safe default


def head_to_hanger_min_m() -> float:
    """헤드 ↔ 행거 최소 간격 (NFTC 2.5.10.1, 상향식 헤드 8 cm)."""
    return 0.08


# ---------------------------------------------------------------------------
# 10. NFTC 2.7.7.2 — Head-to-ceiling distance
# ---------------------------------------------------------------------------


def validate_head_to_ceiling(
    *,
    head_z_m: float,
    ceiling_z_m: float,
    has_beam: bool = False,
    beam_clear_m: float = 0.0,
) -> RuleDecision:
    """Validate head-to-ceiling distance per NFTC 2.7.7.2.

    Default: ≤30 cm. With qualifying beam (천장~보 하단 ≥55 cm), ≤55 cm allowed.
    """
    distance = ceiling_z_m - head_z_m
    if distance < 0:
        return RuleDecision(
            rule_id="NFTC-2772-INVALID",
            verdict=Verdict.FAIL,
            value=distance,
            trace=TripleTrace(nftc="NFTC 103 §2.7.7.2"),
            detail=f"FAIL: 헤드가 천장 위에 있음 ({distance:.3f} m)",
        )
    limit = 0.55 if (has_beam and beam_clear_m >= 0.55) else 0.30
    if distance <= limit:
        return RuleDecision(
            rule_id="NFTC-2772",
            verdict=Verdict.PASS,
            value={"distance_m": distance, "limit_m": limit},
            trace=TripleTrace(
                nftc="NFTC 103 §2.7.7.2" + (" (보 예외)" if limit == 0.55 else ""),
                hb="HB §2.4.11 보 예외 동일",
                phd=None,
            ),
            detail=f"PASS: {distance:.3f} m ≤ {limit:.2f} m",
        )
    return RuleDecision(
        rule_id="NFTC-2772",
        verdict=Verdict.FAIL,
        value={"distance_m": distance, "limit_m": limit},
        trace=TripleTrace(nftc="NFTC 103 §2.7.7.2"),
        detail=f"FAIL: {distance:.3f} m > {limit:.2f} m",
    )


# ---------------------------------------------------------------------------
# 11. Pressure / flow minimums [NFTC 2.2.1.11]
# ---------------------------------------------------------------------------


def head_pressure_min_mpa() -> float:
    """헤드 최소 방수압력 0.1 MPa [NFTC 2.2.1.11]."""
    return 0.1


def head_pressure_max_mpa() -> float:
    """헤드 최대 방수압력 1.2 MPa [HB §2.4.2 — NFTC 침묵, 한백 추가]."""
    return 1.2


def head_flow_min_lpm() -> float:
    """헤드 최소 방수량 80 LPM [NFTC 2.2.1.11]."""
    return 80.0


def emergency_power_min_minutes() -> float:
    """비상전원 최소 시간 20분 [NFTC 2.9.3.2]."""
    return 20.0


# ---------------------------------------------------------------------------
# 12. Public summary helpers
# ---------------------------------------------------------------------------


def summarize_nftc_decisions(decisions: list[RuleDecision]) -> dict[str, Any]:
    """Summarize a list of RuleDecision into PASS/FAIL/REVIEW counts + details."""
    by_verdict: dict[str, int] = {v.value: 0 for v in Verdict}
    fail_items: list[dict[str, Any]] = []
    review_items: list[dict[str, Any]] = []
    for d in decisions:
        by_verdict[d.verdict.value] += 1
        if d.verdict == Verdict.FAIL:
            fail_items.append(d.to_dict())
        elif d.verdict == Verdict.REVIEW:
            review_items.append(d.to_dict())
    overall = "PASS" if by_verdict["FAIL"] == 0 else "FAIL"
    return {
        "overall": overall,
        "counts": by_verdict,
        "fails": fail_items,
        "reviews": review_items,
    }


__all__ = [
    "Verdict",
    "TripleTrace",
    "RuleDecision",
    "CombinedWaterSupplyResult",
    "ESFRDecision",
    "decide_reference_count",
    "decide_temperature_rating",
    "hb_temperature_formula",
    "decide_horizontal_distance",
    "is_fast_response_required",
    "validate_head_clearance",
    "decide_combined_water_supply",
    "decide_esfr_branch",
    "decide_alarm_cascade",
    "hanger_max_spacing_m",
    "head_to_hanger_min_m",
    "validate_head_to_ceiling",
    "head_pressure_min_mpa",
    "head_pressure_max_mpa",
    "head_flow_min_lpm",
    "emergency_power_min_minutes",
    "summarize_nftc_decisions",
]
