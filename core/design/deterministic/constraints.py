# -*- coding: utf-8 -*-
"""NFTC 수치의 유일한 진실 출처 — building.json → Constraints.

`Constraints` 는 frozen 이다. 하류 단계가 값을 고치려 하면 예외가 난다. 설계
중간에 "이 실만 R 을 늘리자" 같은 국소 변경이 들어오면 그 뒤로 어떤 검사도
무엇을 기준으로 통과했는지 말할 수 없게 된다.

값을 만들지 못하면 **기본값으로 채우지 말고 예외를 던진다**(`MissingBuildingFact`).
조용히 채운 값은 뒤의 모든 계산을 틀린 기준 위에서 통과시킨다. 결손은 GATE 로
돌아가 사람이 확정할 항목이지 코드가 추측할 항목이 아니다.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any

from . import nftc_tables as T


class MissingBuildingFact(Exception):
    """확정되지 않은 입력. `field` 는 GATE 결손 항목 폼이 그대로 쓴다."""

    def __init__(self, field: str, detail: str = ""):
        self.field = field
        self.detail = detail
        super().__init__(f"{field} 미확정{f' — {detail}' if detail else ''}")


@dataclass(frozen=True)
class FieldTrace:
    """어떤 필드가 어느 조문에서 나왔는지. constraints.json 의 `trace` 항목."""

    field: str
    code: str
    article: str
    text_hash: str
    effective_date: str
    note: str = ""

    def to_dict(self) -> dict[str, str]:
        return {
            "field": self.field, "code": self.code, "article": self.article,
            "text_hash": self.text_hash, "effective_date": self.effective_date,
            "note": self.note,
        }


@dataclass(frozen=True)
class Constraints:
    # ── 기준개수·수원 ──────────────────────────────
    scenario_head_count: int
    water_supply_m3: float
    discharge_minutes: int
    emergency_power_minutes: int

    # ── 헤드 ───────────────────────────────────────
    horizontal_distance_m: float
    head_spacing_square_m: float
    wall_clearance_max_m: float
    temp_rating_c: int
    quick_response_required: bool
    k_factor: float
    flow_lpm_min: float
    pressure_mpa_min: float
    pressure_mpa_max: float
    head_clearance_radius_m: float
    head_to_wall_clearance_m: float
    head_to_ceiling_max_m: float

    # ── 구역 ───────────────────────────────────────
    zone_area_max_m2: float
    zone_grid_relief_m2: float | None
    zone_floors_max: int
    spray_zone_heads_max: int
    spray_zone_heads_min_when_split: int

    # ── 배관 ───────────────────────────────────────
    branch_heads_per_side_max: int
    cross_main_min_dn: int
    tournament_forbidden: bool
    pipe_size_table: dict
    velocity_limit_mps: dict

    # ── 표 ─────────────────────────────────────────
    beam_clearance_table: tuple
    head_exempt_places: tuple

    # ── 추적 ───────────────────────────────────────
    nftc_effective_date: str
    trace: tuple

    def to_dict(self) -> dict[str, Any]:
        """constraints.json (schema fncadnet.constraints/1) 직렬화."""
        return {
            "schema": "fncadnet.constraints/1",
            "nftc_effective_date": self.nftc_effective_date,
            "scenario_head_count": self.scenario_head_count,
            "water_supply_m3": self.water_supply_m3,
            "discharge_minutes": self.discharge_minutes,
            "emergency_power_minutes": self.emergency_power_minutes,
            "horizontal_distance_m": self.horizontal_distance_m,
            "head_spacing_square_m": self.head_spacing_square_m,
            "wall_clearance_max_m": self.wall_clearance_max_m,
            "temp_rating_c": self.temp_rating_c,
            "quick_response_required": self.quick_response_required,
            "k_factor": self.k_factor,
            "flow_lpm_min": self.flow_lpm_min,
            "pressure_mpa": {"min": self.pressure_mpa_min, "max": self.pressure_mpa_max},
            "head_clearance_radius_m": self.head_clearance_radius_m,
            "head_to_wall_clearance_m": self.head_to_wall_clearance_m,
            "head_to_ceiling_max_m": self.head_to_ceiling_max_m,
            "zone_area_max_m2": self.zone_area_max_m2,
            "zone_grid_relief_m2": self.zone_grid_relief_m2,
            "zone_floors_max": self.zone_floors_max,
            "spray_zone_heads_max": self.spray_zone_heads_max,
            "spray_zone_heads_min_when_split": self.spray_zone_heads_min_when_split,
            "branch_heads_per_side_max": self.branch_heads_per_side_max,
            "cross_main_min_dn": self.cross_main_min_dn,
            "tournament_forbidden": self.tournament_forbidden,
            "pipe_size_table": {
                col: {str(dn): cap for dn, cap in rows.items()}
                for col, rows in self.pipe_size_table.items()
            },
            "velocity_limit_mps": {"_note": "사내 기준 — 법정 아님", **self.velocity_limit_mps},
            "beam_clearance_table": [
                {"horizontal_lt_m": lim, "vertical": req} if isinstance(req, str)
                else {"horizontal_lt_m": lim, "vertical_lt_m": req}
                for lim, req in self.beam_clearance_table
            ],
            "head_exempt_places": list(self.head_exempt_places),
            "trace": [t.to_dict() for t in self.trace],
        }


# ────────────────────────────────────────────────────────────────────────────
# 기준개수 (표 2.1.1.1)
# ────────────────────────────────────────────────────────────────────────────

def scenario_head_count(
    *,
    use: str,
    floors_total: int,
    is_underground_arcade: bool = False,
    has_special_combustible: bool = False,
    head_mount_height_m: float = 0.0,
    is_apartment_unit: bool = False,
    is_connected_parking: bool = False,
    has_retail_occupancy: bool = False,
) -> tuple[int, str]:
    """표 2.1.1.1 조회. 반환 `(기준개수, rule_code)`.

    행 순서가 곧 규칙이다. 아파트를 11층 행 뒤에 두면 국내 아파트 대부분이
    11층 이상으로 잡혀 기준개수가 10에서 30으로 뒤집힌다.

    [문서정합] 지시서 §5.2 스니펫은 본문에서 `building_has_retail` 을 참조하지만
    시그니처에 그 인자가 없다. 여기서는 `has_retail_occupancy` 로 명시 인자화했다.
    복합건축물의 30/20 판정이 이 값 하나에 달려 있어 암묵 전역으로 둘 수 없다.
    """
    if is_apartment_unit:
        return 10, T.RULE_APT.code
    if is_connected_parking:
        return 30, T.RULE_PARK.code
    if floors_total >= 11 or is_underground_arcade:
        return 30, T.RULE_HIGH.code

    if use in T.FACTORY_USES:
        if has_special_combustible:
            return 30, T.RULE_FACTORY_SPECIAL.code
        return 20, T.RULE_FACTORY.code

    if use in T.TWENTY_OR_THIRTY_USES:
        if use in T.RETAIL_USES or (use == "복합건축물" and has_retail_occupancy):
            return 30, T.RULE_RETAIL.code
        return 20, T.RULE_NEIGHBORHOOD.code

    if head_mount_height_m >= 8.0:
        return 20, T.RULE_OTHER_HIGH_MOUNT.code
    return 10, T.RULE_OTHER_LOW_MOUNT.code


# ────────────────────────────────────────────────────────────────────────────
# 수원·방사시간·비상전원
# ────────────────────────────────────────────────────────────────────────────

def water_supply(n: int, floors_total: int) -> tuple[float, int]:
    """반환 `(수원량 ㎥, 방사시간 분)`. 방사시간을 상수로 두면 안 된다.

    비상전원 시간도 같은 분기를 쓴다 — `emergency_power_minutes()` 참조.
    """
    for floor_min, per_head_m3, minutes, _rule in T.WATER_SUPPLY_TIERS:
        if floors_total >= floor_min:
            return n * per_head_m3, minutes
    raise AssertionError("WATER_SUPPLY_TIERS 마지막 행의 층수 하한이 0 이어야 한다")


def water_supply_rule(floors_total: int) -> str:
    for floor_min, _per_head_m3, _minutes, rule in T.WATER_SUPPLY_TIERS:
        if floors_total >= floor_min:
            return rule.code
    raise AssertionError("WATER_SUPPLY_TIERS 마지막 행의 층수 하한이 0 이어야 한다")


def emergency_power_minutes(floors_total: int) -> int:
    """비상전원 용량 시간. 방사시간과 같은 층수 분기를 따른다."""
    return water_supply(1, floors_total)[1]


# ────────────────────────────────────────────────────────────────────────────
# 수평거리 R (2.7.3)
# ────────────────────────────────────────────────────────────────────────────

def horizontal_distance(
    *,
    is_stage: bool = False,
    has_special_combustible: bool = False,
    is_rack_storage: bool = False,
    is_apartment_unit: bool = False,
    structure: str = "비내화구조",
) -> tuple[float, str]:
    """반환 `(R m, rule_code)`. 위에서부터 먼저 맞는 행이 이긴다."""
    if is_rack_storage:
        rule = T.RULE_R_RACK_SPECIAL if has_special_combustible else T.RULE_R_RACK
        return T.HORIZONTAL_DISTANCE_M[rule.code], rule.code
    if is_stage or has_special_combustible:
        return T.HORIZONTAL_DISTANCE_M[T.RULE_R_STAGE.code], T.RULE_R_STAGE.code
    if is_apartment_unit:
        return T.HORIZONTAL_DISTANCE_M[T.RULE_R_APT.code], T.RULE_R_APT.code
    rule = T.RULE_R_FIRE_RESISTANT if structure == "내화구조" else T.RULE_R_NON_FIRE_RESISTANT
    return T.HORIZONTAL_DISTANCE_M[rule.code], rule.code


def head_spacing_square(r_m: float) -> float:
    """정방형 배치 헤드 간격 S = √2 R."""
    return r_m * math.sqrt(2.0)


# ────────────────────────────────────────────────────────────────────────────
# 표시온도 (표 2.7.6)
# ────────────────────────────────────────────────────────────────────────────

def temp_rating(ambient_temp_max_c: float) -> tuple[int, str]:
    """최고 주위온도 → `(표준 표시온도 ℃, rule_code)`.

    구간은 NFTC 가 정하고, 그 구간 안에서 어떤 제품을 쓸지는
    `nftc_tables.STANDARD_MARKING_C` 가 정한다(시장 규격, 법정 아님).
    """
    for i, (amb_lo, amb_hi, _rate_lo, _rate_hi) in enumerate(T.TEMPERATURE_BANDS):
        if amb_lo <= ambient_temp_max_c < amb_hi:
            return T.STANDARD_MARKING_C[i], T.RULE_TEMP.code
    raise AssertionError("TEMPERATURE_BANDS 가 실수 전 구간을 덮지 않는다")


# ────────────────────────────────────────────────────────────────────────────
# building.json → Constraints
# ────────────────────────────────────────────────────────────────────────────

def _require(container: dict, key: str, path: str) -> Any:
    if key not in container or container[key] is None:
        raise MissingBuildingFact(path)
    return container[key]


def _head_mount_height_m(rooms: list[dict]) -> float:
    """헤드 부착높이 = 반자가 있으면 반자 높이, 없으면 슬래브 높이. 최댓값 채택.

    기준개수 표의 '그 밖의 것' 8m 경계에만 쓰이므로 건물 대푯값이면 충분하고,
    경계에서 보수적(큰 값 → 기준개수 20)으로 기울어야 한다.
    """
    heights: list[float] = []
    for room in rooms:
        ceiling = room.get("ceiling") or {}
        mm = (ceiling.get("finish_height_mm") if ceiling.get("has_finish")
              else ceiling.get("slab_height_mm"))
        if mm:
            heights.append(float(mm) / 1000.0)
    if not heights:
        raise MissingBuildingFact(
            "rooms[].ceiling", "반자 유무와 높이가 한 실도 확정되지 않았다")
    return max(heights)


def build_constraints(building: dict) -> Constraints:
    """building.json → Constraints. 이 함수 밖에서 NFTC 값을 만들지 마라."""
    b = _require(building, "building", "building")
    rooms: list[dict] = building.get("rooms") or []

    floors_total = int(_require(b, "floors_total", "building.floors_total"))
    structure = str(_require(b, "structure", "building.structure"))
    use = str(_require(b, "use", "building.use"))
    is_underground_arcade = bool(b.get("is_underground_arcade", False))
    is_connected_parking = bool(b.get("is_connected_parking", False))

    # 무대부·특수가연물·랙크식은 **실별로** 확정된다(GATE `special_hazard`). 건물
    # 키만 읽으면 사람이 실에 붙인 확정이 조용히 힘을 잃고 R 이 1.7 대신 2.3 으로
    # 남는다 — 한 실이라도 붙어 있으면 건물 전체 기준을 그쪽으로 당긴다.
    hazards = {room.get("special_hazard") for room in rooms}
    has_special_combustible = "특수가연물" in hazards
    is_stage = "무대부" in hazards
    is_rack_storage = use == "랙크식창고" or "랙크식창고" in hazards
    # 아파트 세대·판매시설 포함 여부도 사람이 확정한 용도에서 읽는다. 따로 묻지
    # 않는 이유는 같은 사실을 두 번 물으면 두 답이 어긋나기 때문이다.
    is_apartment_unit = use == "공동주택"
    has_retail_occupancy = any((room.get("use") or "") in T.RETAIL_USES for room in rooms)

    mount_h = _head_mount_height_m(rooms) if rooms else float(
        _require(b, "head_mount_height_m", "building.head_mount_height_m"))

    n, count_rule = scenario_head_count(
        use=use, floors_total=floors_total,
        is_underground_arcade=is_underground_arcade,
        has_special_combustible=has_special_combustible,
        head_mount_height_m=mount_h,
        is_apartment_unit=is_apartment_unit,
        is_connected_parking=is_connected_parking,
        has_retail_occupancy=has_retail_occupancy,
    )
    volume_m3, minutes = water_supply(n, floors_total)

    r_m, r_rule = horizontal_distance(
        is_stage=is_stage,
        has_special_combustible=has_special_combustible,
        is_rack_storage=is_rack_storage,
        is_apartment_unit=is_apartment_unit,
        structure=structure,
    )
    spacing = head_spacing_square(r_m)

    ambient = [room.get("ambient_temp_max_c") for room in rooms]
    ambient = [float(a) for a in ambient if a is not None]
    if rooms and not ambient:
        raise MissingBuildingFact(
            "rooms[].ambient_temp_max_c", "최고 주위온도가 한 실도 확정되지 않았다")
    rating_c, temp_rule = temp_rating(max(ambient) if ambient else 0.0)

    quick = any((room.get("use") or "") in T.QUICK_RESPONSE_ROOM_USES for room in rooms)

    trace = (
        _trace("scenario_head_count", count_rule),
        _trace("water_supply_m3", water_supply_rule(floors_total)),
        _trace("discharge_minutes", water_supply_rule(floors_total)),
        _trace("emergency_power_minutes", water_supply_rule(floors_total)),
        _trace("horizontal_distance_m", r_rule),
        _trace("temp_rating_c", temp_rule,
               note=f"구간은 법정, 대표값 {rating_c}℃ 는 시장 표준 제품"),
        _trace("quick_response_required", T.RULE_QUICK_RESPONSE.code),
        _trace("flow_lpm_min", T.RULE_PRESSURE_FLOW.code),
        _trace("head_clearance_radius_m", T.RULE_HEAD_CLEARANCE.code),
        _trace("head_to_ceiling_max_m", T.RULE_HEAD_TO_CEILING.code),
        _trace("zone_area_max_m2", T.RULE_ZONE_AREA.code),
        _trace("spray_zone_heads_max", T.RULE_SPRAY_ZONE.code),
        _trace("branch_heads_per_side_max", T.RULE_BRANCH_HEADS.code,
               note="분기점 기준 한쪽. 양쪽 합계 16개는 적법"),
        _trace("cross_main_min_dn", T.RULE_CROSS_MAIN.code),
        _trace("tournament_forbidden", T.RULE_TOURNAMENT.code),
        _trace("pipe_size_table", T.RULE_PIPE_SIZE.code),
        _trace("beam_clearance_table", T.RULE_BEAM.code),
        _trace("head_exempt_places", T.RULE_HEAD_EXEMPT.code),
    )

    return Constraints(
        scenario_head_count=n,
        water_supply_m3=volume_m3,
        discharge_minutes=minutes,
        emergency_power_minutes=minutes,
        horizontal_distance_m=r_m,
        head_spacing_square_m=spacing,
        wall_clearance_max_m=spacing / 2.0,
        temp_rating_c=rating_c,
        quick_response_required=quick,
        k_factor=T.K_FACTOR_STANDARD,
        flow_lpm_min=T.FLOW_LPM_MIN,
        pressure_mpa_min=T.PRESSURE_MPA_MIN,
        pressure_mpa_max=T.PRESSURE_MPA_MAX,
        head_clearance_radius_m=T.HEAD_CLEARANCE_RADIUS_M,
        head_to_wall_clearance_m=T.HEAD_TO_WALL_CLEARANCE_M,
        head_to_ceiling_max_m=T.HEAD_TO_CEILING_MAX_M,
        zone_area_max_m2=T.ZONE_AREA_MAX_M2,
        # 격자형 배관방식 면적 완화는 결정 대기다. C5 가 트리 라우팅만 하는 한
        # 격자형이 성립하지 않아 완화의 근거 자체가 없다. 루프 라우팅 모드가
        # 생기기 전까지 None 을 유지한다.
        zone_grid_relief_m2=None,
        zone_floors_max=T.ZONE_FLOORS_MAX,
        spray_zone_heads_max=T.SPRAY_ZONE_HEADS_MAX,
        spray_zone_heads_min_when_split=T.SPRAY_ZONE_HEADS_MIN_WHEN_SPLIT,
        branch_heads_per_side_max=T.BRANCH_HEADS_PER_SIDE_MAX,
        cross_main_min_dn=T.CROSS_MAIN_MIN_DN,
        tournament_forbidden=True,
        pipe_size_table=T.PIPE_SIZE_TABLE,
        velocity_limit_mps=T.VELOCITY_LIMIT_MPS,
        beam_clearance_table=T.BEAM_CLEARANCE,
        head_exempt_places=T.HEAD_EXEMPT_PLACES,
        nftc_effective_date=T.NFTC_EFFECTIVE_DATE,
        trace=trace,
    )


def _trace(field: str, rule_code: str, note: str = "") -> FieldTrace:
    rule = T.RULES[rule_code]
    return FieldTrace(
        field=field, code=rule.code, article=rule.article,
        text_hash=rule.text_hash, effective_date=rule.effective_date, note=note,
    )
