# 파일명 models.py
import math
from dataclasses import dataclass, field, fields
from typing import Any
from enum import Enum

from domain.component_spec_adapter import normalize_node_meta_for_load
from domain.component_specs import (
    ComponentSpec,
    component_spec_from_flat_meta,
    infer_component_type_id_from_flat_meta,
)
# 가속기 셋포인트 SSOT는 계산 코어(domain.fdt.accelerator)다 — 노드 기본도 그대로 참조한다.
from domain.fdt.accelerator import (
    DEFAULT_MECHANICAL_DROP_BAR as FDT_DEFAULT_ACCEL_DROP_BAR,
    DEFAULT_ELECTRONIC_TRIP_S as FDT_DEFAULT_ACCEL_TRIP_S,
)

# === FDT(물 도달시간) 밸브 기본 입력값 — 단일 출처(SSOT) ===
# 건식 밸브는 Node dataclass 기본값(초기공기압 2.4, 차압비 5.5)을 그대로 쓰며 아래 준비작동식
# 전용 상수의 영향을 받지 않는다. 준비작동식 노드에 한해 update_from_meta가 이 상수로 기본값을 채운다.
# UI(다이얼로그)·계산(엔진 요청)은 모두 노드 값을 읽으므로 표시·저장·계산이 항상 일치한다.
# 노드 dataclass 기본값·파싱 폴백·다이얼로그/서비스 폴백이 모두 아래 상수를 참조하므로,
# 여기 한 곳만 바꾸면 표시·저장·계산에 전파된다(한 기본값을 여러 파일에서 고치던 문제 제거).
DRY_DEFAULT_INITIAL_AIR_BAR = 2.4
PREACTION_DEFAULT_INITIAL_AIR_SINGLE_BAR = 1.5   # 싱글인터락 초기(감시) 공기압
PREACTION_DEFAULT_INITIAL_AIR_DOUBLE_BAR = 2.4   # 더블인터락 초기(감시) 공기압
PREACTION_DEFAULT_TRIP_BAR = 1.5                 # 트립압(더블인터락 직접입력 기본)
DRY_DEFAULT_DIFFERENTIAL_RATIO = 5.5             # 건식 차압비(트립 = 1차측 ÷ 차압비)
DRY_DEFAULT_PRIMARY_WATER_BAR = 5.0              # 건식 1차측 수압 [bar, 계기]
# 가스 온도 기본값 [°C] — 노드/입력 UI SSOT. (air_model.DEFAULT_GAS_TEMP_C=5℃는 '온도 미지정 시'
# 계산 폴백이라 의도적으로 다르다. 노드는 항상 이 값을 들고 있으므로 표시·저장·계산은 20℃로 일치.)
FDT_DEFAULT_GAS_TEMP_C = 20.0
# 가속기 ΔP·전자식 트립시간 기본값은 FDT_DEFAULT_ACCEL_*(domain.fdt.accelerator) 참조 — 위 import.
# 옛(공용) 기본값 — 신규 기본 도입 이전에 저장된 준비작동식 노드를 1회 보정하기 위한 센티넬.
# (건식 기본값과 독립 — 건식 기본을 바꿔도 이 값은 과거 준비작동식 2.5를 가리키므로 유지한다.)
_LEGACY_FDT_INITIAL_AIR_BAR = 2.5
_LEGACY_FDT_TRIP_BAR = 0.0

# [1] 노드 타입 표준 정의 (Enum)
class NodeTypeId(str, Enum):
    BASE   = "base"      # 기본
    HEAD   = "head"      # 헤드
    NOZZLE = "nozzle"    # 노즐
    PUMP   = "pump"      # 펌프
    WT     = "wt"        # 수조
    RES    = "reservoir" # 상수원
    PRV    = "prv"       # 감압밸브
    VALVE  = "valve"     # 밸브류
    ORIFICE = "orifice"  # 오리피스
    HOSE_STREAM = "hose_stream"  # 호스스트림 (NFPA 13 외부방출형)

# --- 통합 노드타입 문자열 → NodeTypeId 룩업 ---
# 한글, 영문, 약어 등 모든 알려진 입력값을 커버한다.
# 새 언어 추가 시 이 딕셔너리에 항목만 추가하면 된다.
_NODE_TYPE_LOOKUP: dict[str, str] = {
    # --- 기본 ---
    "base":      "base",
    "기본":      "base",
    # --- 펌프 ---
    "pump":      "pump",
    "펌프":      "pump",
    # --- 수조 ---
    "wt":        "wt",
    "수조":      "wt",
    # --- 상수원 ---
    "reservoir": "reservoir",
    "상수원":    "reservoir",
    "res":       "reservoir",
    # --- 감압밸브 ---
    "prv":       "prv",
    "감압변":    "prv",
    "감압밸브":  "prv",
    # --- 밸브 ---
    "valve":     "valve",
    "밸브":      "valve",
    # --- 오리피스 ---
    "orifice":   "orifice",
    "오리피스":  "orifice",
    # --- 헤드 ---
    "head":      "head",
    "헤드":      "head",
    # --- 노즐 ---
    "nozzle":    "nozzle",
    "노즐":      "nozzle",
    "nz":        "nozzle",
    # --- 호스스트림 ---
    "hose_stream":  "hose_stream",
    "hose stream":  "hose_stream",   # fixed_after "Hose Stream" 소문자화 결과
    "hosestream":   "hose_stream",
    "호스스트림":   "hose_stream",
}

# [2] 타입 정규화 함수 (심플 버전)
def normalize_node_type_id(display_type: str, meta: dict | None = None) -> str:
    """
    복잡한 추측 없이, 확실한 데이터(Meta)가 있을 때만 타입을 인정합니다.

    [DR-1/DR-2 이후 주의]
    to_dict() 는 항상 모든 Node 필드를 포함한다.
    따라서 'rated_q' in meta 같은 key presence 검사는 신뢰할 수 없다.
    → 값(value) 기반 검사로 전환:
      - rated_q, water_level, hole_diameter 등 항상 존재하는 숫자 필드는 > 0 조건
      - fitting_id, head_spec_name, nozzle_name 등 Optional 필드는 truthy 조건

    [Stage 1b 정책 강화]
    명시적 base 표시 또는 meta.type_id="base"는 stale 메타보다 우선한다.
    stale pump_curve_data 가 있어도 base 노드를 pump로 되살리지 않는다.
    """
    meta = meta or {}

    # [신규] 명시적 base 우선 — stale 메타가 base 노드를 다른 타입으로 되살리지 못하게 한다
    _s = str(display_type).lower().strip()
    if _NODE_TYPE_LOOKUP.get(_s) == NodeTypeId.BASE.value:
        return NodeTypeId.BASE.value

    _explicit_tid = str(meta.get("type_id", "")).lower().strip()
    if _explicit_tid == NodeTypeId.BASE.value:
        # display_type이 다른 명시적 비-base 타입을 가리키면 display 우선
        if _s in _NODE_TYPE_LOOKUP and _NODE_TYPE_LOOKUP[_s] != NodeTypeId.BASE.value:
            return _NODE_TYPE_LOOKUP[_s]
        return NodeTypeId.BASE.value

    inferred_from_meta = infer_component_type_id_from_flat_meta(meta)
    if inferred_from_meta:
        return inferred_from_meta

    # 2. 통합 룩업 (한글/영문/약어 → NodeTypeId)
    s = str(display_type).lower().strip()
    looked_up = _NODE_TYPE_LOOKUP.get(s)
    if looked_up:
        return looked_up

    # 3. 부분 매칭 Fallback (스펙명 등 룩업에 없는 문자열 대응)
    if "nozzle" in s or "nz" in s:
        return NodeTypeId.NOZZLE.value
    if "head" in s:
        return NodeTypeId.HEAD.value

    # 4. 모르면 기본값
    return NodeTypeId.BASE.value

# [3] 배관 클래스
@dataclass
class Pipe:
    """배관 도메인 모델"""
    id: str
    start: str
    end: str
    
    # 속성
    type: str = ""
    diameter: float = 0.0
    nominal_mm: float = 0.0
    length_m: float = 0.0
    equivalent_length: float = 0.0
    C: float = 120.0
    roughness_mm: float = 0.085
    fittings: list[str] = field(default_factory=list)
    
    # 결과값
    flow_lpm: float = 0.0
    velocity_mps: float = 0.0
    headloss_m: float = 0.0

    def to_dict(self) -> dict[str, Any]:
        return {
            "start": self.start,
            "end": self.end,
            "type": self.type,
            "diameter": self.diameter,
            "nominal_mm": self.nominal_mm,
            "length_m": self.length_m,
            "equivalent_length": self.equivalent_length,
            "C": self.C,
            "roughness_mm": self.roughness_mm,
            "fittings": self.fittings,
            "flow_lpm": self.flow_lpm,
            "velocity_mps": self.velocity_mps,
            "headloss_m": self.headloss_m
        }
    
    def to_legacy_dict(self) -> dict[str, Any]:
        """하위 호환성을 위해 유지 (editor_core.py에서 사용)"""
        return self.to_dict()

    @classmethod
    def from_dict(cls, pipe_id: str, data: dict[str, Any]) -> 'Pipe':
        """
        [로드용 팩토리 메서드]
        JSON 데이터를 읽어 Pipe 객체를 생성합니다.
        기본값은 Pipe 클래스 필드 정의와 동일합니다 (SSOT).
        """
        return cls(
            id=pipe_id,
            start=data.get('start', ''),
            end=data.get('end', ''),
            type=data.get('type', ''),
            diameter=data.get('diameter', 0.0),
            nominal_mm=data.get('nominal_mm', 0.0),
            length_m=data.get('length_m', 0.0),
            equivalent_length=data.get('equivalent_length', 0.0),
            C=data.get('C', 120.0),
            roughness_mm=data.get('roughness_mm', 0.085),
            fittings=data.get('fittings', []),
            flow_lpm=data.get('flow_lpm', 0.0),
            velocity_mps=data.get('velocity_mps', 0.0),
            headloss_m=data.get('headloss_m', 0.0),
        )

# [4] 수리 계산 모델 클래스
@dataclass
class HydraulicModel:
    """수력학 해석 결과 모델"""
    pipe_results: dict[str, Any] = field(default_factory=dict)
    node_results: dict[str, Any] = field(default_factory=dict)
    summary: dict[str, Any] = field(default_factory=dict)
    node_meta_patches: list[Any] = field(default_factory=list)
        
    def to_dict(self) -> dict[str, Any]:
        return {
            "pipe_results": self.pipe_results,
            "node_results": self.node_results,
            "summary": self.summary
        }
        
    @staticmethod
    def from_dict(data: dict[str, Any]) -> 'HydraulicModel':
        return HydraulicModel(
            pipe_results=data.get("pipe_results", {}),
            node_results=data.get("node_results", {}),
            summary=data.get("summary", {})
        )
    
# [5-0] 안전 형변환 헬퍼 (update_from_meta 전용 — 모듈 외부 사용 금지)
def _safe_float(v: Any, default: float = 0.0) -> float:
    """None·빈 문자열·잘못된 타입을 안전하게 float으로 변환합니다."""
    if v is None or v == "":
        return default
    try:
        return float(v)
    except (TypeError, ValueError):
        return default

def _safe_bool(v: Any, default: bool = False) -> bool:
    """None·잘못된 타입을 안전하게 bool로 변환합니다."""
    if isinstance(v, bool):
        return v
    if v is None:
        return default
    try:
        return bool(v)
    except (TypeError, ValueError):
        return default

def _safe_str(v: Any, default: str = "") -> str:
    """None을 안전하게 str로 변환합니다. 빈 문자열은 default로 대체하지 않습니다."""
    if v is None:
        return default
    return str(v)

# [5] 노드 클래스
@dataclass
class Node:
    """
    [통합 데이터 모델]
    기존 분산 노드 저장 구조를 단일 객체로 통합한 모델 클래스.
    """
    # 1. 기본 식별 정보 (필수)
    id: str
    coords: tuple[float, float, float]
    elevation_m: float
    
    # 2. 타입 정보
    type: str = "기본"                          # UI 표시 이름 (구 node_types 값)
    type_id: str = field(default_factory=lambda: NodeTypeId.BASE.value)  # 내부 로직용 ID
    category_id: str = ""                       # 라이브러리 카테고리 ID (예: "head", "water_spray", "alarm_valve")
    
    # 3. 상태 제어 (Status)
    is_active: bool = True                      # 헤드/노즐 개방 여부
    is_closed: bool = False                     # 밸브 잠금 여부
    
    # 4. 수리 계산 공통 속성 (Hydraulic Props)
    k_factor_si: float | None = None            # K-Factor (헤드/노즐)
    required_pressure_bar: float = 0.0          # 최소 방수압력
    base_demand_lps: float = 0.0                # 고정 유량 수요
    
    # 5. 장치별 상세 참조 (Reference IDs)
    fitting_id: str | None = None               # 밸브 라이브러리 ID
    head_spec_name: str | None = None           # 헤드 템플릿 이름
    nozzle_name: str | None = None              # 노즐 라이브러리 이름
    pump_library_id: str | None = None          # 펌프 라이브러리 ID
    
    # 6. 펌프/수조 속성 (Pump/Tank Props)
    rated_q: float = 0.0                        # 펌프 정격 유량
    rated_p: float = 0.0                        # 펌프 정격 양정
    shutoff_p: float = 0.0                      # 펌프 체절 압력
    peak_q: float = 0.0                         # 펌프 최대 유량
    peak_p: float = 0.0                         # 펌프 최대 압력
    water_level: float = 0.0                    # 수조 초기 수위
    
    # 7. 기타 장치 속성
    pressure_setting_bar: float | None = None   # PRV/감압밸브 설정압력
    loss_coefficient: float = 0.0               # 손실 계수 (Minor Loss K)
    hole_diameter: float = 0.0                  # 오리피스 구경
    orifice_discharge_coeff: float = 0.65       # 오리피스 유출계수
    is_check_valve: bool = False                # 밸브 타입 — 체크밸브 기능 여부

    # 10. 호스스트림 전용 속성 (NFPA 13)
    hose_stream_flow_lps: float = 400.0 / 60.0  # 기본값 400 L/min → L/s 내부 저장

    # 8. 과거 메타 딕셔너리에서 승격된 필드 (DR-1)
    pipe_dn: float = 0.0                        # 오리피스 연결 배관 DN (mm)
    check_valve_direction: str = ""             # 체크밸브 업스트림 노드명 (밸브 타입 전용)
    has_check_valve: bool = True                # 수조 출구 역류방지 여부 (WT 타입 전용)
    pump_curve_data: list = field(default_factory=list)  # 커스텀 펌프 커브 데이터

    # 11. FDT(물 도달시간) — 건식/준비작동식 밸브 전용 입력(UI에서 채움, 엔진은 요청으로 받음)
    fdt_differential_ratio: float = DRY_DEFAULT_DIFFERENTIAL_RATIO  # 차압비(건식 트립 = 1차측 ÷ 차압비)
    fdt_primary_water_gauge_bar: float = DRY_DEFAULT_PRIMARY_WATER_BAR  # 1차측 수압 [bar, 계기]
    fdt_p_initial_gauge_bar: float = DRY_DEFAULT_INITIAL_AIR_BAR  # 초기(감시) 공기압 [bar, 계기] (건식 기본; 준비작동식은 SSOT 상수로 보정)
    fdt_gas: str = "air"                         # 가스 종류 'air' | 'nitrogen'
    fdt_temp_c: float = FDT_DEFAULT_GAS_TEMP_C   # 가스 온도 [°C]
    fdt_trip_pressure_gauge_bar: float = 0.0     # 트립압 직접입력 [bar, 계기](준비작동식은 SSOT 상수로 보정)
    fdt_interlock_mode: str = "single"           # 준비작동식 인터락 방식 'double' | 'single' | 'non'
    # 가속기(QOD) — 건식밸브 전용. 신규 밸브는 '설치'가 기본(현장 통례). 엔진은 요청으로 받음.
    fdt_accel_installed: bool = True             # 가속기 설치 여부(기본 설치)
    fdt_accel_kind: str = "mechanical"           # 가속기 유형 'mechanical' | 'electronic'(기본 기계식)
    fdt_accel_drop_bar: float = FDT_DEFAULT_ACCEL_DROP_BAR   # 기계식 트립 공기압 강하 ΔP [bar](≈4.4psi)
    fdt_accel_trip_time_s: float = FDT_DEFAULT_ACCEL_TRIP_S  # 전자식 고정 트립시간 [s]
    # 물도달시간 헤드 선택 기억 — 선택 모드에서 고른 개방 헤드 id 목록. 도면의 is_active(수리계산)와
    # 독립인 FDT 전용 상태라 밸브 노드에 저장한다(파일과 함께 영속). 삭제된 헤드는 사용 시점에 걸러낸다.
    fdt_selected_head_ids: list = field(default_factory=list)

    # 9. 저장되지 않는 런타임 결과값 (Optional)
    _runtime_result_p: float = field(init=False, default=0.0)
    _runtime_result_q: float = field(init=False, default=0.0)
    
    @classmethod
    def create(cls, name: str, x: float, y: float, z: float, node_type: str = "기본") -> 'Node':
        """기존 __init__(name, x, y, z, node_type) 시그니처와 호환되는 팩토리 메서드"""
        return cls(
            id=str(name),
            coords=(float(x), float(y), float(z)),
            elevation_m=float(z),
            type=str(node_type)
        )

    def update_from_meta(self, meta: dict):
        """
        [마이그레이션 헬퍼]
        과거 메타 딕셔너리의 내용을 이 객체 속성으로 일괄 업데이트합니다.

        안전 형변환: 레거시 딕셔너리에는 UI 경유로 빈 문자열·None이 섞일 수 있으므로
        모든 float/bool/str 변환에 _safe_* 헬퍼를 사용합니다.
        """
        if not meta:
            return
        meta = normalize_node_meta_for_load(meta)

        # --- 타입 정보 ---
        if "type_id" in meta:
            self.type_id = _safe_str(meta["type_id"], default=NodeTypeId.BASE.value) or NodeTypeId.BASE.value
        if "category_id" in meta:
            self.category_id = _safe_str(meta["category_id"], default="")

        # --- 상태 제어 ---
        if "is_active" in meta:
            self.is_active = _safe_bool(meta["is_active"], default=True)
        if "is_closed" in meta:
            self.is_closed = _safe_bool(meta["is_closed"], default=False)
        if "is_check_valve" in meta:
            self.is_check_valve = _safe_bool(meta["is_check_valve"], default=False)

        # --- 수리 계산 공통 ---
        # k_factor_si 우선, k_factor는 구 버전 호환 별칭
        if "k_factor_si" in meta and meta["k_factor_si"] is not None:
            self.k_factor_si = _safe_float(meta["k_factor_si"])
        elif "k_factor" in meta and meta["k_factor"] is not None:
            self.k_factor_si = _safe_float(meta["k_factor"])

        if "required_pressure_bar" in meta:
            self.required_pressure_bar = _safe_float(meta["required_pressure_bar"])
        if "base_demand_lps" in meta:
            self.base_demand_lps = _safe_float(meta["base_demand_lps"])

        # --- 장치별 참조 ID ---
        if "fitting_id" in meta:
            self.fitting_id = meta["fitting_id"]  # str | None — 그대로 유지
        if "head_spec_name" in meta:
            self.head_spec_name = meta["head_spec_name"]
        if "nozzle_name" in meta:
            self.nozzle_name = meta["nozzle_name"]
        if "pump_library_id" in meta:
            self.pump_library_id = meta.get("pump_library_id")

        # --- 펌프/수조 속성 ---
        if "rated_q" in meta:
            self.rated_q = _safe_float(meta["rated_q"])
        if "rated_p" in meta:
            self.rated_p = _safe_float(meta["rated_p"])
        if "shutoff_p" in meta:
            self.shutoff_p = _safe_float(meta["shutoff_p"])
        if "peak_q" in meta:
            self.peak_q = _safe_float(meta["peak_q"])
        if "peak_p" in meta:
            self.peak_p = _safe_float(meta["peak_p"])
        if "water_level" in meta:
            self.water_level = _safe_float(meta["water_level"])
        if "pump_curve_data" in meta and isinstance(meta["pump_curve_data"], list):
            self.pump_curve_data = list(meta["pump_curve_data"])

        # --- PRV / 손실 ---
        if "pressure_setting_bar" in meta and meta["pressure_setting_bar"] is not None:
            self.pressure_setting_bar = _safe_float(meta["pressure_setting_bar"])
        if "loss_coefficient" in meta:
            self.loss_coefficient = _safe_float(meta["loss_coefficient"])

        # --- 오리피스 속성 ---
        if "hole_diameter" in meta:
            self.hole_diameter = _safe_float(meta["hole_diameter"])
        if "orifice_discharge_coeff" in meta:
            self.orifice_discharge_coeff = _safe_float(meta["orifice_discharge_coeff"])
        if "pipe_dn" in meta:
            self.pipe_dn = _safe_float(meta["pipe_dn"])

        # --- 체크밸브 방향 (밸브 타입 전용) ---
        if "check_valve_direction" in meta:
            self.check_valve_direction = _safe_str(meta["check_valve_direction"])

        # --- 수조 체크밸브 (WT 타입 전용) ---
        if "has_check_valve" in meta:
            self.has_check_valve = _safe_bool(meta["has_check_valve"], default=True)

        # --- 호스스트림 전용 유량 (HOSE_STREAM 타입 전용) ---
        if "hose_stream_flow_lps" in meta:
            self.hose_stream_flow_lps = _safe_float(meta["hose_stream_flow_lps"], default=400.0 / 60.0)

        # --- FDT 밸브 설정(건식/준비작동식 전용) ---
        if "fdt_differential_ratio" in meta:
            self.fdt_differential_ratio = _safe_float(meta["fdt_differential_ratio"], default=DRY_DEFAULT_DIFFERENTIAL_RATIO)
        if "fdt_primary_water_gauge_bar" in meta:
            self.fdt_primary_water_gauge_bar = _safe_float(meta["fdt_primary_water_gauge_bar"], default=DRY_DEFAULT_PRIMARY_WATER_BAR)
        if "fdt_p_initial_gauge_bar" in meta:
            self.fdt_p_initial_gauge_bar = _safe_float(meta["fdt_p_initial_gauge_bar"], default=DRY_DEFAULT_INITIAL_AIR_BAR)
        if "fdt_gas" in meta:
            self.fdt_gas = _safe_str(meta["fdt_gas"], default="air")
        if "fdt_temp_c" in meta:
            self.fdt_temp_c = _safe_float(meta["fdt_temp_c"], default=FDT_DEFAULT_GAS_TEMP_C)
        if "fdt_trip_pressure_gauge_bar" in meta:
            self.fdt_trip_pressure_gauge_bar = _safe_float(meta["fdt_trip_pressure_gauge_bar"], default=0.0)
        if "fdt_interlock_mode" in meta:
            self.fdt_interlock_mode = _safe_str(meta["fdt_interlock_mode"], default="single")
        if self.category_id == "preaction_valve":
            self.apply_preaction_fdt_defaults(meta)
        if "fdt_accel_installed" in meta:
            self.fdt_accel_installed = _safe_bool(meta["fdt_accel_installed"], default=True)
        if "fdt_accel_kind" in meta:
            self.fdt_accel_kind = _safe_str(meta["fdt_accel_kind"], default="mechanical")
        if "fdt_accel_drop_bar" in meta:
            self.fdt_accel_drop_bar = _safe_float(meta["fdt_accel_drop_bar"], default=FDT_DEFAULT_ACCEL_DROP_BAR)
        if "fdt_accel_trip_time_s" in meta:
            self.fdt_accel_trip_time_s = _safe_float(meta["fdt_accel_trip_time_s"], default=FDT_DEFAULT_ACCEL_TRIP_S)
        if "fdt_selected_head_ids" in meta and isinstance(meta["fdt_selected_head_ids"], list):
            self.fdt_selected_head_ids = [str(h) for h in meta["fdt_selected_head_ids"]]

    def apply_preaction_fdt_defaults(self, meta: dict) -> None:
        """준비작동식 밸브 FDT 기본값을 단일 지점에서 적용한다(SSOT).

        - 신규 노드(meta에 키 없음): 모드별 기본값을 채운다.
        - 레거시 노드(신규 기본 도입 이전의 옛 공용 기본값이 저장됨): 신규 기본값으로 보정한다.
          옛 기본값(2.5/0.0)과 값이 정확히 같을 때만 바꾸므로 사용자가 입력한 다른 값은 보존된다.

        표시(다이얼로그)·저장·계산(엔진 요청)이 모두 이 결과(노드 값)를 읽으므로, 화면에 보이는
        값과 실제 계산에 쓰이는 값이 항상 일치한다(표시 전용 보정으로 인한 불일치 없음).

        파일 로드(update_from_meta)와 앱 내 편집(editor.update_node_via_graph) 두 경로 모두
        이 메서드를 호출해 동일 규칙을 적용한다.
        """
        mode = str(self.fdt_interlock_mode or "single")
        # 초기(감시) 공기압 — 싱글 1.5 / 더블은 SSOT 상수.
        # 싱글은 '미입력' 또는 '일반(건식/공용) 기본값이 새어든 경우'를 1.5로 보정한다.
        #   · 옛 공용 기본값(_LEGACY 2.5): 신규 기본 도입 이전 저장 노드.
        #   · 현행 건식/공용 기본값(DRY_DEFAULT 2.4): category_id가 부속과 어긋나 도메인 보정을
        #     거치지 못한 채 dataclass 기본값(2.4)이 싱글 준비작동식에 남은 경우.
        # 두 센티넬과 정확히 같을 때만 바꾸므로 사용자가 직접 넣은 다른 값은 보존된다.
        if mode == "double":
            if "fdt_p_initial_gauge_bar" not in meta:
                self.fdt_p_initial_gauge_bar = PREACTION_DEFAULT_INITIAL_AIR_DOUBLE_BAR
        elif ("fdt_p_initial_gauge_bar" not in meta
              or math.isclose(self.fdt_p_initial_gauge_bar, _LEGACY_FDT_INITIAL_AIR_BAR,
                              rel_tol=0.0, abs_tol=1e-9)
              or math.isclose(self.fdt_p_initial_gauge_bar, DRY_DEFAULT_INITIAL_AIR_BAR,
                              rel_tol=0.0, abs_tol=1e-9)):
            self.fdt_p_initial_gauge_bar = PREACTION_DEFAULT_INITIAL_AIR_SINGLE_BAR
        # 트립압 — 미입력 또는 옛 기본(0.0)이면 1.5(더블인터락 직접입력 기본).
        if ("fdt_trip_pressure_gauge_bar" not in meta
                or math.isclose(self.fdt_trip_pressure_gauge_bar, _LEGACY_FDT_TRIP_BAR,
                                rel_tol=0.0, abs_tol=1e-9)):
            self.fdt_trip_pressure_gauge_bar = PREACTION_DEFAULT_TRIP_BAR

    def to_dict(self) -> dict[str, Any]:
        """저장용 딕셔너리 변환 — dataclasses.fields() 자동 순회"""
        result = {}
        for f in fields(self):
            # 런타임 전용 필드 제외
            if f.name.startswith("_runtime_"):
                continue
            val = getattr(self, f.name)
            # coords는 tuple → list 변환 (JSON 호환)
            if f.name == "coords":
                val = list(val)
            result[f.name] = val
        return result

    @property
    def component_spec(self) -> ComponentSpec:
        """기존 flat 필드에서 만든 읽기 전용 ComponentSpec view."""
        return self.to_component_spec()

    def to_component_spec(self) -> ComponentSpec:
        """Return a ComponentSpec view without mutating or changing serialization."""
        return component_spec_from_flat_meta(self.to_dict())

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> 'Node':
        """
        [로드용 팩토리 메서드] 
        JSON 데이터를 읽어 Node 객체를 생성합니다.
        """
        name = data.get("id", "Unknown")
        coords_list = data.get("coords", [0, 0, 0])
        coords = tuple(coords_list) if isinstance(coords_list, list) else coords_list
        node_type = data.get("type", "기본")

        # 객체 생성 (@dataclass 생성자 사용)
        node = cls(
            id=name,
            coords=coords,
            elevation_m=coords[2],
            type=node_type
        )

        # 딕셔너리 데이터로 속성 채우기 (update_from_meta 재활용)
        node.update_from_meta(data)
        
        # type_id 별도 처리 (update_from_meta에 포함되어 있지만 확실히 하기 위해)
        if "type_id" in data:
            node.type_id = data["type_id"]
            
        return node
