"""
domain/constants.py

K-Fire 도메인 계층 — 배관 규격 상수 데이터
PySide6 의존성 없음.
"""
from __future__ import annotations

from enum import Enum

# ---------------------------------------------------------------------------
# (1) 배관 규격 참조표 (KSD3507 기준 내경)
# ---------------------------------------------------------------------------
# K-Factor(손실계수)는 소방 공학 추정치.
# 추후 제조사 카탈로그 값으로 수정 가능.

PIPE_DIAMETER_DN: dict[int, dict] = {
    20:  {"id": "DN20",  "D_mm": 21.9,  "K_CheckValve": 2.5,  "K_AlarmV": 5.0,  "K_DryV": 5.0,  "K_PreactionV": 5.0},
    25:  {"id": "DN25",  "D_mm": 27.5,  "K_CheckValve": 2.5,  "K_AlarmV": 5.0,  "K_DryV": 5.0,  "K_PreactionV": 5.0},
    32:  {"id": "DN32",  "D_mm": 36.2,  "K_CheckValve": 2.5,  "K_AlarmV": 5.0,  "K_DryV": 5.0,  "K_PreactionV": 5.0},
    40:  {"id": "DN40",  "D_mm": 42.1,  "K_CheckValve": 3.0,  "K_AlarmV": 6.0,  "K_DryV": 6.0,  "K_PreactionV": 6.0},
    50:  {"id": "DN50",  "D_mm": 53.2,  "K_CheckValve": 3.8,  "K_AlarmV": 8.0,  "K_DryV": 8.0,  "K_PreactionV": 8.0},
    65:  {"id": "DN65",  "D_mm": 69.0,  "K_CheckValve": 5.0,  "K_AlarmV": 10.0, "K_DryV": 10.0, "K_PreactionV": 10.0},
    80:  {"id": "DN80",  "D_mm": 81.0,  "K_CheckValve": 6.5,  "K_AlarmV": 12.0, "K_DryV": 12.0, "K_PreactionV": 12.0},
    100: {"id": "DN100", "D_mm": 105.3, "K_CheckValve": 9.0,  "K_AlarmV": 15.0, "K_DryV": 15.0, "K_PreactionV": 15.0},
    125: {"id": "DN125", "D_mm": 130.1, "K_CheckValve": 12.0, "K_AlarmV": 18.0, "K_DryV": 18.0, "K_PreactionV": 18.0},
    150: {"id": "DN150", "D_mm": 155.5, "K_CheckValve": 15.0, "K_AlarmV": 22.0, "K_DryV": 22.0, "K_PreactionV": 22.0},
    200: {"id": "DN200", "D_mm": 204.6, "K_CheckValve": 22.0, "K_AlarmV": 30.0, "K_DryV": 30.0, "K_PreactionV": 30.0},
    250: {"id": "DN250", "D_mm": 254.6, "K_CheckValve": 35.0, "K_AlarmV": 45.0, "K_DryV": 45.0, "K_PreactionV": 45.0},
    300: {"id": "DN300", "D_mm": 304.5, "K_CheckValve": 50.0, "K_AlarmV": 60.0, "K_DryV": 60.0, "K_PreactionV": 60.0},
}


# ---------------------------------------------------------------------------
# (2) v3.0 — NFPA 13 호칭경 도메인 (ADR-001 §Domain)
# ---------------------------------------------------------------------------
# DN90(3½")은 KS 표준 미존재 및 실무 희소로 제외 (ADR-001 Rev 3)
# DN350·DN400은 국내 현장 수요 반영 추가

DN_STANDARD: list[int] = [
    15, 20, 25, 32, 40, 50,
    65, 80, 100, 125, 150,
    200, 250, 300, 350, 400,
]


# ---------------------------------------------------------------------------
# (3) v3.0 — 부속 카테고리 열거형 (ADR-001 §Detailed Design)
# ---------------------------------------------------------------------------
# M1.2 기본 15종 + M1.2-A 하이브리드 5종 = 20종
# NFPA 13 + KFI 명문 등재 + 국내 주요 제조사 카탈로그 수록 부속은 enum 확장으로 추가.
# 진짜 현장 특수 부속만 USER_DEFINED + UUID 키잉.

class FittingCategory(str, Enum):
    """부속·밸브 카테고리 열거형. str 상속으로 직렬화 호환."""
    ELBOW_45            = "elbow_45"
    ELBOW_90_STANDARD   = "elbow_90_standard"
    ELBOW_90_LONG_TURN  = "elbow_90_long"
    TEE_RUN             = "tee_run"           # 직류 — 항상 0, 경고 없음
    TEE_BRANCH_90       = "tee_branch_90"
    CROSS_BRANCH_90     = "cross_branch_90"
    BUTTERFLY_VALVE     = "butterfly_valve"
    GATE_VALVE          = "gate_valve"
    CHECK_VALVE_SWING   = "swing_check"
    ALARM_VALVE         = "alarm_valve"
    DRY_PIPE_VALVE      = "dry_pipe_valve"
    DELUGE_VALVE        = "deluge_valve"
    FLOW_SWITCH_VANE    = "flow_switch_vane"
    STRAINER            = "strainer"
    USER_DEFINED        = "user_defined"
    # M1.2-A 하이브리드 — NFPA 표 외이나 국내·국제 실무 표준 부속
    PREACTION_VALVE     = "preaction_valve"
    SMOLENSKY_CHECK     = "smolensky_check"
    GLOBE_VALVE         = "globe_valve"
    ANGLE_VALVE         = "angle_valve"
    BALL_VALVE          = "ball_valve"


# ---------------------------------------------------------------------------
# (4) v3.0 — type_id 문자열 → FittingCategory 매핑 (M1.3)
# ---------------------------------------------------------------------------

TYPE_ID_TO_CATEGORY: dict[str, FittingCategory] = {
    "elbow_45":             FittingCategory.ELBOW_45,
    "elbow_90_standard":    FittingCategory.ELBOW_90_STANDARD,
    "elbow_90_long":        FittingCategory.ELBOW_90_LONG_TURN,
    "tee_run":              FittingCategory.TEE_RUN,
    "tee_branch_90":        FittingCategory.TEE_BRANCH_90,
    "cross_branch_90":      FittingCategory.CROSS_BRANCH_90,
    "butterfly_valve":      FittingCategory.BUTTERFLY_VALVE,
    "gate_valve":           FittingCategory.GATE_VALVE,
    "swing_check":          FittingCategory.CHECK_VALVE_SWING,
    "alarm_valve":          FittingCategory.ALARM_VALVE,
    "dry_pipe_valve":       FittingCategory.DRY_PIPE_VALVE,
    "deluge_valve":         FittingCategory.DELUGE_VALVE,
    "flow_switch_vane":     FittingCategory.FLOW_SWITCH_VANE,
    "strainer":             FittingCategory.STRAINER,
    "preaction_valve":      FittingCategory.PREACTION_VALVE,
    "smolensky_check":      FittingCategory.SMOLENSKY_CHECK,
    "globe_valve":          FittingCategory.GLOBE_VALVE,
    "angle_valve":          FittingCategory.ANGLE_VALVE,
    "ball_valve":           FittingCategory.BALL_VALVE,
}


def is_tee_run(category: FittingCategory) -> bool:
    """직류 티 여부 판정. 반드시 category 필드 기반 — ID 부재로 판정 금지."""
    return category is FittingCategory.TEE_RUN


# ---------------------------------------------------------------------------
# (4b) 체크밸브 카테고리 SSOT
# ---------------------------------------------------------------------------
# fitting_id 문자열(`"CHECK" in id`)로 판정하지 않는다.
# 유저 추가 부속은 UUID hex id라 ID 휴리스틱이 실패한다.
# 새 체크 카테고리 추가 시 이 frozenset만 확장한다.

CHECK_VALVE_CATEGORIES: frozenset[FittingCategory] = frozenset({
    FittingCategory.CHECK_VALVE_SWING,
    FittingCategory.SMOLENSKY_CHECK,
})


def is_check_valve_category(category: FittingCategory) -> bool:
    """체크밸브 계열 여부. 반드시 category 필드 기반 — ID 부재로 판정 금지."""
    return category in CHECK_VALVE_CATEGORIES


def fitting_category_from_id(cat_id: str) -> FittingCategory | None:
    """Normalize category_id to a known fitting category, or return None."""
    key = str(cat_id or "").strip().lower().replace(" ", "_")
    mapped = TYPE_ID_TO_CATEGORY.get(key)
    if mapped is not None:
        return mapped
    try:
        return FittingCategory(key)
    except ValueError:
        return None


# ---------------------------------------------------------------------------
# (5) v3.0 — 자동 검출 부속 ID 상수 (H2 — ADR-003 §_apply_fitting)
# ---------------------------------------------------------------------------
# flow_analyzer.py의 하드코딩 문자열 참조 대신 이 상수를 사용한다.
# 새 자동 검출 부속 추가 시 이 dict만 확장 (확장점).

AUTO_DETECTED_FITTING_IDS: dict[str, str] = {
    "elbow_90":  "ELBOW_90_STD",
    "elbow_45":  "ELBOW_45",
    "tee_branch": "TEE_BRANCH",
}


# ---------------------------------------------------------------------------
# (6) v3.0 — 셀 출처 열거형 (H3 — ADR-001 §Detailed Design)
# ---------------------------------------------------------------------------
# str 상속으로 `cell.origin == "USER_FILLED"` 패턴과 하위 호환.
# 단, 신규 코드에서는 반드시 enum 멤버 비교 사용 (문자열 리터럴 금지).

class CellOrigin(str, Enum):
    """CellValue.origin 열거형. 오타 silent bug 방지."""
    NFPA_ORIGINAL = "NFPA_ORIGINAL"   # NFPA 시드값 그대로
    USER_OVERRIDE = "USER_OVERRIDE"   # NFPA 시드값을 사용자가 수정한 셀
    USER_FILLED   = "USER_FILLED"     # 원래 null이었으나 사용자가 채운 셀


# ---------------------------------------------------------------------------
# (7) v3.0 — 노드 배치 가능 밸브 카테고리 SSOT (D2)
# ---------------------------------------------------------------------------
# 노드 우클릭 메뉴 "Change Node Type" + 라이브러리 추가 다이얼로그 카테고리 콤보가 공유한다.
# 표시명 dict는 두지 않는다 (D6 — 4계층 분리: domain 계층은 i18n 의존 금지).
# 표시명/i18n 처리는 services 또는 ui 계층에서 한다.
#
# 정렬 순서 = 노드 메뉴 표시 순서 = 라이브러리 카테고리 컬럼 정렬 우선순위 (D8)
# 방재밸브 → 체크밸브 → 개폐밸브 → 특수

NODE_PLACEABLE_VALVE_CATEGORIES: tuple[FittingCategory, ...] = (
    FittingCategory.ALARM_VALVE,
    FittingCategory.FLOW_SWITCH_VANE,
    FittingCategory.PREACTION_VALVE,
    FittingCategory.DRY_PIPE_VALVE,
    FittingCategory.CHECK_VALVE_SWING,
    FittingCategory.SMOLENSKY_CHECK,
    FittingCategory.GATE_VALVE,
    FittingCategory.BUTTERFLY_VALVE,
    FittingCategory.GLOBE_VALVE,
    FittingCategory.ANGLE_VALVE,
    FittingCategory.BALL_VALVE,
    FittingCategory.STRAINER,
)


# ---------------------------------------------------------------------------
# __all__
# ---------------------------------------------------------------------------

__all__ = [
    "PIPE_DIAMETER_DN",
    "DN_STANDARD",
    "FittingCategory",
    "TYPE_ID_TO_CATEGORY",
    "is_tee_run",
    "CHECK_VALVE_CATEGORIES",
    "is_check_valve_category",
    "fitting_category_from_id",
    "AUTO_DETECTED_FITTING_IDS",
    "CellOrigin",
    "NODE_PLACEABLE_VALVE_CATEGORIES",
]
