"""변환 폼 ↔ 엔진 계약 SSOT. 입력 값만. Z 계산 없음."""
import os

BRANCH_DEFAULT_M = 0.3
UPRIGHT_DEFAULT_M = 0.3
PENDANT_DEFAULT_M = 0.3
PENDANT_2_DEFAULT_M = 0.3
COMBO_1_DEFAULT_M = 0.2
COMBO_2_DEFAULT_M = 0.3
COMBO_3_DEFAULT_M = 0.5
COMBO_UP_DEFAULT_M = 0.3
HEAD_K_DEFAULT = 80.0
VALVE_1_DEFAULT_M = 2.5
VALVE_2_DEFAULT_M = 0.5

# (dto 키, 라벨, 칸 기본 문자열, 빈칸일 때 값)
BRANCH_FIELDS = (
    ("branch_rise_m", "메인→가지 수직 (m)", "0.3", BRANCH_DEFAULT_M),
)
UPRIGHT_FIELDS = (
    ("upright_1_m", "① (m)", "0.3", UPRIGHT_DEFAULT_M),
)
PENDANT_FIELDS = (
    ("pendant_1_m", "① (m)", "0.3", PENDANT_DEFAULT_M),
    ("pendant_2_m", "② (m)", "0.3", PENDANT_2_DEFAULT_M),
)
COMBO_FIELDS = (
    ("combo_1_m", "① (m)", "0.2", COMBO_1_DEFAULT_M),
    ("combo_up_m", "② (m)", "0.3", COMBO_UP_DEFAULT_M),
    ("combo_2_m", "③ (m)", "0.3", COMBO_2_DEFAULT_M),
    ("combo_3_m", "④ (m)", "0.5", COMBO_3_DEFAULT_M),
)
FLEX_FIELDS = (
    ("flex_c", "후렉시블 조도C (하향 2, 상하향 4)", "", None),
    ("flex_roughness_mm", "후렉시블 거칠기 (mm)", "", None),
)
SHARED_FIELDS = (
    ("head_k", "헤드 K", "80", HEAD_K_DEFAULT),
)
VALVE_FIELDS = (
    ("valve_1_m", "① (m)", "2.5", VALVE_1_DEFAULT_M),
    ("valve_2_m", "② (m)", "0.5", VALVE_2_DEFAULT_M),
)
FIELDS = (
    BRANCH_FIELDS + UPRIGHT_FIELDS + PENDANT_FIELDS + COMBO_FIELDS
    + FLEX_FIELDS + SHARED_FIELDS + VALVE_FIELDS
)


def default_dto():
    """폼 기본값. 길이·K·밸브 칸은 비우지 않는다. C·거칠기는 체크 꺼짐이면 None."""
    return {
        "branch_rise_m": BRANCH_DEFAULT_M,
        "upright_1_m": UPRIGHT_DEFAULT_M,
        "pendant_1_m": PENDANT_DEFAULT_M,
        "pendant_2_m": PENDANT_2_DEFAULT_M,
        "combo_1_m": COMBO_1_DEFAULT_M,
        "combo_2_m": COMBO_2_DEFAULT_M,
        "combo_3_m": COMBO_3_DEFAULT_M,
        "combo_up_m": COMBO_UP_DEFAULT_M,
        "flex_c": None,
        "flex_roughness_mm": None,
        "head_k": HEAD_K_DEFAULT,
        "head_active": False,
        "min_pressure_bar": None,
        "valve_1_m": VALVE_1_DEFAULT_M,
        "valve_2_m": VALVE_2_DEFAULT_M,
    }


def dto_to_convert_kwargs(dto):
    """폼 DTO 그대로 넘긴다. Z 계산은 엔진.

    pendant_1_m 은 하향 ① (0 이면 중간이 가지에 직접). ② 는 pendant_2_m.
    """
    return dict(dto or {})


def default_out_path(key):
    return os.path.join(os.path.expanduser("~"), "Desktop",
                        f"{key}_변환.kfp")
