"""domain/fdt/air_model.py

K-Fire FDT 모듈 — 공기 거동 물리 코어 (Phase A 블로다운 + Phase B 충수 공용).

건식/준비작동식 스프링클러의 압축공기 오리피스 배출을 다루는 순수 함수 모음이다.
명세서 §4.2~4.4(오리피스 배출식·가스법칙·초크 판정)·상수요약을 따른다.

설계 원칙:
- 순수 Python. PySide6·WNTR·EPANET 의존 없음(계획서 §6 4계층 규약).
- 단위는 SI 기본(Pa, m², kg/s, K). 입출력 경계에서만 bar·mm·°C로 환산한다.
- 가스(공기/질소)는 기체상수 Rs만 다르고 비열비 k=1.40 공통(둘 다 2원자 기체).
"""
from __future__ import annotations

import math

__all__ = [
    "P_ATM_PA",
    "KAPPA",
    "RHO_WATER",
    "GAS_CONSTANTS",
    "GAS_AIR",
    "GAS_NITROGEN",
    "CRITICAL_PRESSURE_RATIO",
    "CHOKE_FLOW_CONST",
    "DEFAULT_GAS_TEMP_C",
    "DEFAULT_CD_AIR",
    "gas_constant",
    "celsius_to_kelvin",
    "gauge_bar_to_abs_pa",
    "abs_pa_to_gauge_bar",
    "choke_lower_pressure_abs_pa",
    "is_choked",
    "orifice_mass_flow",
    "nominal_orifice_area_from_k",
    "nominal_orifice_diameter_mm_from_k",
    "orifice_area_circle",
    "US_TO_METRIC_K_FACTOR",
    "us_k_to_metric_k",
]

# --- 상수 (명세서 상수요약) ---
P_ATM_PA: float = 101_325.0          # 표준 대기압 [Pa] = 1.013 bar
KAPPA: float = 1.40                  # 비열비 (공기·질소 공통)
RHO_WATER: float = 1000.0            # 물 밀도 [kg/m³]

# 가스별 기체상수 [J/(kg·K)]
GAS_AIR: str = "air"
GAS_NITROGEN: str = "nitrogen"
GAS_CONSTANTS: dict[str, float] = {
    GAS_AIR: 287.0,
    GAS_NITROGEN: 296.8,
}

# 초크(음속) 판정·계수 — k=1.40에서 임계압력비 0.5283, 상수부 0.5787
CRITICAL_PRESSURE_RATIO: float = (2.0 / (KAPPA + 1.0)) ** (KAPPA / (KAPPA - 1.0))
CHOKE_FLOW_CONST: float = (2.0 / (KAPPA + 1.0)) ** ((KAPPA + 1.0) / (2.0 * (KAPPA - 1.0)))

DEFAULT_GAS_TEMP_C: float = 5.0      # 가스 온도 기본값 [°C] (보수적; 냉동창고 대응)
DEFAULT_CD_AIR: float = 0.70         # 공기 배출계수 기본값 (FM A4 실측 앵커)

# 미국식 K(gpm/psi^0.5) → SI K(L/min/bar^0.5) 환산계수.
#   = (3.785412 L/gal) / sqrt(0.0689476 bar/psi) = 14.4164
# 방출시간(FDT) 계산은 라이브러리 관용 반올림값(예: 160) 대신 이 정밀 환산값(161.5)을
# 써서 SprinkCALC(동일 환산)와 같은 K로 비교한다. 설계 수리계산은 관용값을 유지한다.
US_TO_METRIC_K_FACTOR: float = 14.4164


def gas_constant(gas: str) -> float:
    """가스 종류 → 기체상수 [J/(kg·K)]. 미지원 가스는 ValueError."""
    try:
        return GAS_CONSTANTS[gas]
    except KeyError:
        raise ValueError(
            f"지원하지 않는 가스: {gas!r} — {tuple(GAS_CONSTANTS)} 중 하나여야 함"
        ) from None


def celsius_to_kelvin(temp_c: float) -> float:
    return temp_c + 273.15


def gauge_bar_to_abs_pa(p_gauge_bar: float) -> float:
    """계기압[bar] → 절대압[Pa]."""
    return p_gauge_bar * 1.0e5 + P_ATM_PA


def abs_pa_to_gauge_bar(p_abs_pa: float) -> float:
    """절대압[Pa] → 계기압[bar]."""
    return (p_abs_pa - P_ATM_PA) / 1.0e5


def choke_lower_pressure_abs_pa() -> float:
    """초크가 유지되는 상류 절대압 하한 [Pa] (= P_atm / 임계압력비, ≈ 1.918 bar abs)."""
    return P_ATM_PA / CRITICAL_PRESSURE_RATIO


def is_choked(p_upstream_abs_pa: float) -> bool:
    """상류 절대압에서 대기 방출이 초크(음속)인가."""
    if p_upstream_abs_pa <= P_ATM_PA:
        return False
    return (P_ATM_PA / p_upstream_abs_pa) <= CRITICAL_PRESSURE_RATIO


def orifice_mass_flow(
    p_upstream_abs_pa: float,
    area_m2: float,
    cd: float,
    *,
    gas: str = GAS_AIR,
    temp_k: float,
) -> float:
    """오리피스를 통한 공기 질량유량 [kg/s]. 초크/아음속 자동 분기 (명세서 §4.3).

    하류는 대기압으로 본다(헤드는 대기로 배출). 상류 절대압이 대기압 이하이거나
    면적이 0이면 배출 구동압이 없어 0을 반환한다.

    p_upstream_abs_pa : 경계면(상류) 절대압 [Pa]
    area_m2           : 총 개구 면적 [m²]
    cd                : 공기 배출계수
    gas               : 'air' | 'nitrogen'
    temp_k            : 상류(경계면) 온도 [K]
    """
    if p_upstream_abs_pa <= P_ATM_PA or area_m2 <= 0.0:
        return 0.0
    rs = gas_constant(gas)
    r = P_ATM_PA / p_upstream_abs_pa
    base = cd * area_m2 * p_upstream_abs_pa
    if r <= CRITICAL_PRESSURE_RATIO:  # 초크(음속)
        return base * math.sqrt(KAPPA / (rs * temp_k)) * CHOKE_FLOW_CONST
    # 아음속
    return base * math.sqrt(
        (2.0 * KAPPA / ((KAPPA - 1.0) * rs * temp_k))
        * (r ** (2.0 / KAPPA) - r ** ((KAPPA + 1.0) / KAPPA))
    )


def nominal_orifice_area_from_k(k_factor_si: float) -> float:
    """헤드 K값[lpm/bar^0.5] → Cd=1 등가 명목 오리피스 면적 [m²].

    K-factor에는 물의 배출계수가 이미 내장돼 있다(Q=K·√P 가 실측 방수식).
    이를 'Cd=1 비압축 오리피스'로 환산한 명목 면적이며, 공기 배기 면적의 기준이다 —
    실제 공기 배출계수 Cd_air는 orifice_mass_flow에서 별도로 곱한다(물 Cd 이중계산 방지).
    검증 종결 시 확정된 규약(메모리 fdt-input-contract).

    유도: Q[m³/s] = K/60000·√(P_bar);  Cd=1 오리피스 Q = A·√(2·P/ρ) = A·√(2e5/ρ)·√(P_bar)
          ⇒ A = K / (60000·√(2e5/ρ_water)).
    """
    if k_factor_si <= 0.0:
        return 0.0
    return k_factor_si / (60000.0 * math.sqrt(2.0e5 / RHO_WATER))


def nominal_orifice_diameter_mm_from_k(k_factor_si: float) -> float:
    """헤드 K값 → Cd=1 등가 명목 오리피스 구경 [mm] (표시·검증 보조)."""
    area = nominal_orifice_area_from_k(k_factor_si)
    if area <= 0.0:
        return 0.0
    return math.sqrt(4.0 * area / math.pi) * 1000.0


def orifice_area_circle(diameter_mm: float, count: int = 1) -> float:
    """원형 오리피스 면적 합 [m²] — 구경[mm], 개수."""
    d_m = diameter_mm / 1000.0
    return count * math.pi * (d_m / 2.0) ** 2


def us_k_to_metric_k(k_us: float) -> float:
    """미국식 K값(gpm/psi^0.5) → SI K값(L/min/bar^0.5).

    방출시간 계산이 SprinkCALC와 동일한 K(미국 K의 정밀 환산값)를 쓰도록 한다.
    라이브러리에 미국 K가 있는 표준 헤드에만 적용하고, 커스텀 헤드는 입력 K를 그대로 쓴다.
    예: 11.2 → 161.5, 8.0 → 115.3, 5.6 → 80.7.
    """
    if k_us <= 0.0:
        return 0.0
    return k_us * US_TO_METRIC_K_FACTOR
