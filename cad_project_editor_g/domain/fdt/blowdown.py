"""domain/fdt/blowdown.py

K-Fire FDT 모듈 — Phase A 블로다운(트립시간) 코어.

밸브 하류 압축공기가 개방 헤드로 빠지며 공기압이 트립 임계까지 떨어지는 시간(t_trip)을
계산한다. 순수 0차원 공기 모델 — EPANET·전선추적 불필요. 계획서 §5.1, 명세서 §4.0.

열역학 2모드:
- adiabatic(단열): 제품 기본. 빠른 블로다운이라 벽면 열교환이 적어 단열에 가깝고,
  등온보다 트립이 짧게(실측에 근접) 예측된다.
- isothermal(등온): 닫힌해가 존재해 검증 정답지로 쓰며, 보수측(더 길게) 비교용이다.

트립 임계압(밸브가 열리는 공기압)은 설비 종류로 결정되지만 본 코어는 임계압 자체만 받는다:
- 건식밸브        : 트립임계 = 1차측 수압 / 차압비   → dry_pipe_trip_threshold_gauge_bar
- 더블인터록 프리액션: 트립압 직접입력
환산 책임은 상위 계층(밸브 노드 속성을 읽는 곳)에 둔다 — 코어는 물리만.
"""
from __future__ import annotations

from dataclasses import dataclass

from domain.fdt import air_model as air

__all__ = [
    "BlowdownInput",
    "BlowdownResult",
    "trip_time",
    "dry_pipe_trip_threshold_gauge_bar",
    "primary_pressure_ceiling_gauge_bar",
]

# 1 bar ≈ 10.197 m 수두 (엔진 BAR_TO_M과 일치)
_BAR_TO_M_WATER: float = 10.197


@dataclass(frozen=True)
class BlowdownInput:
    """트립(블로다운) 계산 입력. 압력은 계기압[bar], 체적[m³], 온도[°C]."""

    volume_m3: float                       # 건식밸브 하류 공기 체적
    p_initial_gauge_bar: float             # 초기 공기압(감시압)
    p_trip_gauge_bar: float                # 트립 임계 공기압
    orifice_area_m2: float                 # 개방 헤드 총 배기 면적
    cd_air: float = air.DEFAULT_CD_AIR     # 공기 배출계수
    gas: str = air.GAS_AIR                 # 'air' | 'nitrogen'
    temp_c: float = air.DEFAULT_GAS_TEMP_C  # 가스 온도
    thermo: str = "adiabatic"              # 'adiabatic' | 'isothermal'


@dataclass(frozen=True)
class BlowdownResult:
    t_trip_s: float                        # 트립까지 시간 [s]
    reached_trip: bool                     # 트립 임계 도달 여부(False = stall/시간초과)
    choked_throughout: bool                # 전 구간 초크 유지 여부
    p_final_gauge_bar: float               # 종료 시 공기압 [bar, 게이지]
    mass_balance_residual: float           # |Σṁ·dt + m_final − m_init| / m_init
    trajectory: list[tuple[float, float, float]] | None = None  # (t, P_abs[Pa], m[kg])


def dry_pipe_trip_threshold_gauge_bar(
    primary_water_gauge_bar: float, differential_ratio: float
) -> float:
    """건식밸브 트립 임계 공기압[bar, 게이지] = 1차측 수압 / 차압비."""
    if differential_ratio <= 0.0:
        raise ValueError("차압비는 양수여야 함")
    return primary_water_gauge_bar / differential_ratio


def primary_pressure_ceiling_gauge_bar(
    pump_shutoff_gauge_bar: float, elevation_head_m: float
) -> float:
    """1차측 수압 입력 상한[bar, 게이지] = 펌프 체절압 − 표고수두.

    1차측(밸브 공급측) 실수압은 펌프 체절압에서 밸브까지의 표고수두를 뺀 값을 넘을 수 없다.
    elevation_head_m = (건식밸브 표고 − 펌프/수원 표고) [m].
    """
    return pump_shutoff_gauge_bar - elevation_head_m / _BAR_TO_M_WATER


def _validate(inp: BlowdownInput) -> None:
    if inp.volume_m3 <= 0.0:
        raise ValueError("체적(volume_m3)은 양수여야 함")
    if inp.orifice_area_m2 <= 0.0:
        raise ValueError("배기 면적(orifice_area_m2)은 양수여야 함")
    if inp.cd_air <= 0.0:
        raise ValueError("배출계수(cd_air)는 양수여야 함")
    if inp.thermo not in ("adiabatic", "isothermal"):
        raise ValueError("thermo는 'adiabatic' 또는 'isothermal'")
    air.gas_constant(inp.gas)  # 미지원 가스면 여기서 ValueError


def trip_time(
    inp: BlowdownInput,
    *,
    dt: float = 1.0e-3,
    max_time_s: float = 600.0,
    record_trajectory: bool = False,
) -> BlowdownResult:
    """블로다운 적분 → 트립시간. RK4(공기질량 1상태), 트립 도달은 가스법칙으로 정확 보간.

    dt               : 시간스텝 [s] (적분 정확도, 결과는 충분히 작으면 수렴)
    max_time_s       : 시뮬레이션 상한(미도달 stall 안전장치)
    record_trajectory: True면 (t, P_abs, m) 이력 반환(검증·가시화 보조)
    """
    _validate(inp)

    rs = air.gas_constant(inp.gas)
    t0_k = air.celsius_to_kelvin(inp.temp_c)
    v = inp.volume_m3
    p0_abs = air.gauge_bar_to_abs_pa(inp.p_initial_gauge_bar)
    p_trip_abs = air.gauge_bar_to_abs_pa(inp.p_trip_gauge_bar)
    iso = inp.thermo == "isothermal"

    m0 = p0_abs * v / (rs * t0_k)
    traj: list[tuple[float, float, float]] | None = (
        [(0.0, p0_abs, m0)] if record_trajectory else None
    )

    # F3: 초기압이 이미 트립압 이하 → 즉시 트립(0초).
    if p0_abs <= p_trip_abs:
        return BlowdownResult(
            0.0, True, air.is_choked(p0_abs), inp.p_initial_gauge_bar, 0.0, traj
        )

    # F4: 트립압이 대기압 이하 → 대기 방출로는 도달 불가(구동압 소멸). stall.
    if p_trip_abs <= air.P_ATM_PA:
        return BlowdownResult(
            max_time_s, False, air.is_choked(p0_abs),
            inp.p_initial_gauge_bar, 0.0, traj,
        )

    # 트립 시점 공기질량(가스법칙): 등온 m/m0 = P/P0, 단열 m/m0 = (P/P0)^(1/k)
    ratio = p_trip_abs / p0_abs
    m_trip = m0 * (ratio if iso else ratio ** (1.0 / air.KAPPA))

    def pressure_abs(m: float) -> float:
        rr = m / m0
        return p0_abs * (rr if iso else rr ** air.KAPPA)

    def temp_k(m: float) -> float:
        if iso:
            return t0_k
        return t0_k * (m / m0) ** (air.KAPPA - 1.0)

    def mdot(m: float) -> float:
        return air.orifice_mass_flow(
            pressure_abs(m), inp.orifice_area_m2, inp.cd_air, gas=inp.gas, temp_k=temp_k(m)
        )

    t = 0.0
    m = m0
    choked_throughout = True
    vented = 0.0  # ∫ṁ·dt 사다리꼴 누적 (I1 질량보존 독립 검산용)

    while t < max_time_s:
        if not air.is_choked(pressure_abs(m)):
            choked_throughout = False

        rate = mdot(m)
        if rate <= 0.0:
            # 배출 정지인데 트립 미도달 → stall.
            return BlowdownResult(
                t, False, choked_throughout, air.abs_pa_to_gauge_bar(pressure_abs(m)),
                _residual(m0, m, vented), traj,
            )

        # RK4 (autonomous: dm/dt = −mdot(m))
        k1 = -rate
        k2 = -mdot(m + 0.5 * dt * k1)
        k3 = -mdot(m + 0.5 * dt * k2)
        k4 = -mdot(m + dt * k3)
        m_next = m + dt / 6.0 * (k1 + 2.0 * k2 + 2.0 * k3 + k4)

        if m_next <= m_trip:
            # 트립 도달 — 질량으로 선형보간해 정확한 교차시간.
            frac = (m - m_trip) / (m - m_next) if m != m_next else 1.0
            t_trip = t + frac * dt
            vented += 0.5 * (mdot(m) + mdot(m_trip)) * (frac * dt)
            if traj is not None:
                traj.append((t_trip, p_trip_abs, m_trip))
            return BlowdownResult(
                t_trip, True, choked_throughout, inp.p_trip_gauge_bar,
                _residual(m0, m_trip, vented), traj,
            )

        vented += 0.5 * (rate + mdot(m_next)) * dt
        m = m_next
        t += dt
        if traj is not None:
            traj.append((t, pressure_abs(m), m))

    # 시간 초과(미도달).
    return BlowdownResult(
        t, False, choked_throughout, air.abs_pa_to_gauge_bar(pressure_abs(m)),
        _residual(m0, m, vented), traj,
    )


def _residual(m0: float, m_final: float, vented: float) -> float:
    """질량보존 잔차 = |(m0 − m_final) − ∫ṁ·dt| / m0."""
    if m0 <= 0.0:
        return 0.0
    return abs((m0 - m_final) - vented) / m0
