"""domain/fdt/accelerator.py

K-Fire FDT 모듈 — 가속기(Quick-Opening Device) 트립 가속.

건식/더블인터록 밸브에 가속기가 설치되면 밸브가 블로다운 완료 전에 트립한다.
새 물리를 만들지 않고 **트립 코어(blowdown)를 '정지조건만 바꿔' 재사용**한다.
근거: 리서치(게이트 docs/FDT/C_accelerator_step.md) — Tyco/JCI 백서가 SprinkFDT의
가속기 모델링을 직접 명시(전자식=고정 트립시간, 기계식=고정 공기압 강하).

두 유형:
- 전자식(electronic): 압력강하율을 감지해 고정 트립시간 t_acc 후 솔레노이드로 트립.
  트립시간 = min(t_acc, 무가속기 트립시간) — 초소형 설비서 역효과(트립 지연) 금지.
  앵커: Reliable Model C 3초·Tyco VIZOR 2초 → 보수 기본 3초.
- 기계식(mechanical): 차압챔버가 공기압 강하 ΔP를 감지하면 시스템 공기를 중간챔버로
  보내 차압을 소멸시켜 트립. 트립 = 감시압에서 max(P0−ΔP, 트립임계)까지 블로다운.
  ΔP가 차압여유(P0−트립임계)보다 크면 차압 클래퍼가 먼저 트립(가속기 무이득).
  앵커: Tyco ACC-1·Viking E-1 3~5 psi + 응답지연 → 기본 0.30 bar.

가속기는 트립 시점 공기압이 높게 남으므로, 트랜짓은 그 **잔류 공기압**에서 출발해야 한다
(상위 계층이 residual_air_gauge_bar를 도달엔진 시작압으로 사용). 백서 입증: 같은 망에서
가속기를 쓰면 트립은 줄지만 트랜짓은 (잔류공기 배기 탓에) 같거나 약간 길다.

순수 Python — PySide6·WNTR·infra 의존 없음(계획서 §6 4계층).
"""
from __future__ import annotations

from dataclasses import dataclass, replace

from domain.fdt import air_model as air
from domain.fdt.blowdown import BlowdownInput, trip_time

__all__ = [
    "ACCEL_ELECTRONIC",
    "ACCEL_MECHANICAL",
    "DEFAULT_ELECTRONIC_TRIP_S",
    "DEFAULT_MECHANICAL_DROP_BAR",
    "AcceleratorSpec",
    "AcceleratorTripResult",
    "accelerated_trip",
]

ACCEL_ELECTRONIC: str = "electronic"
ACCEL_MECHANICAL: str = "mechanical"

# 디폴트 셋포인트 (사용자 확정 2026-06-20)
DEFAULT_ELECTRONIC_TRIP_S: float = 3.0      # VIZOR 2초·Reliable Model C 3초 중 보수측
DEFAULT_MECHANICAL_DROP_BAR: float = 0.30   # ACC-1/Valve A 3~5psi+지연의 라운드값


@dataclass(frozen=True)
class AcceleratorSpec:
    """가속기 설정. 전자식은 trip_time_s, 기계식은 pressure_drop_bar만 의미 있다."""

    kind: str                                          # 'electronic' | 'mechanical'
    trip_time_s: float = DEFAULT_ELECTRONIC_TRIP_S     # 전자식 고정 트립시간 [s]
    pressure_drop_bar: float = DEFAULT_MECHANICAL_DROP_BAR  # 기계식 트립 공기압 강하 [bar]


@dataclass(frozen=True)
class AcceleratorTripResult:
    t_trip_s: float                    # 가속기 트립시간 [s]
    residual_air_gauge_bar: float      # 트랜짓 시작 공기압(잔류) [bar, 게이지]
    reached_trip: bool                 # 트립 도달 여부
    baseline_t_trip_s: float           # 무가속기 트립시간(대비) [s]
    baseline_residual_gauge_bar: float  # 무가속기 트랜짓 시작압(= 트립임계) [bar]
    effective: bool                    # 가속기가 실제로 트립을 줄였는가
    saved_s: float                     # 트립 절감(무가속기 − 가속기) [s]


def _validate_spec(spec: AcceleratorSpec) -> None:
    if spec.kind not in (ACCEL_ELECTRONIC, ACCEL_MECHANICAL):
        raise ValueError(
            f"지원하지 않는 가속기 유형: {spec.kind!r} — "
            f"'{ACCEL_ELECTRONIC}' 또는 '{ACCEL_MECHANICAL}'"
        )
    if spec.kind == ACCEL_ELECTRONIC and spec.trip_time_s <= 0.0:
        raise ValueError("전자식 가속기 트립시간(trip_time_s)은 양수여야 함")
    if spec.kind == ACCEL_MECHANICAL and spec.pressure_drop_bar <= 0.0:
        raise ValueError("기계식 가속기 공기압 강하(pressure_drop_bar)는 양수여야 함")


def _pressure_gauge_at_time(
    trajectory: list[tuple[float, float, float]] | None, t_query: float
) -> float:
    """블로다운 이력 (t, P_abs[Pa], m)에서 t_query 시점 공기압[bar, 게이지]을 선형보간."""
    if not trajectory:
        return 0.0
    for (t0, p0, _), (t1, p1, _) in zip(trajectory, trajectory[1:]):
        if t0 <= t_query <= t1:
            frac = (t_query - t0) / (t1 - t0) if t1 > t0 else 0.0
            return air.abs_pa_to_gauge_bar(p0 + frac * (p1 - p0))
    return air.abs_pa_to_gauge_bar(trajectory[-1][1])


def accelerated_trip(
    inp: BlowdownInput,
    spec: AcceleratorSpec,
    *,
    dt: float = 1.0e-3,
    max_time_s: float = 600.0,
) -> AcceleratorTripResult:
    """가속기 설치 시 트립시간 + 트랜짓 시작 공기압. 무가속기 기준도 함께 반환(대비표용).

    inp  : 무가속기와 동일한 블로다운 입력(체적·초기압·트립임계·배기면적·가스·온도·열역학).
    spec : 가속기 설정.

    트립 코어를 정지조건만 바꿔 재사용한다 — 코어(blowdown) 무수정.
    """
    _validate_spec(spec)

    # 무가속기 기준(대비·캡). 전자식은 잔류압 보간용 이력이 필요하다.
    need_traj = spec.kind == ACCEL_ELECTRONIC
    base = trip_time(inp, dt=dt, max_time_s=max_time_s, record_trajectory=need_traj)

    if spec.kind == ACCEL_MECHANICAL:
        # 차압챔버가 ΔP를 감지하면 트립 — 단 차압여유보다 ΔP가 크면 클래퍼가 먼저(=무이득).
        stop_gauge = max(
            inp.p_initial_gauge_bar - spec.pressure_drop_bar, inp.p_trip_gauge_bar
        )
        res = trip_time(replace(inp, p_trip_gauge_bar=stop_gauge), dt=dt, max_time_s=max_time_s)
        t_acc = res.t_trip_s
        residual = stop_gauge
        reached = res.reached_trip

    else:  # ACCEL_ELECTRONIC
        if spec.trip_time_s >= base.t_trip_s:
            # 고정 트립시간이 무가속기보다 느림 → 가속기 무이득(트립 지연 금지).
            t_acc = base.t_trip_s
            residual = inp.p_trip_gauge_bar
            reached = base.reached_trip
        else:
            # 솔레노이드가 t_acc에 강제 트립 — 잔류압 = 그 시각 블로다운 압력.
            t_acc = spec.trip_time_s
            residual = _pressure_gauge_at_time(base.trajectory, t_acc)
            reached = True

    return AcceleratorTripResult(
        t_trip_s=t_acc,
        residual_air_gauge_bar=residual,
        reached_trip=reached,
        baseline_t_trip_s=base.t_trip_s,
        baseline_residual_gauge_bar=inp.p_trip_gauge_bar,
        effective=t_acc < base.t_trip_s - 1.0e-9,
        saved_s=base.t_trip_s - t_acc,
    )
