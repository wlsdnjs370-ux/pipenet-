"""Scenario water demand, not a complete statutory tank-size certification."""
from __future__ import annotations

from dataclasses import asdict, dataclass
from math import isfinite


@dataclass(frozen=True)
class WaterDemand:
    """Effective scenario volume in cubic metres; no refill credit is assumed."""

    duration_minutes: float | None
    flow_lpm: float
    trial_volume_m3: float | None
    required_volume_m3: float | None
    usable: bool
    basis: str = "selected_active_scenario_only"
    regulatory_compliance: bool = False


def water_demand(flow_lpm: float, duration_minutes: float | None, *, feasible: bool) -> dict:
    """Expose a requirement only after hydraulic convergence AND constraints pass.

    The caller supplies the jurisdiction/scenario duration. Roof storage,
    shared systems, other design areas, dead storage and refill are not included.
    """
    if not isfinite(flow_lpm) or flow_lpm < 0:
        raise ValueError("수원 계산 유량이 유효하지 않습니다.")
    if duration_minutes is not None and (not isfinite(duration_minutes) or duration_minutes <= 0):
        raise ValueError("방수 지속시간은 양수여야 합니다.")
    volume = None if duration_minutes is None else flow_lpm * duration_minutes / 1000
    usable = feasible and volume is not None
    return asdict(WaterDemand(duration_minutes, flow_lpm, volume, volume if usable else None, usable))
