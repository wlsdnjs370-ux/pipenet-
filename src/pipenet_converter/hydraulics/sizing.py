"""Bounded, non-destructive standard-size proposals and pump duty search.

Only enlarges unlocked existing sizes. This is a feasible-proposal heuristic,
not a minimum-cost or globally optimal design. Pump duty is not a pump curve.
"""
from __future__ import annotations

from dataclasses import asdict
from typing import Callable

from .models import BAR_PER_M, Network, Settings
from .solver import Solution, solve
from .water_supply import water_demand


def violations(net: Network, result: Solution, settings: Settings) -> list[dict]:
    """Report hydraulic convergence independently of design acceptance."""
    out = []
    if not result.converged:
        return [{"kind": "nonconverged", "message": "유량·에너지 잔차가 수렴하지 않았습니다."}]
    for h in net.nozzles:
        q, p = result.nozzle_flow_lpm[h.label], result.pressure_bar[h.node]
        if q + .001 < h.min_flow_lpm or p + .00001 < h.min_pressure_bar:
            out.append(dict(kind="head_min", label=h.label, node=h.node, flow_lpm=q, pressure_bar=p,
                            min_flow_lpm=h.min_flow_lpm, min_pressure_bar=h.min_pressure_bar))
        if p > h.max_pressure_bar + .00001:
            out.append(dict(kind="head_max", label=h.label, pressure_bar=p, limit=h.max_pressure_bar))
    for label, p in result.pressure_bar.items():
        if p < -.00001:
            out.append(dict(kind="negative_pressure", label=label, pressure_bar=p))
    for pipe in net.pipes:
        limit = settings.other_velocity_mps if pipe.role == "other" else settings.branch_velocity_mps
        v = result.velocity_mps[pipe.label]
        if v > limit + 1e-5:
            out.append(dict(kind="velocity", label=pipe.label, velocity_mps=v, limit=limit))
    return out


def _required_supply(net, choices, settings, checkpoint):
    """Minimum source pressure for active heads, bounded by user ceiling."""
    def enough(r):
        return r.converged and not any(x["kind"] in {"head_min", "negative_pressure"}
                                      for x in violations(net, r, settings))
    high = settings.max_supply_pressure_bar
    best = solve(net, choices, high, checkpoint=checkpoint)
    if not enough(best):
        return best
    low = 0.0
    for _ in range(25):
        checkpoint()
        mid = (low + high) / 2
        result = solve(net, choices, mid, checkpoint=checkpoint)
        if enough(result):
            high, best = mid, result
        else:
            low = mid
        if high-low < 1e-5:
            break
    return best


def propose(net: Network, settings: Settings, *,
            checkpoint: Callable[[], None] = lambda: None,
            progress: Callable[[str], None] = lambda text: None) -> dict:
    """Return auditable values without changing any input object."""
    net.validate(); settings.validate()
    choices = tuple(p.initial for p in net.pipes)
    trace = []
    status = "iteration_limit"
    before = None
    for iteration in range(1, settings.max_iterations + 1):
        checkpoint()
        progress(f"역산 {iteration}/{settings.max_iterations} — 전체 연결 동시 계산")
        if settings.mode == "fixed_supply":
            result = solve(net, choices, settings.supply_pressure_bar, checkpoint=checkpoint)
        else:
            result = _required_supply(net, choices, settings, checkpoint)
        bad = violations(net, result, settings)
        if before is None:
            before = asdict(result)
        trace.append(dict(iteration=iteration, supply_bar=result.source_pressure_bar,
                          flow_lpm=result.source_flow_lpm, violations=len(bad),
                          converged=result.converged))
        if not result.converged:
            status = "nonconverged"; break
        if not bad:
            status = "feasible_review"; break
        if settings.mode == "pump_only":
            status = "constraints_unmet"; break
        # A high pressure at a head cannot be fixed safely by blindly enlarging
        # upstream pipes. A PRV/zoning design is a separate task.
        if any(x["kind"] == "head_max" for x in bad):
            status = "constraints_unmet"; break
        velocity = {x["label"] for x in bad if x["kind"] == "velocity"}
        head_short = any(x["kind"] in {"head_min", "negative_pressure"} for x in bad)
        nxt = list(choices)
        for i, pipe in enumerate(net.pipes):
            if not pipe.locked and choices[i]+1 < len(pipe.sizes) and (head_short or pipe.label in velocity):
                nxt[i] += 1
        if tuple(nxt) == choices:
            status = "bounds_exhausted"; break
        # Do not return uncalculated sizes when the iteration budget expires.
        if iteration < settings.max_iterations:
            choices = tuple(nxt)
    source_z = next(n.elevation_m for n in net.nodes if n.label == net.source)
    discharge_h = source_z + result.source_pressure_bar / BAR_PER_M + settings.discharge_velocity_head_m
    duty_h = max(0., discharge_h-settings.suction_total_head_m)
    pipe_rows = []
    for p, selected in zip(net.pipes, choices):
        old, new = p.sizes[p.initial], p.sizes[selected]
        pipe_rows.append(dict(label=p.label, before_dn=old.nominal_mm, after_dn=new.nominal_mm,
            inner_mm=new.inner_mm, equivalent_m=new.equivalent_m, locked=p.locked, role=p.role,
            flow_lpm=result.flow_lpm[p.label], velocity_mps=result.velocity_mps[p.label]))
    return dict(schema="module-f-hydraulic-sizing-v1", status=status,
        feasible=status == "feasible_review", review_only=True, settings=asdict(settings),
        source=net.source, iterations=iteration, trace=trace, violations=bad,
        before=before, solution=asdict(result), pipes=pipe_rows,
        water_supply=water_demand(result.source_flow_lpm, settings.duration_minutes,
                                 feasible=status == "feasible_review"),
        pump=dict(usable=status == "feasible_review",
            flow_lpm=result.source_flow_lpm, differential_head_m=duty_h,
            discharge_pressure_bar=result.source_pressure_bar,
            suction_total_head_m=settings.suction_total_head_m,
            discharge_velocity_head_m=settings.discharge_velocity_head_m,
            gravity_surplus_m=max(0.,settings.suction_total_head_m-discharge_h),
            hydraulic_power_kw=998.2*9.80665*result.source_flow_lpm/60000*duty_h/1000,
            kind="required_duty_point_not_manufacturer_curve"),
        nozzles=[dict(label=h.label,node=h.node,k=h.k_lpm_sqrt_bar,
            min_flow_lpm=h.min_flow_lpm,min_pressure_bar=h.min_pressure_bar,
            max_pressure_bar=h.max_pressure_bar,flow_lpm=result.nozzle_flow_lpm[h.label],
            pressure_bar=result.pressure_bar[h.node]) for h in net.nozzles])
