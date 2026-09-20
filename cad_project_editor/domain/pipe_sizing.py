"""Pure domain primitives for automatic and manual pipe sizing.

This module intentionally has no Qt, EPANET, or editor dependency.  It owns the
network classification, velocity/capacity arithmetic, range validation, and the
preview data contract used by the application service and UI.
"""
from __future__ import annotations

from collections import deque
from dataclasses import dataclass, field
from enum import Enum
import hashlib
import json
import math
from types import SimpleNamespace
from typing import Any, Iterable, Mapping, Sequence

from domain.models import NodeTypeId
from domain.node_predicates import is_pump_graph_node, is_tank_graph_node
from domain.topology import analyze_pump_topology


DEFAULT_BRANCH_VELOCITY_MPS = 6.0
DEFAULT_MAIN_VELOCITY_MPS = 10.0
DEFAULT_SUPPLY_VELOCITY_MPS = 3.0
UNUSUALLY_HIGH_VELOCITY_MPS = 20.0
FLOW_TOLERANCE_LPM = 1.0e-6
# Grid no-flow threshold: above solver residual noise, well below a single head.
FLOW_EPS_LPM = 0.1
# 헤드·노즐 짧은 드롭(니플). 그리드 판별(touches_short_drop)에 쓴다.
_HEAD_DROP_MAX_M = 1.0
_DISCHARGE_CONTROL_VALVE_TOKENS = frozenset(
    {
        "alarm_valve",
        "alarm valve",
        "알람밸브",
        "dry_pipe_valve",
        "dry_pipe",
        "건식밸브",
        "preaction_valve",
        "preaction",
        "준비작동식",
        "준비작동식밸브",
        "준비작동",
    }
)
# 헤드·물분무 노즐을 하류에 두는 배관의 법규 최소 구경.
MIN_DISCHARGE_PATH_NOMINAL_MM = 25
# 호스노즐(옥내소화전)을 하류에 두는 배관의 법규 최소 구경.
MIN_HOSE_PATH_NOMINAL_MM = 40
# 가지(BRANCH) 배관의 최소 구경 — 장치 없는 막다른 관말(스터브)도 이
# 밑으로 줄이지 않는다(무유량이라도 라이브러리 최소 규격까지 축소 금지).
MIN_BRANCH_NOMINAL_MM = 25
# 교차·수평주행(MAIN) 배관의 최소 구경 — 가지와 같다.
MIN_MAIN_NOMINAL_MM = 25
# 펌프 토출측(SUPPLY) 배관의 최소 구경.
MIN_SUPPLY_NOMINAL_MM = 65


class SizingBasis(str, Enum):
    HEAD_COUNT = "head_count"
    FLOW = "flow"


class PipeClass(str, Enum):
    BRANCH = "branch"
    MAIN = "main"
    SUPPLY = "supply"


class ChangeStatus(str, Enum):
    DECREASE = "decrease"
    INCREASE = "increase"
    UNCHANGED = "unchanged"
    INVALID = "invalid"
    VALVE_ADJUSTED = "valve_adjusted"
    MANUAL_OVERRIDE = "manual_override"


@dataclass(frozen=True)
class VelocityLimits:
    branch_mps: float = DEFAULT_BRANCH_VELOCITY_MPS
    main_mps: float = DEFAULT_MAIN_VELOCITY_MPS
    supply_mps: float = DEFAULT_SUPPLY_VELOCITY_MPS

    def for_class(self, pipe_class: PipeClass) -> float:
        if pipe_class is PipeClass.SUPPLY:
            return float(self.supply_mps)
        # MAIN is no longer assigned; leftover MAIN reads as branch velocity.
        return float(self.branch_mps)

    def validate(self) -> tuple[list[str], list[str]]:
        errors: list[str] = []
        warnings: list[str] = []
        for key, value in (
            ("branch", self.branch_mps),
            ("main", self.main_mps),
            ("supply", self.supply_mps),
        ):
            try:
                numeric = float(value)
            except (TypeError, ValueError, OverflowError):
                errors.append(f"velocity_not_numeric:{key}")
                continue
            if not math.isfinite(numeric) or numeric <= 0.0:
                errors.append(f"velocity_not_positive:{key}")
            elif numeric > UNUSUALLY_HIGH_VELOCITY_MPS:
                warnings.append(f"velocity_unusually_high:{key}:{numeric:g}")
        return errors, warnings

    def to_dict(self) -> dict[str, Any]:
        def serialized(value: Any) -> Any:
            try:
                return float(value)
            except (TypeError, ValueError, OverflowError):
                # Preserve the bad scalar so the subsequent domain validation
                # can report its field-specific code; do not raise while merely
                # normalizing persisted settings.
                return value

        return {
            "branch_mps": serialized(self.branch_mps),
            "main_mps": serialized(self.main_mps),
            "supply_mps": serialized(self.supply_mps),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any] | None) -> "VelocityLimits":
        # Keep malformed scalar values intact so ``validate`` can return a
        # structured error instead of leaking ValueError from a persistence or
        # UI boundary. A non-mapping container is treated like missing data.
        raw = data if isinstance(data, Mapping) else {}
        return cls(
            branch_mps=raw.get("branch_mps", DEFAULT_BRANCH_VELOCITY_MPS),
            main_mps=raw.get("main_mps", DEFAULT_MAIN_VELOCITY_MPS),
            supply_mps=raw.get("supply_mps", DEFAULT_SUPPLY_VELOCITY_MPS),
        )


@dataclass(frozen=True, order=True)
class PipeSizeSpec:
    nominal_mm: int
    inner_d_mm: float
    standard: str = ""
    name: str = ""
    c_factor: float = 120.0
    roughness_mm: float = 0.1

    def validate(self) -> list[str]:
        """Return machine-readable diagnostics without raising on bad input."""
        errors: list[str] = []
        try:
            nominal = int(self.nominal_mm)
        except (TypeError, ValueError, OverflowError):
            errors.append("pipe_spec_nominal_not_integer")
        else:
            if nominal <= 0:
                errors.append("pipe_spec_nominal_not_positive")
        for name, value, positive in (
            ("inner_d_mm", self.inner_d_mm, True),
            ("c_factor", self.c_factor, True),
            ("roughness_mm", self.roughness_mm, False),
        ):
            try:
                numeric = float(value)
            except (TypeError, ValueError, OverflowError):
                errors.append(f"pipe_spec_{name}_not_numeric")
                continue
            if not math.isfinite(numeric):
                errors.append(f"pipe_spec_{name}_not_finite")
            elif positive and numeric <= 0.0:
                errors.append(f"pipe_spec_{name}_not_positive")
            elif not positive and numeric < 0.0:
                errors.append(f"pipe_spec_{name}_negative")
        return errors


@dataclass(frozen=True)
class RangeRow:
    start: float
    end: float | None
    nominal_mm: int
    _input_errors: tuple[str, ...] = field(default=(), repr=False, compare=False)

    def contains(self, value: float, *, first: bool = False) -> bool:
        """Return whether value belongs to the range.

        Shared boundaries belong to the smaller/previous range.  Rows are
        therefore evaluated in order and use an inclusive upper boundary.
        """
        v = float(value)
        lower_ok = v >= self.start - FLOW_TOLERANCE_LPM if first else v > self.start + FLOW_TOLERANCE_LPM
        if first and math.isclose(v, self.start, abs_tol=FLOW_TOLERANCE_LPM, rel_tol=0.0):
            lower_ok = True
        upper_ok = self.end is None or v <= self.end + FLOW_TOLERANCE_LPM
        return lower_ok and upper_ok

    def to_dict(self) -> dict[str, Any]:
        return {
            "start": float(self.start),
            "end": None if self.end is None else float(self.end),
            "nominal_mm": int(self.nominal_mm),
        }

    @classmethod
    def from_dict(cls, data: Mapping[str, Any] | Any) -> "RangeRow":
        if not isinstance(data, Mapping):
            return cls(math.nan, math.nan, 0, ("range_row_not_mapping",))

        errors: list[str] = []
        try:
            start = float(data.get("start", 0.0))
        except (TypeError, ValueError, OverflowError):
            start = math.nan
            errors.append("range_start_not_numeric")
        raw_end = data.get("end")
        if raw_end in (None, ""):
            end = None
        else:
            try:
                end = float(raw_end)
            except (TypeError, ValueError, OverflowError):
                end = math.nan
                errors.append("range_end_not_numeric")
        try:
            nominal = int(data.get("nominal_mm", 0))
        except (TypeError, ValueError, OverflowError):
            nominal = 0
            errors.append("range_dn_not_integer")
        return cls(
            start=start,
            end=end,
            nominal_mm=nominal,
            _input_errors=tuple(errors),
        )


@dataclass(frozen=True)
class PipeSizingContext:
    pipe_id: str
    basis: SizingBasis
    pipe_class: PipeClass
    downstream_head_count: int
    estimated_flow_lpm: float
    is_cycle: bool = False
    is_supply_path: bool = False
    # 헤드·물분무 노즐이 배관 끝점에 직접 달려 있는가(=소방 기준의 가지배관).
    has_discharge_endpoint: bool = False
    # 짧은 니플(≤1 m) 드롭이 붙은 정션에 닿는 cycle/트리 배관.
    touches_short_drop: bool = False
    downstream_node_ids: Sequence[str] = ()
    # 하류 호스노즐(옥내소화전) 수 — 설계 수량(비활성 포함).
    downstream_hose_count: int = 0
    # 하류에서 실제 유량을 만드는 장치 수(활성 헤드·노즐/호스).
    downstream_active_head_count: int = 0
    downstream_active_hose_count: int = 0


def is_grid_pipe(context: PipeSizingContext) -> bool:
    """True when a discharge endpoint sits directly on a cycle edge.

    Pure loops (cycle mains with tree branches only) are not grids.
    Short-drop grids are detected via ``is_grid_network`` instead.
    """
    return bool(context.is_cycle and context.has_discharge_endpoint)


def count_grid_pipes(contexts: Mapping[str, PipeSizingContext]) -> int:
    """Count cycle pipes that carry a direct head/nozzle endpoint."""
    return sum(1 for context in contexts.values() if is_grid_pipe(context))


def is_grid_network(contexts: Mapping[str, PipeSizingContext]) -> bool:
    """True when heads/nozzles sit on or hang from cycle edges.

    Direct-on-cycle heads and short-drop (≤1 m) nipples off a loop both count.
    Pure loops with tree-only branches remain non-grid.
    """
    return any(
        bool(context.is_cycle)
        and (
            bool(context.has_discharge_endpoint)
            or bool(getattr(context, "touches_short_drop", False))
        )
        for context in contexts.values()
    )


def count_grid_network_pipes(contexts: Mapping[str, PipeSizingContext]) -> int:
    """Count cycle pipes that make the network a sizing grid."""
    return sum(
        1
        for context in contexts.values()
        if bool(context.is_cycle)
        and (
            bool(context.has_discharge_endpoint)
            or bool(getattr(context, "touches_short_drop", False))
        )
    )


def grid_uniform_velocity_limits(limits: VelocityLimits) -> VelocityLimits:
    """Apply the branch velocity limit to every class (grid has no class split)."""
    v = float(limits.branch_mps)
    return VelocityLimits(branch_mps=v, main_mps=v, supply_mps=v)


def minimum_nominal_for_context(
    context: PipeSizingContext,
    *,
    ignore_pipe_class: bool = False,
) -> int:
    """Return the code-minimum DN for a pipe, 0 when unconstrained.

    Floors combine (largest wins): heads/nozzles downstream → DN25, hose
    streams (옥내소화전) downstream → DN40, BRANCH(가지) → DN25,
    MAIN(교차·수평주행) → DN25, SUPPLY(펌프 토출측) → DN65. The BRANCH floor
    covers device-less dead-end stubs so zero flow never selects below DN25.
    For loops the downstream summary covers the whole network, so cycle pipes
    inherit device floors network-wide. Grid sizing sets ``ignore_pipe_class``
    so class floors are skipped (the grid path reasserts DN25 itself).
    """
    floor = 0
    if int(context.downstream_head_count) > 0:
        floor = max(floor, MIN_DISCHARGE_PATH_NOMINAL_MM)
    if int(getattr(context, "downstream_hose_count", 0) or 0) > 0:
        floor = max(floor, MIN_HOSE_PATH_NOMINAL_MM)
    if not ignore_pipe_class:
        if context.pipe_class is PipeClass.MAIN:
            floor = max(floor, MIN_MAIN_NOMINAL_MM)
        elif context.pipe_class is PipeClass.SUPPLY:
            floor = max(floor, MIN_SUPPLY_NOMINAL_MM)
        else:
            floor = max(floor, MIN_BRANCH_NOMINAL_MM)
    return floor


def sizing_flow_producers_missing(context: PipeSizingContext) -> bool:
    """True when a flow-basis pipe serves devices that are all inactive.

    Such pipes would size from a meaningless 0 lpm, so automatic sizing
    keeps their current diameter instead of shrinking (freeze rule).
    """
    if context.basis is not SizingBasis.FLOW:
        return False
    design_devices = int(context.downstream_head_count) + int(
        getattr(context, "downstream_hose_count", 0) or 0
    )
    if design_devices <= 0:
        return False
    active_devices = int(
        getattr(context, "downstream_active_head_count", 0) or 0
    ) + int(getattr(context, "downstream_active_hose_count", 0) or 0)
    return active_devices <= 0


def build_grid_sizing_flows(
    graph: Any,
    contexts: Mapping[str, PipeSizingContext],
    *,
    default_pressure_bar: float = 1.0,
) -> dict[str, float]:
    """Legacy design-flow map for grids (not used by auto sizing).

    Auto sizing uses Demand reverse-calc pipe flows instead: flowing pipes
    size by solved Q + uniform velocity; near-zero Q pipes propose DN25.
    Kept for diagnostics / warm-start experiments only.

    - Short drops: that head's design discharge.
    - Head-lines (short-drop junctions linked by cycle edges): sum of all heads
      on the line for every cycle segment on that line.
    - Everything else: max(estimated_flow, incident line flows).
    """
    nodes, pipes, _adjacency = _graph_parts(graph)
    flows: dict[str, float] = {
        str(pipe_id): max(float(context.estimated_flow_lpm), 0.0)
        for pipe_id, context in contexts.items()
    }
    if not pipes:
        return flows

    # Drop pipe → junction heads (short nipple ≤ 1 m).
    junction_heads: dict[str, list[tuple[str, float]]] = {}
    short_drop_pipes: set[str] = set()
    for pipe_id, pipe in pipes.items():
        pid = str(pipe_id)
        start = str(_get(pipe, "start", "") or "")
        end = str(_get(pipe, "end", "") or "")
        if start not in nodes or end not in nodes or start == end:
            continue
        start_d = _is_discharge_node(nodes, start)
        end_d = _is_discharge_node(nodes, end)
        if start_d == end_d:
            continue
        if _pipe_length_m(pipe) > _HEAD_DROP_MAX_M:
            continue
        head_id = start if start_d else end
        junction = end if start_d else start
        head_flow = max(design_discharge_lpm(nodes[head_id], default_pressure_bar), 0.0)
        junction_heads.setdefault(junction, []).append((head_id, head_flow))
        # Short drops size to their own head only (not line / network estimate).
        flows[pid] = head_flow
        short_drop_pipes.add(pid)

    short_junctions = set(junction_heads)
    # Cycle edges that link two short-drop junctions form head-lines.
    line_edge_adj: dict[str, list[tuple[str, str]]] = {
        node_id: [] for node_id in short_junctions
    }
    for pipe_id, context in contexts.items():
        if not context.is_cycle:
            continue
        pipe = pipes.get(pipe_id)
        if pipe is None:
            continue
        start = str(_get(pipe, "start", "") or "")
        end = str(_get(pipe, "end", "") or "")
        if start in short_junctions and end in short_junctions:
            line_edge_adj[start].append((end, pipe_id))
            line_edge_adj[end].append((start, pipe_id))

    visited_junctions: set[str] = set()
    line_flow_by_pipe: dict[str, float] = {}
    line_flow_by_junction: dict[str, float] = {}
    for seed in short_junctions:
        if seed in visited_junctions:
            continue
        stack = [seed]
        component: list[str] = []
        seen = {seed}
        while stack:
            current = stack.pop()
            component.append(current)
            for neighbor, _edge in line_edge_adj.get(current, ()):
                if neighbor in seen:
                    continue
                seen.add(neighbor)
                stack.append(neighbor)
        visited_junctions.update(component)
        # Same head may hang from multiple junctions; count once per line.
        seen_heads: set[str] = set()
        line_flow = 0.0
        for junction in component:
            for head_id, head_flow in junction_heads.get(junction, ()):
                if head_id in seen_heads:
                    continue
                seen_heads.add(head_id)
                line_flow += head_flow
        for junction in component:
            line_flow_by_junction[junction] = line_flow
        for junction in component:
            for neighbor, edge_id in line_edge_adj.get(junction, ()):
                if neighbor not in seen:
                    continue
                line_flow_by_pipe[edge_id] = max(
                    line_flow_by_pipe.get(edge_id, 0.0), line_flow
                )

    for pipe_id, line_flow in line_flow_by_pipe.items():
        if pipe_id in short_drop_pipes:
            continue
        flows[pipe_id] = max(flows.get(pipe_id, 0.0), line_flow)

    # Direct-on-cycle heads (no short drop): size that edge by endpoint heads.
    for pipe_id, context in contexts.items():
        if pipe_id in short_drop_pipes:
            continue
        if not (context.is_cycle and context.has_discharge_endpoint):
            continue
        pipe = pipes.get(pipe_id)
        if pipe is None:
            continue
        direct = 0.0
        for node_id in (
            str(_get(pipe, "start", "") or ""),
            str(_get(pipe, "end", "") or ""),
        ):
            if node_id in nodes and _is_discharge_node(nodes, node_id):
                direct += max(
                    design_discharge_lpm(nodes[node_id], default_pressure_bar),
                    0.0,
                )
        flows[pipe_id] = max(flows.get(pipe_id, 0.0), direct)

    # Cross-mains / supply leftovers: keep estimate, raise by incident line flow.
    for pipe_id, context in contexts.items():
        if pipe_id in short_drop_pipes:
            continue
        pipe = pipes.get(pipe_id)
        if pipe is None:
            continue
        incident = 0.0
        for node_id in (
            str(_get(pipe, "start", "") or ""),
            str(_get(pipe, "end", "") or ""),
        ):
            incident = max(incident, line_flow_by_junction.get(node_id, 0.0))
        flows[pipe_id] = max(flows.get(pipe_id, 0.0), incident)

    return {str(pid): max(float(q), 0.0) for pid, q in flows.items()}



@dataclass
class PipeSizingProposal:
    pipe_id: str
    current_nominal_mm: int
    proposed_nominal_mm: int
    current_inner_d_mm: float
    proposed_inner_d_mm: float
    standard: str
    calculated_flow_lpm: float
    velocity_mps: float
    velocity_limit_mps: float
    basis: SizingBasis
    pipe_class: PipeClass
    downstream_head_count: int = 0
    reason: str = ""
    status: ChangeStatus = ChangeStatus.UNCHANGED
    valid: bool = True
    valve_adjusted: bool = False

    def refresh_status(self) -> None:
        manual_override = self.status is ChangeStatus.MANUAL_OVERRIDE
        if not self.valid:
            self.status = ChangeStatus.INVALID
        elif manual_override:
            # Recomputing display status must not erase an explicit user
            # override marker for a numerically valid, unusual manual choice.
            self.status = ChangeStatus.MANUAL_OVERRIDE
        elif self.valve_adjusted:
            self.status = ChangeStatus.VALVE_ADJUSTED
        elif self.proposed_nominal_mm < self.current_nominal_mm:
            self.status = ChangeStatus.DECREASE
        elif self.proposed_nominal_mm > self.current_nominal_mm:
            self.status = ChangeStatus.INCREASE
        else:
            self.status = ChangeStatus.UNCHANGED


@dataclass(frozen=True)
class GridSweepPositionDuty:
    """Required pump duty with the design area at one lattice position."""

    row0: int
    col0: int
    variant: int
    head_ids: tuple[str, ...]
    pump_pressure_bar: float | None
    total_flow_lpm: float | None
    critical_node_id: str = ""


@dataclass
class GridSweepReport:
    """Verification sweep on the final DN map: the duty point cloud.

    Pump selection must clear every duty point, not just the governing one —
    a curve can pass the peak pressure yet dip under a higher-flow position.
    """

    n_rows: int = 0
    n_cols: int = 0
    shape_cells: int = 0
    n_variants: int = 0
    duties: list[GridSweepPositionDuty] = field(default_factory=list)
    governing_index: int | None = None
    # Placement matching the head activity the sizing started from.
    origin_index: int | None = None
    # Every grid head the sweep toggles: relocating the design area sets
    # exactly these actives (heads outside the lattice are untouched).
    lattice_head_ids: tuple[str, ...] = ()
    demand_runs: int = 0
    sweep_rounds: int = 0

    def _duty_at(self, index: int | None) -> GridSweepPositionDuty | None:
        if index is None or not 0 <= index < len(self.duties):
            return None
        return self.duties[index]

    @property
    def governing(self) -> GridSweepPositionDuty | None:
        return self._duty_at(self.governing_index)

    @property
    def origin(self) -> GridSweepPositionDuty | None:
        return self._duty_at(self.origin_index)


@dataclass
class PipeSizingPreview:
    mode: str
    success: bool
    proposals: dict[str, PipeSizingProposal] = field(default_factory=dict)
    head_ranges: list[RangeRow] = field(default_factory=list)
    flow_ranges: list[RangeRow] = field(default_factory=list)
    valve_nominal_mm: dict[str, int] = field(default_factory=dict)
    invalid_node_ids: set[str] = field(default_factory=set)
    issues: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    iterations: int = 0
    model_signature: str = ""
    hydraulic_result: Any = field(default=None, repr=False, compare=False)
    calculation_source: str = ""
    calculation_mode: str = ""
    fallback_used: bool = False
    fallback_reason: str = ""
    velocity_limits: VelocityLimits | None = None
    grid_sweep: GridSweepReport | None = None
    # Internal diagnosis for general velocity auto sizing (analysis only).
    auto_diagnostics: Any = field(default=None, repr=False, compare=False)

    @property
    def can_apply(self) -> bool:
        return bool(self.success and self.proposals and not self.issues and all(p.valid for p in self.proposals.values()))


@dataclass
class RangeValidation:
    valid: bool
    errors: list[str] = field(default_factory=list)


def _get(obj: Any, name: str, default: Any = None) -> Any:
    if isinstance(obj, Mapping):
        return obj.get(name, default)
    return getattr(obj, name, default)


def _float(value: Any, default: float = 0.0) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError, OverflowError):
        return float(default)
    return result if math.isfinite(result) else float(default)


def flow_capacity_lpm(inner_d_mm: float, velocity_mps: float) -> float:
    """Maximum flow at velocity limit using the real internal diameter."""
    try:
        d_m = float(inner_d_mm) / 1000.0
        speed = float(velocity_mps)
    except (TypeError, ValueError, OverflowError):
        return 0.0
    if not math.isfinite(d_m) or not math.isfinite(speed):
        return 0.0
    if d_m <= 0.0 or speed <= 0.0:
        return 0.0
    return speed * math.pi * d_m * d_m / 4.0 * 60000.0


def velocity_mps(flow_lpm: float, inner_d_mm: float) -> float:
    try:
        flow = float(flow_lpm)
        diameter = float(inner_d_mm)
    except (TypeError, ValueError, OverflowError):
        return math.inf
    if not math.isfinite(flow) or not math.isfinite(diameter):
        return math.inf
    area = math.pi * (diameter / 1000.0) ** 2 / 4.0
    if area <= 0.0 or not math.isfinite(area):
        return math.inf
    return abs(flow) / 60000.0 / area


def _pipe_library_rows(pipe_library: Mapping[str, Any] | None) -> list[Mapping[str, Any]]:
    if not isinstance(pipe_library, Mapping):
        return []
    raw_rows = pipe_library.get("pipe_types", [])
    if isinstance(raw_rows, (str, bytes, Mapping)) or not isinstance(raw_rows, Sequence):
        return []
    return [row for row in raw_rows if isinstance(row, Mapping)]


def pipe_specs_for_standard(
    pipe_library: Mapping[str, Any] | None,
    standard: str,
) -> list[PipeSizeSpec]:
    result: dict[int, PipeSizeSpec] = {}
    target = str(standard or "").strip().casefold()
    for row in _pipe_library_rows(pipe_library):
        row_std = str(row.get("standard", "") or "").strip()
        if target and row_std.casefold() != target:
            continue
        raw_dn = row.get("nominal_mm", row.get("nominal_d_mm"))
        raw_id = row.get(
            "inner_d_mm", row.get("inner_mm", row.get("ID_mm", row.get("diameter")))
        )
        try:
            dn = int(raw_dn)
            inner = float(raw_id)
        except (TypeError, ValueError, OverflowError):
            continue
        spec = PipeSizeSpec(
            nominal_mm=dn,
            inner_d_mm=inner,
            standard=row_std,
            name=str(row.get("name", "") or f"{row_std} DN{dn}"),
            c_factor=_float(row.get("C", row.get("C_HW_default", 120.0)), 120.0),
            roughness_mm=_float(row.get("roughness_mm", 0.1), 0.1),
        )
        if not spec.validate():
            result[dn] = spec
    return sorted(result.values(), key=lambda item: (item.nominal_mm, item.inner_d_mm))


def available_standards(pipe_library: Mapping[str, Any] | None) -> list[str]:
    return sorted({
        str(row.get("standard", "") or "").strip()
        for row in _pipe_library_rows(pipe_library)
        if str(row.get("standard", "") or "").strip()
    })


def resolve_pipe_standard(
    pipe: Any,
    pipe_library: Mapping[str, Any] | None,
    fallback_standard: str = "",
) -> str:
    raw = str(_get(pipe, "type", "") or "").strip()
    standards = available_standards(pipe_library)
    by_fold = {item.casefold(): item for item in standards}
    if raw.casefold() in by_fold:
        return by_fold[raw.casefold()]
    for row in _pipe_library_rows(pipe_library):
        if str(row.get("name", "") or "").strip().casefold() == raw.casefold():
            return str(row.get("standard", "") or "").strip()
    if str(fallback_standard or "").strip().casefold() in by_fold:
        return by_fold[str(fallback_standard).strip().casefold()]
    # Prefer a standard that contains the current DN before falling back to the
    # first project standard.  This preserves mixed-material projects.
    current_dn = int(_float(_get(pipe, "nominal_mm", 0), 0.0))
    for standard in standards:
        if any(spec.nominal_mm == current_dn for spec in pipe_specs_for_standard(pipe_library, standard)):
            return standard
    return standards[0] if standards else raw


def select_spec_for_flow(
    specs: Sequence[PipeSizeSpec],
    flow_lpm: float,
    limit_mps: float,
    *,
    tolerance_lpm: float = FLOW_TOLERANCE_LPM,
) -> tuple[PipeSizeSpec | None, bool]:
    """Select first valid spec whose inclusive capacity contains the flow."""
    ordered = sorted(
        (spec for spec in specs if isinstance(spec, PipeSizeSpec) and not spec.validate()),
        key=lambda item: (int(item.nominal_mm), float(item.inner_d_mm)),
    )
    if not ordered:
        return None, True
    try:
        q = abs(float(flow_lpm))
        limit = float(limit_mps)
        tolerance = float(tolerance_lpm)
    except (TypeError, ValueError, OverflowError):
        return None, True
    if not all(math.isfinite(item) for item in (q, limit, tolerance)) or limit <= 0.0:
        return None, True
    for spec in ordered:
        cap = flow_capacity_lpm(spec.inner_d_mm, limit)
        tol = max(tolerance, cap * 1.0e-9)
        if q <= cap + tol:
            return spec, False
    return ordered[-1], True


def spec_for_nominal(specs: Sequence[PipeSizeSpec], nominal_mm: int) -> PipeSizeSpec | None:
    try:
        target = int(nominal_mm)
    except (TypeError, ValueError, OverflowError):
        return None
    return next(
        (
            item for item in specs
            if isinstance(item, PipeSizeSpec)
            and not item.validate()
            and int(item.nominal_mm) == target
        ),
        None,
    )


def build_flow_ranges(specs: Sequence[PipeSizeSpec], limit_mps: float) -> list[RangeRow]:
    ordered = sorted(
        (spec for spec in specs if isinstance(spec, PipeSizeSpec) and not spec.validate()),
        key=lambda item: (int(item.nominal_mm), float(item.inner_d_mm)),
    )
    try:
        limit = float(limit_mps)
    except (TypeError, ValueError, OverflowError):
        return []
    if not ordered or not math.isfinite(limit) or limit <= 0.0:
        return []
    rows: list[RangeRow] = []
    start = 0.0
    for index, spec in enumerate(ordered):
        end = None if index == len(ordered) - 1 else flow_capacity_lpm(spec.inner_d_mm, limit)
        rows.append(RangeRow(start=start, end=end, nominal_mm=int(spec.nominal_mm)))
        if end is not None:
            start = end
    return rows


def _coerce_range_row(raw: Any) -> tuple[RangeRow, list[str]]:
    row = raw if isinstance(raw, RangeRow) else RangeRow.from_dict(raw)
    row_errors = list(row._input_errors)
    try:
        start = float(row.start)
    except (TypeError, ValueError, OverflowError):
        start = math.nan
        row_errors.append("range_start_not_numeric")
    if row.end is None:
        end = None
    else:
        try:
            end = float(row.end)
        except (TypeError, ValueError, OverflowError):
            end = math.nan
            row_errors.append("range_end_not_numeric")
    try:
        nominal = int(row.nominal_mm)
    except (TypeError, ValueError, OverflowError):
        nominal = 0
        row_errors.append("range_dn_not_integer")
    return RangeRow(start, end, nominal, tuple(dict.fromkeys(row_errors))), row_errors


def validate_ranges(
    rows: Sequence[RangeRow],
    *,
    integer_bounds: bool = False,
    require_unbounded_last: bool = False,
    required_values: Iterable[float] = (),
    expected_first_start: float | None = None,
) -> RangeValidation:
    errors: list[str] = []
    try:
        raw_rows = list(rows)
    except TypeError:
        return RangeValidation(False, ["range_table_not_iterable"])
    if not raw_rows:
        return RangeValidation(False, ["range_table_empty"])

    ordered: list[RangeRow] = []
    for idx, raw in enumerate(raw_rows):
        row, input_errors = _coerce_range_row(raw)
        ordered.append(row)
        errors.extend(f"{code}:{idx}" for code in input_errors)

    expected_first: float | None = None
    if expected_first_start is not None:
        try:
            expected_first = float(expected_first_start)
        except (TypeError, ValueError, OverflowError):
            errors.append("range_expected_first_start_not_numeric")
        if expected_first is not None and not math.isfinite(expected_first):
            errors.append("range_expected_first_start_not_finite")
            expected_first = None
    if expected_first is not None:
        tolerance = 1.0e-9 if integer_bounds else FLOW_TOLERANCE_LPM
        if not math.isfinite(ordered[0].start) or not math.isclose(
            ordered[0].start, expected_first, abs_tol=tolerance, rel_tol=0.0
        ):
            errors.append(f"range_first_start_must_be:{expected_first:g}")

    for idx, row in enumerate(ordered):
        if row.nominal_mm <= 0:
            errors.append(f"range_invalid_dn:{idx}")
        start_valid = math.isfinite(row.start) and row.start >= 0.0
        if not start_valid:
            errors.append(f"range_invalid_start:{idx}")
        elif integer_bounds and not math.isclose(row.start, round(row.start), abs_tol=1.0e-9):
            errors.append(f"range_non_integer_start:{idx}")

        end_valid = row.end is None or (
            math.isfinite(row.end) and start_valid and row.end >= row.start
        )
        if row.end is not None:
            if not end_valid:
                errors.append(f"range_invalid_end:{idx}")
            elif integer_bounds and not math.isclose(row.end, round(row.end), abs_tol=1.0e-9):
                errors.append(f"range_non_integer_end:{idx}")
        if idx < len(ordered) - 1 and row.end is None:
            errors.append(f"range_unbounded_not_last:{idx}")
        if idx == 0:
            continue
        prev = ordered[idx - 1]
        if prev.end is None or not math.isfinite(prev.end) or not start_valid:
            continue
        expected_start = prev.end + (1.0 if integer_bounds else 0.0)
        tolerance = 1.0e-9 if integer_bounds else FLOW_TOLERANCE_LPM
        if not math.isclose(row.start, expected_start, abs_tol=tolerance, rel_tol=0.0):
            relation = "overlap" if row.start < expected_start else "gap"
            errors.append(f"range_{relation}:{idx - 1}:{idx}")

    if require_unbounded_last and ordered[-1].end is not None:
        errors.append("range_last_not_unbounded")
    for index, value in enumerate(required_values):
        try:
            numeric = float(value)
        except (TypeError, ValueError, OverflowError):
            errors.append(f"range_required_value_not_numeric:{index}")
            continue
        lookup = lookup_head_range if integer_bounds else lookup_range
        if not math.isfinite(numeric) or lookup(ordered, numeric) is None:
            errors.append(f"range_value_uncovered:{numeric:g}")
    return RangeValidation(not errors, list(dict.fromkeys(errors)))


def lookup_range(rows: Sequence[RangeRow], value: float) -> RangeRow | None:
    try:
        numeric = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    if not math.isfinite(numeric):
        return None
    for idx, raw in enumerate(rows):
        row, input_errors = _coerce_range_row(raw)
        if input_errors or not math.isfinite(row.start):
            continue
        if row.contains(numeric, first=(idx == 0)):
            return raw if isinstance(raw, RangeRow) else row
    return None


def lookup_head_range(rows: Sequence[RangeRow], value: float) -> RangeRow | None:
    """Look up an integer head count using inclusive, non-shared bounds.

    If the last row has a finite end and ``value`` is above it, that last row
    is reused (open-ended extension). Gaps below the last end still miss.
    """
    try:
        numeric = float(value)
    except (TypeError, ValueError, OverflowError):
        return None
    if not math.isfinite(numeric) or not math.isclose(
        numeric, round(numeric), abs_tol=1.0e-9
    ):
        return None
    last_raw: RangeRow | None = None
    for raw in rows:
        row, input_errors = _coerce_range_row(raw)
        if input_errors or not math.isfinite(row.start):
            continue
        last_raw = raw if isinstance(raw, RangeRow) else row
        lower_ok = numeric >= row.start
        upper_ok = row.end is None or (math.isfinite(row.end) and numeric <= row.end)
        if lower_ok and upper_ok:
            return raw if isinstance(raw, RangeRow) else row
    if last_raw is None:
        return None
    last_row, last_errors = _coerce_range_row(last_raw)
    if last_errors or not math.isfinite(last_row.start):
        return None
    if last_row.end is None:
        return None
    if math.isfinite(last_row.end) and numeric > last_row.end:
        return last_raw if isinstance(last_raw, RangeRow) else last_row
    return None


def minimum_discharge_lpm(node: Any, default_pressure_bar: float) -> float:
    """Active sprinkler/nozzle demand used by the hydraulic sizing model.

    Inactive heads contribute zero flow here, but still count toward design
    head totals via ``design_discharge_lpm`` / discharge summaries.
    """
    if not bool(_get(node, "is_active", True)):
        return 0.0
    return design_discharge_lpm(node, default_pressure_bar)


def hose_stream_design_flow_lpm(node: Any) -> float:
    """Design flow of a hose-stream (옥내소화전) node, ignoring ``is_active``."""
    type_id = str(_get(node, "type_id", "") or "").strip().lower()
    if type_id != NodeTypeId.HOSE_STREAM.value:
        return 0.0
    return max(_float(_get(node, "hose_stream_flow_lps", 0.0)), 0.0) * 60.0


def hose_stream_sizing_flow_lpm(node: Any) -> float:
    """Hose-stream flow that participates in sizing (active nodes only)."""
    if not bool(_get(node, "is_active", True)):
        return 0.0
    return hose_stream_design_flow_lpm(node)


def design_discharge_lpm(node: Any, default_pressure_bar: float) -> float:
    """Design demand for a head/nozzle, ignoring ``is_active``.

    Hose streams stay zero here: their fixed flow enters sizing through
    ``hose_stream_sizing_flow_lpm`` and must not affect head counts.
    """
    type_id = str(_get(node, "type_id", "") or "").strip().lower()
    if type_id == NodeTypeId.HOSE_STREAM.value:
        return 0.0
    if type_id not in {NodeTypeId.HEAD.value, NodeTypeId.NOZZLE.value}:
        return 0.0
    fixed = max(_float(_get(node, "base_demand_lps", 0.0)), 0.0) * 60.0
    k = max(_float(_get(node, "k_factor_si", 0.0)), 0.0)
    req = _float(_get(node, "required_pressure_bar", 0.0))
    if req <= 0.0:
        req = max(_float(default_pressure_bar), 0.0)
    return max(fixed, k * math.sqrt(req))


def _predicate_view(node: Any) -> Any:
    """Adapt legacy mapping nodes to the shared node-predicate interface."""
    if isinstance(node, Mapping):
        return SimpleNamespace(**dict(node))
    return node


def _is_non_pump_hydraulic_source(node: Any) -> bool:
    view = _predicate_view(node)
    tid = str(_get(node, "type_id", "") or "").strip().lower()
    text = str(_get(node, "type", "") or "").strip().lower()
    cat = str(_get(node, "category_id", "") or "").strip().lower()
    if is_tank_graph_node(view) or tid in {
        NodeTypeId.WT.value,
        NodeTypeId.RES.value,
    }:
        return True
    # Legacy projects can carry a tank as a base node with only water_level;
    # this mirrors domain.topology's hydraulic-source compatibility rule.
    if _float(_get(node, "water_level", 0.0)) > 0.0:
        return True
    # ``source`` is a supported legacy identifier but is not a NodeTypeId.
    return any("source" in value for value in (tid, text, cat))


def _source_pump_ids(
    nodes: Mapping[str, Any], pipes: Mapping[str, Any]
) -> list[str]:
    """Return topology source pumps (not inline/booster pumps)."""
    pump_ids = [
        str(node_id)
        for node_id, node in nodes.items()
        if is_pump_graph_node(_predicate_view(node))
        or str(_get(node, "type_id", "") or "").strip().lower()
        == NodeTypeId.PUMP.value
    ]
    topology_pipes = {
        str(pipe_id): SimpleNamespace(
            start=str(_get(pipe, "start", "") or ""),
            end=str(_get(pipe, "end", "") or ""),
        )
        for pipe_id, pipe in pipes.items()
    }
    topology = analyze_pump_topology(pump_ids, topology_pipes)
    return [str(pid) for pid in topology.get("source_pumps", ())]


def _is_alarm_valve(node: Any) -> bool:
    """Backward-compatible alarm-valve-only check."""
    values = " ".join(
        str(_get(node, key, "") or "").strip().lower()
        for key in ("type_id", "type", "category_id", "fitting_id")
    )
    return (
        "alarm_valve" in values
        or "alarm valve" in values
        or "알람밸브" in values
    )


def _is_discharge_control_valve(node: Any) -> bool:
    """Alarm / dry-pipe / preaction valve used as discharge-side stop."""
    values = " ".join(
        str(_get(node, key, "") or "").strip().lower()
        for key in ("type_id", "type", "category_id", "fitting_id")
    )
    if not values:
        return False
    return any(token in values for token in _DISCHARGE_CONTROL_VALVE_TOKENS)


def _node_pipe_degree(
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
) -> dict[str, int]:
    return {
        str(node_id): len(edges)
        for node_id, edges in adjacency.items()
    }


def _pump_suction_neighbor_ids(
    nodes: Mapping[str, Any],
    pipes: Mapping[str, Any],
    pump_ids: Sequence[str],
) -> set[str]:
    """Tank/source neighbors attached directly to a pump (흡입측 후보)."""
    pumps = {str(pid) for pid in pump_ids}
    suction: set[str] = set()
    for pipe in pipes.values():
        start = str(_get(pipe, "start", "") or "")
        end = str(_get(pipe, "end", "") or "")
        if start in pumps and end in nodes and _is_non_pump_hydraulic_source(
            nodes[end]
        ):
            suction.add(end)
        if end in pumps and start in nodes and _is_non_pump_hydraulic_source(
            nodes[start]
        ):
            suction.add(start)
    return suction


def _reachable_from_pumps(
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
    pump_ids: Sequence[str],
    *,
    blocked_first_hop: set[str] | None = None,
) -> dict[str, int]:
    """Multi-source BFS distances from pumps; optional first-hop exclusions."""
    pumps = [str(pid) for pid in pump_ids if str(pid) in adjacency]
    if not pumps:
        return {}
    blocked = {str(nid) for nid in (blocked_first_hop or ())}
    distance: dict[str, int] = {pid: 0 for pid in pumps}
    queue: deque[str] = deque(pumps)
    while queue:
        current = queue.popleft()
        current_distance = distance[current]
        for neighbor, _pipe_id in adjacency.get(current, ()):
            neighbor_id = str(neighbor)
            if neighbor_id in distance:
                continue
            if current_distance == 0 and neighbor_id in blocked:
                continue
            distance[neighbor_id] = current_distance + 1
            queue.append(neighbor_id)
    return distance


def _discharge_side_tank_ids(
    nodes: Mapping[str, Any],
    pipes: Mapping[str, Any],
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
    pump_ids: Sequence[str],
) -> list[str]:
    """Elevated/discharge tanks reachable from pumps, excluding suction WT."""
    suction = _pump_suction_neighbor_ids(nodes, pipes, pump_ids)
    reachable = _reachable_from_pumps(
        adjacency, pump_ids, blocked_first_hop=suction
    )
    return sorted(
        node_id
        for node_id in reachable
        if node_id not in {str(pid) for pid in pump_ids}
        and node_id in nodes
        and _is_non_pump_hydraulic_source(nodes[node_id])
    )


def _first_tee_node_ids(
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
    pump_ids: Sequence[str],
) -> list[str]:
    """Closest degree>=3 junctions downstream of pumps (펌프 자신 제외)."""
    degree = _node_pipe_degree(adjacency)
    distance = _reachable_from_pumps(adjacency, pump_ids)
    candidates = [
        node_id
        for node_id, dist in distance.items()
        if dist > 0 and degree.get(node_id, 0) >= 3
    ]
    if not candidates:
        return []
    best = min(distance[node_id] for node_id in candidates)
    return sorted(
        node_id for node_id in candidates if distance[node_id] == best
    )


def _discharge_control_valve_ids(nodes: Mapping[str, Any]) -> list[str]:
    return sorted(
        str(node_id)
        for node_id, node in nodes.items()
        if _is_discharge_control_valve(node)
    )


def _discharge_origin_pump_ids(
    nodes: Mapping[str, Any], pipes: Mapping[str, Any]
) -> list[str]:
    """Pumps that start the discharge-side supply walk.

    Prefer topology source pumps. When the only pump sits after a suction WT
    (treated as inline/booster), fall back to that pump so pump→first-tee still
    works.
    """
    sources = _source_pump_ids(nodes, pipes)
    if sources:
        return sources
    return [
        str(node_id)
        for node_id, node in nodes.items()
        if is_pump_graph_node(_predicate_view(node))
        or str(_get(node, "type_id", "") or "").strip().lower()
        == NodeTypeId.PUMP.value
    ]


def _shortest_node_path(
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
    source: str,
    target: str,
) -> list[str] | None:
    """BFS node path from source to target, or None if unreachable."""
    source_id = str(source)
    target_id = str(target)
    if source_id == target_id:
        return [source_id]
    if source_id not in adjacency or target_id not in adjacency:
        return None
    parent: dict[str, str | None] = {source_id: None}
    queue: deque[str] = deque([source_id])
    while queue:
        current = queue.popleft()
        for neighbor, _pipe_id in adjacency.get(current, ()):
            neighbor_id = str(neighbor)
            if neighbor_id in parent:
                continue
            parent[neighbor_id] = current
            if neighbor_id == target_id:
                path = [target_id]
                while path[-1] != source_id:
                    prev = parent[path[-1]]
                    if prev is None:
                        return None
                    path.append(prev)
                path.reverse()
                return path
            queue.append(neighbor_id)
    return None


def _is_local_pump_discharge_target(
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
    degree: Mapping[str, int],
    pump_id: str,
    target_id: str,
) -> bool:
    """True when target sits on this pump's local discharge manifold.

    A remote tank/valve reached through the distribution corridor (more than
    one degree>=3 hub on the shortest path) is rejected so parallel pumps do
    not pull each other's mains into SUPPLY.
    """
    path = _shortest_node_path(adjacency, pump_id, target_id)
    if not path:
        return False
    hubs = [
        node_id
        for node_id in path[1:-1]
        if int(degree.get(node_id, 0) or 0) >= 3
    ]
    return len(hubs) <= 1


def _discharge_supply_pipe_ids(
    nodes: Mapping[str, Any],
    pipes: Mapping[str, Any],
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
) -> set[str]:
    """Per-pump discharge supply: local tank > first tee > local valve.

    Each origin pump is classified independently. A local elevated tank on one
    pump must not pull another pump's distribution corridor into SUPPLY.
    Intentionally simple — users may override borderline segments manually.
    """
    pump_ids = _discharge_origin_pump_ids(nodes, pipes)
    if not pump_ids:
        return set()

    degree = _node_pipe_degree(adjacency)
    valve_ids = _discharge_control_valve_ids(nodes)
    supply: set[str] = set()

    for pump_id in pump_ids:
        single = [pump_id]
        tank_ids = _discharge_side_tank_ids(nodes, pipes, adjacency, single)
        local_tanks = [
            tank_id
            for tank_id in tank_ids
            if _is_local_pump_discharge_target(
                adjacency, degree, pump_id, tank_id
            )
        ]
        if local_tanks:
            supply |= _pipes_on_any_source_target_path(
                nodes, pipes, adjacency, single, local_tanks
            )
            continue

        tee_ids = _first_tee_node_ids(adjacency, single)
        if tee_ids:
            supply |= _pipes_on_any_source_target_path(
                nodes, pipes, adjacency, single, tee_ids
            )
            continue

        local_valves = [
            valve_id
            for valve_id in valve_ids
            if _is_local_pump_discharge_target(
                adjacency, degree, pump_id, valve_id
            )
        ]
        if local_valves:
            supply |= _pipes_on_any_source_target_path(
                nodes, pipes, adjacency, single, local_valves
            )

    return supply


def _hydraulic_source_ids(
    nodes: Mapping[str, Any], pipes: Mapping[str, Any]
) -> set[str]:
    """Resolve sources through the existing topology/type SSOTs.

    A pump with both an inlet and outlet is an inline/booster pump and must not
    become a second source. Tanks and reservoirs are sources in their own
    right, including the common WT -> pump -> system arrangement.
    """
    sources = {
        str(node_id)
        for node_id, node in nodes.items()
        if _is_non_pump_hydraulic_source(node)
    }
    sources.update(_source_pump_ids(nodes, pipes))
    return sources


def _discharge_signature(node: Any, default_pressure_bar: float) -> tuple[float, float]:
    # Design demand (active or not) so inactive heads with the same K stay in
    # the HEAD_COUNT family instead of looking "mixed" against open heads.
    return (
        round(max(_float(_get(node, "k_factor_si", 0.0)), 0.0), 6),
        round(design_discharge_lpm(node, default_pressure_bar), 6),
    )


def _graph_parts(graph: Any) -> tuple[dict[str, Any], dict[str, Any], dict[str, list[tuple[str, str]]]]:
    nodes = dict(_get(graph, "nodes", {}) or {})
    pipes = dict(_get(graph, "pipes", {}) or {})
    adjacency: dict[str, list[tuple[str, str]]] = {str(node_id): [] for node_id in nodes}
    for pipe_id, pipe in pipes.items():
        start = str(_get(pipe, "start", "") or "")
        end = str(_get(pipe, "end", "") or "")
        if start not in adjacency or end not in adjacency or start == end:
            continue
        adjacency[start].append((end, str(pipe_id)))
        adjacency[end].append((start, str(pipe_id)))
    return nodes, pipes, adjacency


def find_bridge_pipe_ids(graph: Any) -> set[str]:
    """Iterative Tarjan bridge detection for arbitrarily deep networks."""
    nodes, _pipes, adjacency = _graph_parts(graph)
    timer = 0
    tin: dict[str, int] = {}
    low: dict[str, int] = {}
    bridges: set[str] = set()

    for node_id in nodes:
        root = str(node_id)
        if root in tin:
            continue
        timer += 1
        tin[root] = low[root] = timer
        parent_node: dict[str, str] = {}
        # Mutable frame: node, parent pipe ID, next adjacency index.
        stack: list[list[Any]] = [[root, None, 0]]
        while stack:
            current, parent_pipe_id, next_index = stack[-1]
            edges = adjacency.get(current, ())
            if next_index < len(edges):
                neighbor, pipe_id = edges[next_index]
                stack[-1][2] = next_index + 1
                if pipe_id == parent_pipe_id:
                    continue
                if neighbor in tin:
                    low[current] = min(low[current], tin[neighbor])
                    continue
                parent_node[neighbor] = current
                timer += 1
                tin[neighbor] = low[neighbor] = timer
                stack.append([neighbor, pipe_id, 0])
                continue

            stack.pop()
            if parent_pipe_id is None:
                continue
            parent = parent_node[current]
            low[parent] = min(low[parent], low[current])
            if low[current] > tin[parent]:
                bridges.add(str(parent_pipe_id))
    return bridges


def _multi_source_distance(
    adjacency: Mapping[str, Sequence[tuple[str, str]]], sources: Sequence[str]
) -> tuple[dict[str, int], dict[str, tuple[str, str]]]:
    distance: dict[str, int] = {}
    parent: dict[str, tuple[str, str]] = {}
    queue: deque[str] = deque()
    for source in sources:
        distance[source] = 0
        queue.append(source)
    while queue:
        current = queue.popleft()
        for neighbor, pipe_id in adjacency.get(current, ()):
            if neighbor in distance:
                continue
            distance[neighbor] = distance[current] + 1
            parent[neighbor] = (current, pipe_id)
            queue.append(neighbor)
    return distance, parent


def _biconnected_pipe_components(
    nodes: Mapping[str, Any],
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
) -> list[set[str]]:
    """Return vertex-biconnected edge blocks with iterative Tarjan DFS.

    A bridge is represented as a one-edge block. Parallel pipes are handled by
    identifying the DFS parent by pipe ID instead of endpoint. The explicit
    stack avoids Python's recursion limit on long risers and trunk lines.
    """
    timer = 0
    discovered: dict[str, int] = {}
    low: dict[str, int] = {}
    edge_stack: list[str] = []
    components: list[set[str]] = []

    def pop_component(stop_pipe_id: str) -> None:
        component: set[str] = set()
        while edge_stack:
            pipe_id = edge_stack.pop()
            component.add(pipe_id)
            if pipe_id == stop_pipe_id:
                break
        if component:
            components.append(component)

    for node_id in nodes:
        root = str(node_id)
        if root in discovered:
            continue
        timer += 1
        discovered[root] = low[root] = timer
        parent_node: dict[str, str] = {}
        stack: list[list[Any]] = [[root, None, 0]]
        while stack:
            current, parent_pipe_id, next_index = stack[-1]
            edges = adjacency.get(current, ())
            if next_index < len(edges):
                neighbor, pipe_id = edges[next_index]
                stack[-1][2] = next_index + 1
                if pipe_id == parent_pipe_id:
                    continue
                if neighbor not in discovered:
                    parent_node[neighbor] = current
                    edge_stack.append(pipe_id)
                    timer += 1
                    discovered[neighbor] = low[neighbor] = timer
                    stack.append([neighbor, pipe_id, 0])
                    continue
                if discovered[neighbor] < discovered[current]:
                    # Record an undirected back edge once, descendant to
                    # ancestor. A parallel edge to the parent lands here too.
                    edge_stack.append(pipe_id)
                    low[current] = min(low[current], discovered[neighbor])
                continue

            stack.pop()
            if parent_pipe_id is None:
                continue
            parent = parent_node[current]
            low[parent] = min(low[parent], low[current])
            if low[current] >= discovered[parent]:
                pop_component(str(parent_pipe_id))

        if edge_stack:
            pop_component(edge_stack[0])
    return components


def _pipes_on_any_source_target_path(
    nodes: Mapping[str, Any],
    pipes: Mapping[str, Any],
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
    sources: Sequence[str],
    targets: Sequence[str],
) -> set[str]:
    """Return edges on any simple source-to-target path without enumeration.

    Tarjan blocks are contracted into a block-cut forest. A forest edge is on a
    source-target path exactly when deleting it separates at least one source
    from at least one target. Every edge in a selected biconnected block lies on
    a simple path between the block's relevant boundary vertices.
    """
    source_keys = {("node", str(node_id)) for node_id in sources}
    target_keys = {("node", str(node_id)) for node_id in targets}
    if not source_keys or not target_keys:
        return set()

    components = _biconnected_pipe_components(nodes, adjacency)
    tree: dict[tuple[str, str], list[tuple[str, str]]] = {}
    block_edges: dict[tuple[str, str], set[str]] = {}
    for index, component in enumerate(components):
        block_key = ("block", str(index))
        block_edges[block_key] = set(component)
        vertices: set[str] = set()
        for pipe_id in component:
            pipe = pipes.get(pipe_id)
            if pipe is None:
                continue
            vertices.add(str(_get(pipe, "start", "") or ""))
            vertices.add(str(_get(pipe, "end", "") or ""))
        for node_id in vertices:
            node_key = ("node", node_id)
            tree.setdefault(block_key, []).append(node_key)
            tree.setdefault(node_key, []).append(block_key)

    selected_blocks: set[tuple[str, str]] = set()
    visited: set[tuple[str, str]] = set()
    for root in tree:
        if root in visited:
            continue
        parent: dict[tuple[str, str], tuple[str, str] | None] = {root: None}
        order: list[tuple[str, str]] = []
        queue: deque[tuple[str, str]] = deque([root])
        visited.add(root)
        while queue:
            current = queue.popleft()
            order.append(current)
            for neighbor in tree.get(current, ()):
                if neighbor in visited:
                    continue
                visited.add(neighbor)
                parent[neighbor] = current
                queue.append(neighbor)

        subtree_sources = {key: int(key in source_keys) for key in order}
        subtree_targets = {key: int(key in target_keys) for key in order}
        for key in reversed(order[1:]):
            parent_key = parent[key]
            if parent_key is not None:
                subtree_sources[parent_key] += subtree_sources[key]
                subtree_targets[parent_key] += subtree_targets[key]

        total_sources = subtree_sources[root]
        total_targets = subtree_targets[root]
        if total_sources <= 0 or total_targets <= 0:
            continue
        for key in order[1:]:
            source_below = subtree_sources[key]
            target_below = subtree_targets[key]
            separates_pair = (
                source_below > 0 and total_targets - target_below > 0
            ) or (
                target_below > 0 and total_sources - source_below > 0
            )
            if not separates_pair:
                continue
            parent_key = parent[key]
            if key[0] == "block":
                selected_blocks.add(key)
            if parent_key is not None and parent_key[0] == "block":
                selected_blocks.add(parent_key)

    result: set[str] = set()
    for block_key in selected_blocks:
        result.update(block_edges.get(block_key, ()))
    return result


class _LazyDownstreamNodeIds(Sequence[str]):
    """Tuple-compatible lazy view over one side of a bridge forest edge.

    Most sizing callers only need aggregated flow and count. Deferring the
    potentially quadratic *output* of every downstream-node tuple keeps normal
    analysis O(V + E), while iteration/indexing still provides the historical
    sorted node-ID sequence on demand.
    """

    __slots__ = (
        "_component_nodes",
        "_ordered_components",
        "_start",
        "_stop",
        "_use_subtree",
        "_size",
        "_cache",
    )

    def __init__(
        self,
        component_nodes: Sequence[Sequence[str]],
        ordered_components: Sequence[int],
        start: int,
        stop: int,
        use_subtree: bool,
        size: int,
    ) -> None:
        self._component_nodes = component_nodes
        self._ordered_components = ordered_components
        self._start = int(start)
        self._stop = int(stop)
        self._use_subtree = bool(use_subtree)
        self._size = int(size)
        self._cache: tuple[str, ...] | None = None

    def _values(self) -> tuple[str, ...]:
        if self._cache is None:
            if self._use_subtree:
                selected = self._ordered_components[self._start:self._stop]
            else:
                selected = (
                    list(self._ordered_components[:self._start])
                    + list(self._ordered_components[self._stop:])
                )
            self._cache = tuple(sorted(
                node_id
                for component_id in selected
                for node_id in self._component_nodes[component_id]
            ))
        return self._cache

    def __len__(self) -> int:
        return self._size

    def __getitem__(self, index):
        return self._values()[index]

    def __iter__(self):
        return iter(self._values())

    def __eq__(self, other: object) -> bool:
        if isinstance(other, (str, bytes)):
            return False
        try:
            return self._values() == tuple(other)  # type: ignore[arg-type]
        except TypeError:
            return False

    def __repr__(self) -> str:
        return repr(self._values()) if self._cache is not None else f"<lazy node ids: {self._size}>"


@dataclass(frozen=True)
class _DischargeSummary:
    count: int = 0
    active_count: int = 0
    flow_lpm: float = 0.0
    signature: tuple[float, float] | None = None
    mixed: bool = False
    hose_count: int = 0
    active_hose_count: int = 0


def _combine_discharge(
    left: _DischargeSummary, right: _DischargeSummary
) -> _DischargeSummary:
    # Hose streams contribute flow and their own counts, but never join the
    # head signature/mixed decision: head-count tables stay a sprinkler
    # concept even on pipes that also feed a hose outlet.
    left_heads = left.count > 0 or left.active_count > 0
    right_heads = right.count > 0 or right.active_count > 0
    if left_heads and right_heads:
        mixed = (
            left.mixed
            or right.mixed
            or left.signature is None
            or right.signature is None
            or left.signature != right.signature
        )
        signature = None if mixed else left.signature
    elif left_heads:
        mixed, signature = left.mixed, left.signature
    elif right_heads:
        mixed, signature = right.mixed, right.signature
    else:
        mixed, signature = False, None
    return _DischargeSummary(
        count=left.count + right.count,
        active_count=left.active_count + right.active_count,
        flow_lpm=left.flow_lpm + right.flow_lpm,
        signature=signature,
        mixed=mixed,
        hose_count=left.hose_count + right.hose_count,
        active_hose_count=left.active_hose_count + right.active_hose_count,
    )


def _node_discharge_summary(
    node: Any, default_pressure_bar: float
) -> _DischargeSummary | None:
    """Count every head/nozzle/hose for design sizing; flow stays active-only."""
    type_id = str(_get(node, "type_id", "") or "").strip().lower()
    if type_id == NodeTypeId.HOSE_STREAM.value:
        hose_flow = hose_stream_sizing_flow_lpm(node)
        return _DischargeSummary(
            flow_lpm=hose_flow,
            hose_count=1,
            active_hose_count=1 if hose_flow > 0.0 else 0,
        )
    if type_id not in {NodeTypeId.HEAD.value, NodeTypeId.NOZZLE.value}:
        return None
    active_flow = minimum_discharge_lpm(node, default_pressure_bar)
    return _DischargeSummary(
        count=1,
        active_count=1 if active_flow > 0.0 else 0,
        flow_lpm=active_flow,
        signature=_discharge_signature(node, default_pressure_bar),
    )


def _bridge_component_index(
    nodes: Mapping[str, Any],
    adjacency: Mapping[str, Sequence[tuple[str, str]]],
    bridge_ids: set[str],
) -> tuple[
    dict[str, int],
    list[list[str]],
    dict[int, list[tuple[int, str]]],
]:
    """Contract non-bridge regions into a bridge forest in O(V + E)."""
    component_of: dict[str, int] = {}
    component_nodes: list[list[str]] = []
    for raw_node_id in nodes:
        node_id = str(raw_node_id)
        if node_id in component_of:
            continue
        component_id = len(component_nodes)
        members: list[str] = []
        component_of[node_id] = component_id
        queue: deque[str] = deque([node_id])
        while queue:
            current = queue.popleft()
            members.append(current)
            for neighbor, pipe_id in adjacency.get(current, ()):
                if pipe_id in bridge_ids or neighbor in component_of:
                    continue
                component_of[neighbor] = component_id
                queue.append(neighbor)
        component_nodes.append(members)

    forest: dict[int, list[tuple[int, str]]] = {
        component_id: [] for component_id in range(len(component_nodes))
    }
    seen_bridges: set[str] = set()
    for node_id, edges in adjacency.items():
        for neighbor, pipe_id in edges:
            if pipe_id not in bridge_ids or pipe_id in seen_bridges:
                continue
            seen_bridges.add(pipe_id)
            left = component_of[node_id]
            right = component_of[neighbor]
            if left == right:
                continue
            forest[left].append((right, pipe_id))
            forest[right].append((left, pipe_id))
    return component_of, component_nodes, forest


def _pipe_length_m(pipe: Any) -> float:
    return max(_float(_get(pipe, "length_m", 0.0), 0.0), 0.0)


def _is_discharge_node(nodes: Mapping[str, Any], node_id: str) -> bool:
    type_id = str(_get(nodes.get(node_id), "type_id", "") or "").strip().lower()
    return type_id in {NodeTypeId.HEAD.value, NodeTypeId.NOZZLE.value}


def _short_drop_junction_ids(
    nodes: Mapping[str, Any], pipes: Mapping[str, Any]
) -> set[str]:
    """Junction nodes carrying a head/nozzle on a short drop (≤ 1 m nipple)."""
    result: set[str] = set()
    for pipe in pipes.values():
        start = str(_get(pipe, "start", "") or "")
        end = str(_get(pipe, "end", "") or "")
        if start not in nodes or end not in nodes or start == end:
            continue
        start_d = _is_discharge_node(nodes, start)
        end_d = _is_discharge_node(nodes, end)
        if start_d == end_d:
            continue
        if _pipe_length_m(pipe) > _HEAD_DROP_MAX_M:
            continue
        result.add(end if start_d else start)
    return result


def analyze_pipe_contexts(
    graph: Any,
    *,
    default_pressure_bar: float = 1.0,
) -> tuple[dict[str, PipeSizingContext], list[str]]:
    """Classify pipes with iterative, near-linear topology aggregation.

    The bridge forest is walked once. Per-pipe downstream flow and head counts
    then come from subtree/outside summaries instead of two fresh BFS walks per
    bridge. ``downstream_node_ids`` is a lazy compatibility view; its sorted
    tuple is materialized only when a caller explicitly iterates or indexes
    that output.

    Non-supply pipes are always BRANCH. Cycles keep FLOW basis; tree pipes
    with heads (and no mixed demand) use HEAD_COUNT. MAIN is not assigned.
    """
    nodes, pipes, adjacency = _graph_parts(graph)
    issues: list[str] = []
    hydraulic_sources = sorted(_hydraulic_source_ids(nodes, pipes))
    source_distance, _ = _multi_source_distance(adjacency, hydraulic_sources)
    sources = list(hydraulic_sources)
    if not hydraulic_sources:
        issues.append("hydraulic_source_missing")
        discharge_types = {
            NodeTypeId.HEAD.value,
            NodeTypeId.NOZZLE.value,
            NodeTypeId.HOSE_STREAM.value,
        }
        leaf_candidates = [
            node_id
            for node_id, edges in adjacency.items()
            if len(edges) <= 1
            and str(_get(nodes.get(node_id), "type_id", "") or "").lower()
            not in discharge_types
        ]
        if leaf_candidates:
            sources = [sorted(leaf_candidates)[0]]
        elif nodes:
            sources = [sorted(nodes)[0]]

    distance, _parent = _multi_source_distance(adjacency, sources)
    bridge_ids = find_bridge_pipe_ids(graph)
    supply_pipe_ids = _discharge_supply_pipe_ids(nodes, pipes, adjacency)
    short_drop_junctions = _short_drop_junction_ids(nodes, pipes)

    for node_id, node in nodes.items():
        if not bool(_get(node, "is_active", True)):
            continue
        type_id = str(_get(node, "type_id", "") or "").strip().lower()
        if type_id in {NodeTypeId.HEAD.value, NodeTypeId.NOZZLE.value} and (
            minimum_discharge_lpm(node, default_pressure_bar) <= 0.0
        ):
            issues.append(f"discharge_flow_missing:{node_id}")
        if (
            hydraulic_sources
            and type_id in {NodeTypeId.HEAD.value, NodeTypeId.NOZZLE.value}
            and str(node_id) not in source_distance
        ):
            issues.append(f"discharge_source_unreachable:{node_id}")

    component_of, component_nodes, forest = _bridge_component_index(
        nodes, adjacency, bridge_ids
    )
    empty = _DischargeSummary()
    base_summary: dict[int, _DischargeSummary] = {}
    base_source_count: dict[int, int] = {}
    for component_id, member_ids in enumerate(component_nodes):
        summary = empty
        for node_id in member_ids:
            contribution = _node_discharge_summary(
                nodes[node_id], default_pressure_bar
            )
            if contribution is None:
                continue
            summary = _combine_discharge(summary, contribution)
        base_summary[component_id] = summary
        base_source_count[component_id] = sum(
            1 for node_id in member_ids if node_id in sources
        )

    parent: dict[int, int | None] = {}
    parent_bridge: dict[int, str] = {}
    tree_root: dict[int, int] = {}
    entry: dict[int, int] = {}
    exit_index: dict[int, int] = {}
    tree_components: dict[int, list[int]] = {}
    children: dict[int, list[int]] = {
        component_id: [] for component_id in range(len(component_nodes))
    }

    for start_component in range(len(component_nodes)):
        if start_component in parent:
            continue
        root = start_component
        parent[root] = None
        local_order: list[int] = []
        stack: list[list[int | None]] = [[root, None, 0]]
        while stack:
            current = int(stack[-1][0])
            parent_id = stack[-1][1]
            next_index = int(stack[-1][2])
            if next_index == 0:
                tree_root[current] = root
                entry[current] = len(local_order)
                local_order.append(current)
            edges = forest.get(current, ())
            if next_index < len(edges):
                neighbor, pipe_id = edges[next_index]
                stack[-1][2] = next_index + 1
                if neighbor == parent_id or neighbor in parent:
                    continue
                parent[neighbor] = current
                parent_bridge[neighbor] = pipe_id
                children[current].append(neighbor)
                stack.append([neighbor, current, 0])
                continue
            exit_index[current] = len(local_order)
            stack.pop()
        tree_components[root] = local_order

    tree_node_prefix: dict[int, list[int]] = {}
    for root, local_order in tree_components.items():
        prefix = [0]
        for component_id in local_order:
            prefix.append(prefix[-1] + len(component_nodes[component_id]))
        tree_node_prefix[root] = prefix


    subtree_summary: dict[int, _DischargeSummary] = {}
    subtree_source_count: dict[int, int] = {}
    outside_summary: dict[int, _DischargeSummary] = {}
    tree_total_sources: dict[int, int] = {}
    for root, local_order in tree_components.items():
        for component_id in reversed(local_order):
            summary = base_summary[component_id]
            source_count = base_source_count[component_id]
            for child in children[component_id]:
                summary = _combine_discharge(summary, subtree_summary[child])
                source_count += subtree_source_count[child]
            subtree_summary[component_id] = summary
            subtree_source_count[component_id] = source_count
        tree_total_sources[root] = subtree_source_count[root]
        outside_summary[root] = empty

        for component_id in local_order:
            child_ids = children[component_id]
            contributions = [
                outside_summary[component_id],
                base_summary[component_id],
                *(subtree_summary[child] for child in child_ids),
            ]
            prefix = [empty]
            for contribution in contributions:
                prefix.append(_combine_discharge(prefix[-1], contribution))
            suffix = [empty] * (len(contributions) + 1)
            for index in range(len(contributions) - 1, -1, -1):
                suffix[index] = _combine_discharge(
                    contributions[index], suffix[index + 1]
                )
            for child_index, child in enumerate(child_ids, start=2):
                outside_summary[child] = _combine_discharge(
                    prefix[child_index], suffix[child_index + 1]
                )

    bridge_child = {pipe_id: child for child, pipe_id in parent_bridge.items()}
    all_node_ids = tuple(sorted(nodes))
    total_summary = empty
    for root in tree_components:
        total_summary = _combine_discharge(total_summary, subtree_summary[root])
    total_min_flow = total_summary.flow_lpm

    def side_node_ids(child: int, use_subtree: bool) -> Sequence[str]:
        root = tree_root[child]
        ordered_components = tree_components[root]
        start_index = entry[child]
        stop_index = exit_index[child]
        prefix = tree_node_prefix[root]
        subtree_size = prefix[stop_index] - prefix[start_index]
        selected_size = (
            subtree_size if use_subtree else prefix[-1] - subtree_size
        )
        return _LazyDownstreamNodeIds(
            component_nodes=component_nodes,
            ordered_components=ordered_components,
            start=start_index,
            stop=stop_index,
            use_subtree=use_subtree,
            size=selected_size,
        )

    contexts: dict[str, PipeSizingContext] = {}
    pending: dict[str, dict[str, Any]] = {}

    for pipe_id, pipe in pipes.items():
        pid = str(pipe_id)
        start = str(_get(pipe, "start", "") or "")
        end = str(_get(pipe, "end", "") or "")
        if start not in nodes or end not in nodes or start == end:
            issues.append(f"pipe_endpoint_invalid:{pid}")
            continue
        is_cycle = pid not in bridge_ids
        is_supply = pid in supply_pipe_ids

        if is_cycle:
            downstream_summary = total_summary
            downstream_node_ids = all_node_ids
        else:
            child = bridge_child[pid]
            start_component = component_of[start]
            end_component = component_of[end]
            start_is_subtree = start_component == child
            child_sources = subtree_source_count[child]
            total_sources = tree_total_sources[tree_root[child]]
            start_sources = child_sources if start_is_subtree else total_sources - child_sources
            end_sources = total_sources - start_sources
            if start_sources > 0 and end_sources <= 0:
                # Sources only on the start side → demand is toward end.
                use_subtree = not start_is_subtree
            elif end_sources > 0 and start_sources <= 0:
                use_subtree = start_is_subtree
            else:
                # Both sides have sources (or neither). Prefer the farther
                # endpoint as upstream. Equal distance = supply/tank cross-link
                # (e.g. WT↔pump diamond bridge): distance alone is a coin flip
                # and can point at an empty pocket (0 heads) while the other
                # side carries the system. Pick the heavier discharge side.
                d_start = distance.get(start, 10**9)
                d_end = distance.get(end, 10**9)
                if d_start != d_end:
                    use_start_side = d_start > d_end
                    use_subtree = (
                        start_is_subtree if use_start_side else not start_is_subtree
                    )
                else:
                    sub = subtree_summary[child]
                    out = outside_summary[child]
                    use_subtree = (int(sub.count), float(sub.flow_lpm)) >= (
                        int(out.count),
                        float(out.flow_lpm),
                    )
            downstream_summary = (
                subtree_summary[child] if use_subtree else outside_summary[child]
            )
            downstream_node_ids = side_node_ids(child, use_subtree)

        discharge_types = {NodeTypeId.HEAD.value, NodeTypeId.NOZZLE.value}
        pending[pid] = {
            "basis_mixed": bool(downstream_summary.mixed),
            "downstream_head_count": int(downstream_summary.count),
            "active_head_count": int(downstream_summary.active_count),
            "downstream_hose_count": int(downstream_summary.hose_count),
            "active_hose_count": int(downstream_summary.active_hose_count),
            "downstream_node_ids": downstream_node_ids,
            "estimated_flow_lpm": float(downstream_summary.flow_lpm),
            "has_discharge_endpoint": any(
                str(_get(nodes[node_id], "type_id", "") or "").strip().lower()
                in discharge_types
                for node_id in (start, end)
            ),
            "touches_short_drop": (
                start in short_drop_junctions or end in short_drop_junctions
            ),
            "is_cycle": is_cycle,
            "is_supply": is_supply,
        }

    for pid, fact in pending.items():
        is_supply = bool(fact["is_supply"])
        is_cycle = bool(fact["is_cycle"])
        has_discharge_endpoint = bool(fact["has_discharge_endpoint"])
        head_count = int(fact["downstream_head_count"])
        estimated = float(fact["estimated_flow_lpm"])
        if is_supply:
            pipe_class = PipeClass.SUPPLY
            basis = SizingBasis.FLOW
            estimated = max(estimated, total_min_flow)
        else:
            # 펌프 토출측이 아니면 모두 가지. 링은 클래스만 BRANCH, 기준은 FLOW.
            pipe_class = PipeClass.BRANCH
            if is_cycle:
                basis = SizingBasis.FLOW
                estimated = max(estimated, total_min_flow)
            else:
                basis = (
                    SizingBasis.HEAD_COUNT
                    if head_count > 0 and not fact["basis_mixed"]
                    else SizingBasis.FLOW
                )

        contexts[pid] = PipeSizingContext(
            pipe_id=pid,
            basis=basis,
            pipe_class=pipe_class,
            downstream_head_count=head_count,
            estimated_flow_lpm=max(float(estimated), 0.0),
            is_cycle=is_cycle,
            is_supply_path=is_supply,
            has_discharge_endpoint=has_discharge_endpoint,
            touches_short_drop=bool(fact["touches_short_drop"]),
            downstream_node_ids=fact["downstream_node_ids"],
            downstream_hose_count=int(fact["downstream_hose_count"]),
            downstream_active_head_count=int(fact["active_head_count"]),
            downstream_active_hose_count=int(fact["active_hose_count"]),
        )
    return contexts, list(dict.fromkeys(issues))


# Fixed empty slots in the unified sizing dialog (cumulative "until N heads").
# Each row is bound to one DN on this ladder; users edit head counts only.
HEAD_RANGE_FIXED_NOMINALS: tuple[int, ...] = (
    25,
    32,
    40,
    50,
    65,
    80,
    100,
    125,
    150,
    200,
)
HEAD_RANGE_SLOT_COUNT = len(HEAD_RANGE_FIXED_NOMINALS)
# Dialog DN combo upper bound (library may offer larger; UI caps at DN300).
MAX_DIALOG_NOMINAL_MM = 300
# UI stand-in for open-ended last head range (RangeRow.end is None).
OPEN_ENDED_HEAD_UNTIL_UI = 9999


def head_slots_to_ranges(
    slots: Sequence[tuple[Any, Any]],
) -> list[RangeRow]:
    """Convert cumulative (until_count, nominal_mm) slots to contiguous ranges.

    Empty slots (missing until and/or DN) are skipped. Rows are sorted by
    until-count; each range covers (prev_until+1) .. until inclusive, with the
    first range starting at 1.
    """
    parsed: list[tuple[int, int]] = []
    for raw_until, raw_dn in slots:
        if raw_until in (None, "") or raw_dn in (None, ""):
            continue
        try:
            until = int(round(float(raw_until)))
            nominal = int(round(float(raw_dn)))
        except (TypeError, ValueError, OverflowError):
            continue
        if until < 1 or nominal <= 0:
            continue
        parsed.append((until, nominal))
    if not parsed:
        return []
    parsed.sort(key=lambda item: item[0])
    rows: list[RangeRow] = []
    prev_end = 0
    for until, nominal in parsed:
        start = prev_end + 1
        if until < start:
            # Overlapping/non-increasing "until" — keep last DN for that end.
            if rows:
                rows[-1] = RangeRow(rows[-1].start, float(until), nominal)
            else:
                rows.append(RangeRow(1.0, float(until), nominal))
            prev_end = until
            continue
        rows.append(RangeRow(float(start), float(until), nominal))
        prev_end = until
    return rows


def ranges_to_head_slots(
    rows: Sequence[RangeRow],
    *,
    slot_count: int = HEAD_RANGE_SLOT_COUNT,
) -> list[tuple[int | None, int | None]]:
    """Convert contiguous head ranges to cumulative until-count slots.

    Always keeps every populated range. ``slot_count`` is only a *minimum*
    number of rows (empty trailing slots for manual editing); it must not
    truncate auto-sizing results that need more breakpoints.
    """
    slots: list[tuple[int | None, int | None]] = []
    coerced: list[RangeRow] = []
    for raw in rows:
        row, errors = _coerce_range_row(raw)
        if errors or not math.isfinite(row.start) or int(row.nominal_mm) <= 0:
            continue
        coerced.append(row)
    for index, row in enumerate(coerced):
        if row.end is None:
            # Only the trailing open range maps to the UI sentinel.
            if index != len(coerced) - 1:
                continue
            until = OPEN_ENDED_HEAD_UNTIL_UI
        elif not math.isfinite(row.end):
            continue
        else:
            until = int(round(float(row.end)))
            if until < 1:
                continue
        slots.append((until, int(row.nominal_mm)))
    minimum = max(int(slot_count), 0)
    while len(slots) < minimum:
        slots.append((None, None))
    return slots


def ranges_to_fixed_head_slots(
    rows: Sequence[RangeRow],
    *,
    nominals: Sequence[int] = HEAD_RANGE_FIXED_NOMINALS,
) -> list[tuple[int | None, int]]:
    """Map head ranges onto the fixed DN ladder (empty until = unused DN)."""
    until_by_dn: dict[int, int] = {}
    coerced: list[RangeRow] = []
    for raw in rows:
        row, errors = _coerce_range_row(raw)
        if errors or not math.isfinite(row.start) or int(row.nominal_mm) <= 0:
            continue
        coerced.append(row)
    for index, row in enumerate(coerced):
        if row.end is None:
            if index != len(coerced) - 1:
                continue
            until = OPEN_ENDED_HEAD_UNTIL_UI
        elif not math.isfinite(row.end):
            continue
        else:
            until = int(round(float(row.end)))
            if until < 1:
                continue
        until_by_dn[int(row.nominal_mm)] = until
    return [
        (until_by_dn.get(int(nominal)), int(nominal)) for nominal in nominals
    ]


def distinct_proposal_nominals(
    proposals: Mapping[str, PipeSizingProposal] | Mapping[str, Any],
    pipe_class: PipeClass,
) -> list[int]:
    """Sorted unique proposed DN values for a pipe class."""
    values: set[int] = set()
    for proposal in proposals.values():
        raw_class = _get(proposal, "pipe_class", None)
        if isinstance(raw_class, PipeClass):
            matched = raw_class is pipe_class
        else:
            matched = str(raw_class or "").strip().lower() == pipe_class.value
        if not matched:
            continue
        try:
            nominal = int(_get(proposal, "proposed_nominal_mm", 0) or 0)
        except (TypeError, ValueError, OverflowError):
            continue
        if nominal > 0:
            values.add(nominal)
    return sorted(values)


def build_head_ranges(
    contexts: Mapping[str, PipeSizingContext],
    proposals: Mapping[str, PipeSizingProposal],
) -> list[RangeRow]:
    by_count: dict[int, int] = {}
    for pipe_id, context in contexts.items():
        if context.basis is not SizingBasis.HEAD_COUNT or context.downstream_head_count <= 0:
            continue
        proposal = proposals.get(pipe_id)
        if proposal is None:
            continue
        count = int(context.downstream_head_count)
        by_count[count] = max(by_count.get(count, 0), int(proposal.proposed_nominal_mm))
    if not by_count:
        return []

    max_count = max(by_count)
    known = sorted(by_count)
    dense: dict[int, int] = {}
    running = by_count[known[0]]
    cursor = 0
    for count in range(1, max_count + 1):
        while cursor < len(known) and known[cursor] <= count:
            running = max(running, by_count[known[cursor]])
            cursor += 1
        dense[count] = running

    rows: list[RangeRow] = []
    start = 1
    current_dn = dense[1]
    for count in range(2, max_count + 1):
        if dense[count] != current_dn:
            rows.append(RangeRow(float(start), float(count - 1), current_dn))
            start = count
            current_dn = dense[count]
    rows.append(RangeRow(float(start), float(max_count), current_dn))
    return rows


def build_common_flow_ranges(
    specs: Sequence[PipeSizeSpec],
    limits: Iterable[float],
) -> list[RangeRow]:
    """Build one conservative common table from real-ID capacity boundaries."""
    ordered = sorted(
        (spec for spec in specs if isinstance(spec, PipeSizeSpec) and not spec.validate()),
        key=lambda item: (int(item.nominal_mm), float(item.inner_d_mm)),
    )
    clean_limit_values: set[float] = set()
    for value in limits:
        try:
            numeric = float(value)
        except (TypeError, ValueError, OverflowError):
            continue
        if math.isfinite(numeric) and numeric > 0.0:
            clean_limit_values.add(numeric)
    clean_limits = sorted(clean_limit_values)
    if not ordered or not clean_limits:
        return []
    boundaries = sorted({
        flow_capacity_lpm(spec.inner_d_mm, limit)
        for spec in ordered[:-1]
        for limit in clean_limits
    })
    rows: list[RangeRow] = []
    start = 0.0
    for boundary in boundaries:
        required_dn = 0
        for limit in clean_limits:
            selected, _overflow = select_spec_for_flow(ordered, boundary, limit)
            if selected is not None:
                required_dn = max(required_dn, selected.nominal_mm)
        if rows and rows[-1].nominal_mm == required_dn:
            rows[-1] = RangeRow(rows[-1].start, boundary, required_dn)
        else:
            rows.append(RangeRow(start, boundary, required_dn))
        start = boundary
    max_dn = max(item.nominal_mm for item in ordered)
    if rows and rows[-1].nominal_mm == max_dn:
        rows[-1] = RangeRow(rows[-1].start, None, max_dn)
    else:
        rows.append(RangeRow(start, None, max_dn))
    return rows


def model_signature(graph: Any, design_settings: Mapping[str, Any] | None = None) -> str:
    """Hash every persisted input that can alter the hydraulic model.

    Runtime result fields are deliberately excluded. Valve state/direction,
    tank level, PRV/orifice settings and pump curves are deliberately included
    so a preview cannot be applied after any of those inputs changes.
    """
    nodes, pipes, _adj = _graph_parts(graph)

    def stable_value(value: Any) -> Any:
        if value is None or isinstance(value, (str, bool, int)):
            return value
        if isinstance(value, float):
            if math.isnan(value):
                return {"__float__": "nan"}
            if math.isinf(value):
                return {"__float__": "inf" if value > 0 else "-inf"}
            return round(value, 12)
        if isinstance(value, Enum):
            return stable_value(value.value)
        if isinstance(value, Mapping):
            return {
                str(key): stable_value(item)
                for key, item in sorted(value.items(), key=lambda pair: str(pair[0]))
            }
        if isinstance(value, (list, tuple)):
            return [stable_value(item) for item in value]
        if isinstance(value, (set, frozenset)):
            normalized = [stable_value(item) for item in value]
            return sorted(
                normalized,
                key=lambda item: json.dumps(item, ensure_ascii=True, sort_keys=True),
            )
        try:
            numeric = float(value)
        except (TypeError, ValueError, OverflowError):
            return {"__type__": type(value).__qualname__, "value": str(value)}
        return stable_value(numeric)

    node_fields = (
        "type",
        "type_id",
        "category_id",
        "coords",
        "elevation_m",
        "is_active",
        "is_closed",
        "k_factor_si",
        "required_pressure_bar",
        "base_demand_lps",
        "fitting_id",
        "head_spec_name",
        "nozzle_name",
        "pump_library_id",
        "rated_q",
        "rated_p",
        "shutoff_p",
        "peak_q",
        "peak_p",
        "pump_curve_data",
        "water_level",
        "pressure_setting_bar",
        "loss_coefficient",
        "hole_diameter",
        "orifice_discharge_coeff",
        "is_check_valve",
        "hose_stream_flow_lps",
        "pipe_dn",
        "check_valve_direction",
        "has_check_valve",
    )
    pipe_fields = (
        "start",
        "end",
        "type",
        "diameter",
        "nominal_mm",
        "length_m",
        "equivalent_length",
        "C",
        "roughness_mm",
        "fittings",
    )

    payload = {
        "nodes": [
            {
                "id": str(node_id),
                "input": {
                    field_name: stable_value(_get(node, field_name, None))
                    for field_name in node_fields
                },
            }
            for node_id, node in sorted(nodes.items(), key=lambda item: str(item[0]))
        ],
        "pipes": [
            {
                "id": str(pipe_id),
                "input": {
                    field_name: stable_value(_get(pipe, field_name, None))
                    for field_name in pipe_fields
                },
            }
            for pipe_id, pipe in sorted(pipes.items(), key=lambda item: str(item[0]))
        ],
        "design_settings": stable_value(
            {
                str(key): value
                for key, value in design_settings.items()
                if str(key) != "pipe_sizing"
            }
            if isinstance(design_settings, Mapping) else {}
        ),
    }
    raw = json.dumps(payload, ensure_ascii=True, separators=(",", ":"), sort_keys=True)
    return hashlib.sha256(raw.encode("utf-8")).hexdigest()


__all__ = [
    "ChangeStatus",
    "DEFAULT_BRANCH_VELOCITY_MPS",
    "DEFAULT_MAIN_VELOCITY_MPS",
    "DEFAULT_SUPPLY_VELOCITY_MPS",
    "FLOW_EPS_LPM",
    "FLOW_TOLERANCE_LPM",
    "HEAD_RANGE_FIXED_NOMINALS",
    "HEAD_RANGE_SLOT_COUNT",
    "MAX_DIALOG_NOMINAL_MM",
    "OPEN_ENDED_HEAD_UNTIL_UI",
    "PipeClass",
    "PipeSizeSpec",
    "PipeSizingContext",
    "PipeSizingPreview",
    "PipeSizingProposal",
    "RangeRow",
    "RangeValidation",
    "SizingBasis",
    "VelocityLimits",
    "analyze_pipe_contexts",
    "available_standards",
    "build_common_flow_ranges",
    "build_flow_ranges",
    "build_head_ranges",
    "design_discharge_lpm",
    "distinct_proposal_nominals",
    "find_bridge_pipe_ids",
    "flow_capacity_lpm",
    "head_slots_to_ranges",
    "hose_stream_design_flow_lpm",
    "hose_stream_sizing_flow_lpm",
    "lookup_head_range",
    "lookup_range",
    "MIN_BRANCH_NOMINAL_MM",
    "MIN_DISCHARGE_PATH_NOMINAL_MM",
    "MIN_HOSE_PATH_NOMINAL_MM",
    "MIN_MAIN_NOMINAL_MM",
    "MIN_SUPPLY_NOMINAL_MM",
    "minimum_discharge_lpm",
    "minimum_nominal_for_context",
    "sizing_flow_producers_missing",
    "model_signature",
    "pipe_specs_for_standard",
    "ranges_to_fixed_head_slots",
    "ranges_to_head_slots",
    "resolve_pipe_standard",
    "select_spec_for_flow",
    "spec_for_nominal",
    "validate_ranges",
    "velocity_mps",
]
