"""Validated hydraulic inputs; metres, millimetres, L/min and gauge bar."""
from __future__ import annotations

from dataclasses import dataclass
from math import isfinite

BAR_PER_M = 998.2 * 9.80665 / 100000.0  # water at 20 C


@dataclass(frozen=True)
class Node:
    """Physical elevation, never an isometric display coordinate."""
    label: str
    elevation_m: float


@dataclass(frozen=True)
class Size:
    """One available library size, including its diameter-dependent losses."""
    nominal_mm: float
    inner_mm: float
    equivalent_m: float = 0.0


@dataclass(frozen=True)
class Pipe:
    """An undirected physical pipe; signed flows follow a -> b."""
    label: str
    a: str
    b: str
    length_m: float
    c_factor: float
    sizes: tuple[Size, ...]
    initial: int = 0
    locked: bool = False
    role: str = "unknown"


@dataclass(frozen=True)
class Nozzle:
    """An active atmospheric sprinkler: Q[L/min] = K sqrt(P[bar])."""
    label: str
    node: str
    k_lpm_sqrt_bar: float
    min_flow_lpm: float
    min_pressure_bar: float
    max_pressure_bar: float


@dataclass(frozen=True)
class Network:
    """Single prescribed supply boundary and connected physical network."""
    nodes: tuple[Node, ...]
    pipes: tuple[Pipe, ...]
    nozzles: tuple[Nozzle, ...]
    source: str

    def validate(self) -> None:
        """Reject missing physics rather than replacing them with defaults."""
        nodes = {n.label: n for n in self.nodes}
        if len(nodes) != len(self.nodes) or self.source not in nodes:
            raise ValueError("절점 라벨 중복 또는 공급 절점 누락")
        if not self.pipes or not self.nozzles:
            raise ValueError("연결 배관과 작동 헤드가 필요합니다.")
        for group in (self.pipes, self.nozzles):
            if len({r.label for r in group}) != len(group):
                raise ValueError("배관/노즐 라벨 중복")
        adjacent = {n: set() for n in nodes}
        for n in self.nodes:
            if not isfinite(n.elevation_m):
                raise ValueError(f"절점 {n.label}: 표고가 유효하지 않습니다.")
        for p in self.pipes:
            if p.a not in nodes or p.b not in nodes or p.a == p.b:
                raise ValueError(f"배관 {p.label}: 연결 절점 오류")
            if not isfinite(p.length_m) or p.length_m <= 0:
                raise ValueError(f"배관 {p.label}: 실제 길이는 0보다 커야 합니다.")
            if p.length_m + .002 < abs(nodes[p.a].elevation_m - nodes[p.b].elevation_m):
                raise ValueError(f"배관 {p.label}: 길이가 표고차보다 작습니다.")
            if not isfinite(p.c_factor) or p.c_factor <= 0:
                raise ValueError(f"배관 {p.label}: C 값이 필요합니다.")
            if p.role not in {"branch", "other", "unknown"}:
                raise ValueError(f"배관 {p.label}: 배관 역할 오류")
            if not p.sizes or not 0 <= p.initial < len(p.sizes):
                raise ValueError(f"배관 {p.label}: 사용 가능한 관경이 없습니다.")
            for size in p.sizes:
                if (not all(isfinite(v) for v in (size.nominal_mm, size.inner_mm, size.equivalent_m))
                        or min(size.nominal_mm, size.inner_mm) <= 0 or size.equivalent_m < 0):
                    raise ValueError(f"배관 {p.label}: 내경/등가길이 오류")
            if any(a.nominal_mm >= b.nominal_mm or a.inner_mm >= b.inner_mm
                   for a, b in zip(p.sizes, p.sizes[1:])):
                raise ValueError(f"배관 {p.label}: 관경 후보는 증가 순이어야 합니다.")
            adjacent[p.a].add(p.b)
            adjacent[p.b].add(p.a)
        seen, todo = set(), [self.source]
        while todo:
            n = todo.pop()
            if n not in seen:
                seen.add(n)
                todo.extend(adjacent[n] - seen)
        if seen != set(nodes):
            raise ValueError("공급원과 분리된 절점: " + ", ".join(sorted(set(nodes) - seen)[:12]))
        for h in self.nozzles:
            values = (h.k_lpm_sqrt_bar, h.min_flow_lpm, h.min_pressure_bar, h.max_pressure_bar)
            if h.node not in nodes or not all(isfinite(v) and v > 0 for v in values):
                raise ValueError(f"헤드 {h.label}: K/요구 유량/압력 정보 오류")
            if h.max_pressure_bar < h.min_pressure_bar:
                raise ValueError(f"헤드 {h.label}: 최대압력이 최소압력보다 작습니다.")


@dataclass(frozen=True)
class Settings:
    """User design constraints, not a regulatory compliance declaration."""
    mode: str = "joint"
    supply_pressure_bar: float = 8.0
    max_supply_pressure_bar: float = 20.0
    suction_total_head_m: float = 0.0
    discharge_velocity_head_m: float = 0.0
    branch_velocity_mps: float = 6.0
    other_velocity_mps: float = 10.0
    max_iterations: int = 40
    duration_minutes: float | None = None

    def validate(self) -> None:
        """Bound the search and reject non-finite user input."""
        if self.mode not in {"fixed_supply", "pump_only", "joint"}:
            raise ValueError("역산 방식 오류")
        for v in (self.supply_pressure_bar, self.max_supply_pressure_bar,
                  self.branch_velocity_mps, self.other_velocity_mps):
            if not isfinite(v) or v <= 0:
                raise ValueError("압력/유속 제한은 유한한 양수여야 합니다.")
        if (not isfinite(self.suction_total_head_m)
                or not isfinite(self.discharge_velocity_head_m) or self.discharge_velocity_head_m < 0
                or type(self.max_iterations) is not int or not 1 <= self.max_iterations <= 100):
            raise ValueError("흡입 전수두 또는 반복 제한 오류")
        if self.duration_minutes is not None and (
            isinstance(self.duration_minutes, bool)
            or not isinstance(self.duration_minutes, (int, float))
            or not isfinite(self.duration_minutes) or self.duration_minutes <= 0
        ):
            raise ValueError("방수 지속시간은 유한한 양수(분)여야 합니다.")
