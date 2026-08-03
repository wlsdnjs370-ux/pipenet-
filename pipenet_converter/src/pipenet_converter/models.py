"""Core domain models for PipeNet conversion.

The package uses these models as the intermediate representation between CAD
extraction, graph construction, validation, rendering, and SDF XML writing.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import TypeAlias

import pandas as pd

MetadataValue: TypeAlias = str | float | int | bool
Metadata: TypeAlias = dict[str, MetadataValue]

# 값은 호칭경(m) — PIPENET 이 SDF 의 bore 를 SLF <Size-definition nominal> 과
# 매칭해 실내경을 결정하므로 여기에 실내경을 넣으면 안 된다.
# 15A/20A 는 신축배관(FX, 20A)·말단 가지관에 실제로 쓰이는데 빠져 있어
# diameter_label_to_m 이 ValueError 를 내던 구간이다.
SUPPORTED_DIAMETERS_M: dict[str, float] = {
    "15A": 0.015,
    "20A": 0.02,
    "25A": 0.025,
    "32A": 0.032,
    "40A": 0.04,
    "50A": 0.05,
    "65A": 0.065,
    "80A": 0.08,
    "100A": 0.1,
    "125A": 0.125,
    "150A": 0.15,
    "200A": 0.2,
}


def diameter_label_to_m(label: str) -> float:
    """Convert a supported nominal diameter label such as ``150A`` to meters."""
    normalized = label.strip().upper().replace(" ", "")
    try:
        return SUPPORTED_DIAMETERS_M[normalized]
    except KeyError as exc:
        supported = ", ".join(SUPPORTED_DIAMETERS_M)
        raise ValueError(f"Unsupported diameter label {label!r}. Supported labels: {supported}") from exc


def diameter_m_to_label(diameter_m: float) -> str:
    """Convert a supported nominal diameter in meters to its label."""
    rounded = round(float(diameter_m), 6)
    for label, value in SUPPORTED_DIAMETERS_M.items():
        if round(value, 6) == rounded:
            return label
    supported = ", ".join(f"{value:g}m" for value in SUPPORTED_DIAMETERS_M.values())
    raise ValueError(f"Unsupported diameter {diameter_m!r}. Supported diameters: {supported}")


@dataclass(slots=True)
class Node:
    """A calculation node in the 3D sprinkler pipe network."""

    node_id: str
    x: float
    y: float
    z: float
    node_type: str
    source: str | None = None
    metadata: Metadata = field(default_factory=dict)


@dataclass(slots=True)
class Fitting:
    """A PipeNet fitting count associated with a pipe."""

    fitting_type: str
    count: int


@dataclass(slots=True)
class Equipment:
    """Equivalent-length equipment associated with a pipe."""

    equipment_id: str
    description: str
    equivalent_length_m: float
    rel_position: float = 0.5


@dataclass(slots=True)
class Pipe:
    """A directed pipe link between two nodes."""

    pipe_id: str
    from_node: str
    to_node: str
    diameter_m: float
    length_m: float
    rise_m: float
    c_factor: float = 120.0
    material: str | None = None
    status: str = "normal"
    fittings: list[Fitting] = field(default_factory=list)
    equipment: list[Equipment] = field(default_factory=list)
    waypoints: list[tuple[float, float]] = field(default_factory=list)
    metadata: Metadata = field(default_factory=dict)


@dataclass(slots=True)
class Valve:
    """A valve link between two nodes."""

    valve_id: str
    input_node: str
    output_node: str
    valve_type: str
    target_value: float | None = None
    metadata: Metadata = field(default_factory=dict)


@dataclass(slots=True)
class Nozzle:
    """A sprinkler nozzle connected to a pipe node and output node."""

    nozzle_id: str
    input_node: str
    output_node: str
    flow_m3s: float
    status: int = 1
    library_item: str = "SP-HEAD"
    metadata: Metadata = field(default_factory=dict)


@dataclass(slots=True)
class ValidationIssue:
    """A validation issue produced by rule checks."""

    severity: str
    code: str
    message: str
    object_type: str | None
    object_id: str | None


@dataclass(slots=True)
class PipeNetwork:
    """A calculation-ready sprinkler pipe network."""

    title: str
    nodes: dict[str, Node] = field(default_factory=dict)
    pipes: dict[str, Pipe] = field(default_factory=dict)
    nozzles: dict[str, Nozzle] = field(default_factory=dict)
    valves: dict[str, Valve] = field(default_factory=dict)
    metadata: Metadata = field(default_factory=dict)

    def add_node(self, node: Node) -> None:
        """Add a node to the network."""
        if node.node_id in self.nodes:
            raise ValueError(f"Duplicate node_id {node.node_id!r} in PipeNetwork {self.title!r}.")
        self.nodes[node.node_id] = node

    def add_pipe(self, pipe: Pipe) -> None:
        """Add a pipe to the network."""
        if pipe.pipe_id in self.pipes:
            raise ValueError(f"Duplicate pipe_id {pipe.pipe_id!r} in PipeNetwork {self.title!r}.")
        self.pipes[pipe.pipe_id] = pipe

    def add_nozzle(self, nozzle: Nozzle) -> None:
        """Add a nozzle to the network."""
        self.nozzles[nozzle.nozzle_id] = nozzle

    def add_valve(self, valve: Valve) -> None:
        """Add a valve to the network."""
        self.valves[valve.valve_id] = valve

    def get_node(self, node_id: str) -> Node:
        """Return a node by ID, raising ``KeyError`` when absent."""
        return self.nodes[node_id]

    def active_nozzles(self) -> list[Nozzle]:
        """Return nozzles with nonzero status."""
        return [nozzle for nozzle in self.nozzles.values() if nozzle.status != 0]

    def to_node_dataframe(self) -> pd.DataFrame:
        """Return network nodes as a pandas DataFrame."""
        return pd.DataFrame(
            [
                {
                    "node_id": node.node_id,
                    "x": node.x,
                    "y": node.y,
                    "z": node.z,
                    "node_type": node.node_type,
                    "source": node.source,
                    **node.metadata,
                }
                for node in self.nodes.values()
            ]
        )

    def to_pipe_dataframe(self) -> pd.DataFrame:
        """Return network pipes as a pandas DataFrame."""
        return pd.DataFrame(
            [
                {
                    "pipe_id": pipe.pipe_id,
                    "from_node": pipe.from_node,
                    "to_node": pipe.to_node,
                    "diameter_m": pipe.diameter_m,
                    "length_m": pipe.length_m,
                    "rise_m": pipe.rise_m,
                    "c_factor": pipe.c_factor,
                    "material": pipe.material,
                    "status": pipe.status,
                    "fitting_count": sum(fitting.count for fitting in pipe.fittings),
                    "equipment_count": len(pipe.equipment),
                    **pipe.metadata,
                }
                for pipe in self.pipes.values()
            ]
        )

    def to_nozzle_dataframe(self) -> pd.DataFrame:
        """Return network nozzles as a pandas DataFrame."""
        return pd.DataFrame(
            [
                {
                    "nozzle_id": nozzle.nozzle_id,
                    "input_node": nozzle.input_node,
                    "output_node": nozzle.output_node,
                    "flow_m3s": nozzle.flow_m3s,
                    "status": nozzle.status,
                    "library_item": nozzle.library_item,
                    **nozzle.metadata,
                }
                for nozzle in self.nozzles.values()
            ]
        )

    def to_fitting_dataframe(self) -> pd.DataFrame:
        """Return pipe fittings as a pandas DataFrame."""
        rows: list[dict[str, str | int]] = []
        for pipe in self.pipes.values():
            for fitting in pipe.fittings:
                rows.append(
                    {
                        "pipe_id": pipe.pipe_id,
                        "fitting_type": fitting.fitting_type,
                        "count": fitting.count,
                    }
                )
        return pd.DataFrame(rows)
