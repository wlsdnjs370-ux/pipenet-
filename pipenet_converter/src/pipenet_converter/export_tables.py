"""CSV table exports for intermediate PipeNetwork data."""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from pipenet_converter.models import (
    Equipment,
    Fitting,
    Node,
    Nozzle,
    Pipe,
    PipeNetwork,
    Valve,
    diameter_m_to_label,
)


NODE_COLUMNS = ["node_id", "x", "y", "z", "node_type", "source"]
PIPE_COLUMNS = [
    "pipe_id",
    "from_node",
    "to_node",
    "diameter_m",
    "diameter_label",
    "length_m",
    "rise_m",
    "c_factor",
    "material",
    "status",
]
NOZZLE_COLUMNS = ["nozzle_id", "input_node", "output_node", "flow_m3s", "status", "library_item"]
FITTING_COLUMNS = ["pipe_id", "fitting_type", "count"]
EQUIPMENT_COLUMNS = [
    "pipe_id",
    "equipment_id",
    "description",
    "equivalent_length_m",
    "rel_position",
]
VALVE_COLUMNS = ["valve_id", "input_node", "output_node", "valve_type", "target_value"]


def write_network_tables(network: PipeNetwork, output_dir: str | Path) -> None:
    """Write standard network CSV tables to ``output_dir``."""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    _nodes_dataframe(network).to_csv(output_path / "network_3d_nodes.csv", index=False)
    _pipes_dataframe(network).to_csv(output_path / "network_3d_pipes.csv", index=False)
    _nozzles_dataframe(network).to_csv(output_path / "network_3d_nozzles.csv", index=False)
    _fittings_dataframe(network).to_csv(output_path / "network_3d_fittings.csv", index=False)
    _equipment_dataframe(network).to_csv(output_path / "network_3d_equipment.csv", index=False)
    _valves_dataframe(network).to_csv(output_path / "network_3d_valves.csv", index=False)


def read_network_tables(input_dir: str | Path) -> PipeNetwork:
    """Read standard network CSV tables from ``input_dir``."""
    input_path = Path(input_dir)
    network = PipeNetwork(title=input_path.name or "network")

    for row in _read_required_csv(input_path / "network_3d_nodes.csv", NODE_COLUMNS):
        network.add_node(
            Node(
                node_id=row["node_id"],
                x=_float(row["x"]),
                y=_float(row["y"]),
                z=_float(row["z"]),
                node_type=row["node_type"],
                source=row.get("source") or None,
            )
        )

    for row in _read_required_csv(input_path / "network_3d_pipes.csv", PIPE_COLUMNS):
        network.add_pipe(
            Pipe(
                pipe_id=row["pipe_id"],
                from_node=row["from_node"],
                to_node=row["to_node"],
                diameter_m=_float(row["diameter_m"]),
                length_m=_float(row["length_m"]),
                rise_m=_float(row["rise_m"]),
                c_factor=_float(row["c_factor"]),
                material=row.get("material") or None,
                status=row.get("status") or "normal",
            )
        )

    for row in _read_optional_csv(input_path / "network_3d_fittings.csv", FITTING_COLUMNS):
        pipe = network.pipes.get(row["pipe_id"])
        if pipe is not None:
            pipe.fittings.append(Fitting(row["fitting_type"], int(row["count"])))

    for row in _read_optional_csv(input_path / "network_3d_equipment.csv", EQUIPMENT_COLUMNS):
        pipe = network.pipes.get(row["pipe_id"])
        if pipe is not None:
            pipe.equipment.append(
                Equipment(
                    equipment_id=row["equipment_id"],
                    description=row["description"],
                    equivalent_length_m=_float(row["equivalent_length_m"]),
                    rel_position=_float(row["rel_position"]),
                )
            )

    for row in _read_required_csv(input_path / "network_3d_nozzles.csv", NOZZLE_COLUMNS):
        network.add_nozzle(
            Nozzle(
                nozzle_id=row["nozzle_id"],
                input_node=row["input_node"],
                output_node=row["output_node"],
                flow_m3s=_float(row["flow_m3s"]),
                status=int(row["status"]),
                library_item=row["library_item"],
            )
        )

    for row in _read_optional_csv(input_path / "network_3d_valves.csv", VALVE_COLUMNS):
        target_value = row.get("target_value") or None
        network.add_valve(
            Valve(
                valve_id=row["valve_id"],
                input_node=row["input_node"],
                output_node=row["output_node"],
                valve_type=row["valve_type"],
                target_value=_float(target_value) if target_value is not None else None,
            )
        )

    return network


def _nodes_dataframe(network: PipeNetwork) -> pd.DataFrame:
    rows = [
        {
            "node_id": node.node_id,
            "x": node.x,
            "y": node.y,
            "z": node.z,
            "node_type": node.node_type,
            "source": node.source,
        }
        for node in network.nodes.values()
    ]
    return pd.DataFrame(rows, columns=NODE_COLUMNS)


def _pipes_dataframe(network: PipeNetwork) -> pd.DataFrame:
    rows = [
        {
            "pipe_id": pipe.pipe_id,
            "from_node": pipe.from_node,
            "to_node": pipe.to_node,
            "diameter_m": pipe.diameter_m,
            "diameter_label": _safe_diameter_label(pipe.diameter_m),
            "length_m": pipe.length_m,
            "rise_m": pipe.rise_m,
            "c_factor": pipe.c_factor,
            "material": pipe.material,
            "status": pipe.status,
        }
        for pipe in network.pipes.values()
    ]
    return pd.DataFrame(rows, columns=PIPE_COLUMNS)


def _nozzles_dataframe(network: PipeNetwork) -> pd.DataFrame:
    rows = [
        {
            "nozzle_id": nozzle.nozzle_id,
            "input_node": nozzle.input_node,
            "output_node": nozzle.output_node,
            "flow_m3s": nozzle.flow_m3s,
            "status": nozzle.status,
            "library_item": nozzle.library_item,
        }
        for nozzle in network.nozzles.values()
    ]
    return pd.DataFrame(rows, columns=NOZZLE_COLUMNS)


def _fittings_dataframe(network: PipeNetwork) -> pd.DataFrame:
    rows = [
        {
            "pipe_id": pipe.pipe_id,
            "fitting_type": fitting.fitting_type,
            "count": fitting.count,
        }
        for pipe in network.pipes.values()
        for fitting in pipe.fittings
    ]
    return pd.DataFrame(rows, columns=FITTING_COLUMNS)


def _equipment_dataframe(network: PipeNetwork) -> pd.DataFrame:
    rows = [
        {
            "pipe_id": pipe.pipe_id,
            "equipment_id": equipment.equipment_id,
            "description": equipment.description,
            "equivalent_length_m": equipment.equivalent_length_m,
            "rel_position": equipment.rel_position,
        }
        for pipe in network.pipes.values()
        for equipment in pipe.equipment
    ]
    return pd.DataFrame(rows, columns=EQUIPMENT_COLUMNS)


def _valves_dataframe(network: PipeNetwork) -> pd.DataFrame:
    rows = [
        {
            "valve_id": valve.valve_id,
            "input_node": valve.input_node,
            "output_node": valve.output_node,
            "valve_type": valve.valve_type,
            "target_value": valve.target_value,
        }
        for valve in network.valves.values()
    ]
    return pd.DataFrame(rows, columns=VALVE_COLUMNS)


def _safe_diameter_label(diameter_m: float) -> str:
    try:
        return diameter_m_to_label(diameter_m)
    except ValueError:
        return ""


def _read_csv(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    return pd.read_csv(path, dtype=str, keep_default_na=False).to_dict("records")


def _read_optional_csv(path: Path, required_columns: list[str]) -> list[dict[str, str]]:
    if not path.exists():
        return []
    return _read_required_csv(path, required_columns)


def _read_required_csv(path: Path, required_columns: list[str]) -> list[dict[str, str]]:
    if not path.exists():
        raise FileNotFoundError(f"Required network table not found: {path}")
    dataframe = pd.read_csv(path, dtype=str, keep_default_na=False)
    missing = [column for column in required_columns if column not in dataframe.columns]
    if missing:
        raise ValueError(
            f"CSV file {path} is missing required columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(dataframe.columns)}"
        )
    return dataframe.to_dict("records")


def _float(value: str) -> float:
    return float(value or 0.0)
