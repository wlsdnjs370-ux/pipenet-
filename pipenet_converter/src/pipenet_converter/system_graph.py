"""System diagram and riser CSV integration."""

from __future__ import annotations

import csv
from copy import deepcopy
from math import dist
from pathlib import Path

from pipenet_converter.models import Equipment, Fitting, Node, Pipe, PipeNetwork


def parse_system_edges_csv(path: str | Path) -> PipeNetwork:
    """Parse a human-authored ``system_edges.csv`` into a system ``PipeNetwork``."""
    network = PipeNetwork(title="System riser network", metadata={"source": "system_edges_csv"})
    rows = _read_rows(path)

    for row_index, row in enumerate(rows):
        from_node_id = row["from_node"].strip()
        to_node_id = row["to_node"].strip()
        _ensure_system_node(network, from_node_id, float(row["from_z_m"]), row_index)
        _ensure_system_node(network, to_node_id, float(row["to_z_m"]), row_index + 1)

        network.add_pipe(
            Pipe(
                pipe_id=row["edge_id"].strip(),
                from_node=from_node_id,
                to_node=to_node_id,
                diameter_m=float(row["diameter_m"]),
                length_m=float(row["length_m"]),
                rise_m=float(row["rise_m"]),
                c_factor=float(row.get("c_factor") or 120.0),
                material=(row.get("material") or None),
                fittings=_parse_fittings(row.get("fittings", "")),
                equipment=_parse_equipment(row.get("equipment", ""), row["edge_id"].strip()),
                metadata={"description": row.get("description", ""), "source": "system_edges_csv"},
            )
        )

    return network


def merge_networks(
    system_network: PipeNetwork,
    plan_network: PipeNetwork,
    connection_map: dict[str, str],
) -> PipeNetwork:
    """Merge system/riser and plan networks and connect mapped nodes."""
    merged = PipeNetwork(
        title=f"{system_network.title} + {plan_network.title}",
        metadata={"source": "merged_system_plan"},
    )
    system_node_ids = _copy_system_nodes_with_conflict_prefix(system_network, plan_network, merged)
    _copy_plan_nodes(plan_network, merged)
    _copy_system_pipes(system_network, merged, system_node_ids)
    _copy_plan_pipes(plan_network, merged)
    _copy_plan_nozzles(plan_network, merged)
    _copy_plan_valves(plan_network, merged)

    for index, (system_node_id, plan_node_id) in enumerate(connection_map.items(), start=1):
        resolved_system_node_id = system_node_ids[system_node_id]
        if plan_node_id not in merged.nodes:
            raise KeyError(f"Plan connection node {plan_node_id!r} does not exist.")
        _add_connection_pipe(merged, resolved_system_node_id, plan_node_id, index)

    return merged


def _read_rows(path: str | Path) -> list[dict[str, str]]:
    csv_path = Path(path)
    if not csv_path.exists():
        raise FileNotFoundError(f"System edges CSV file not found: {csv_path}")
    with csv_path.open(newline="", encoding="utf-8-sig") as csv_file:
        reader = csv.DictReader(csv_file)
        _require_columns(
            reader.fieldnames,
            {
                "edge_id",
                "from_node",
                "to_node",
                "from_z_m",
                "to_z_m",
                "diameter_m",
                "length_m",
                "rise_m",
                "c_factor",
                "material",
                "fittings",
                "equipment",
                "description",
            },
            csv_path,
        )
        rows = list(reader)
    if not rows:
        raise ValueError(f"System edges CSV file {csv_path} does not contain any rows.")
    return rows


def _ensure_system_node(network: PipeNetwork, node_id: str, z_m: float, order: int) -> None:
    existing = network.nodes.get(node_id)
    if existing is not None:
        existing.z = z_m
        return
    network.add_node(
        Node(
            node_id=node_id,
            x=0.0,
            y=float(order),
            z=z_m,
            node_type="input" if node_id.upper() == "INPUT" else "riser",
            source="system_edges_csv",
        )
    )


def _parse_fittings(value: str) -> list[Fitting]:
    fittings: list[Fitting] = []
    for item in _split_semicolon_items(value):
        fitting_type, count = item.split(":", 1)
        fittings.append(Fitting(fitting_type.strip(), int(count)))
    return fittings


def _parse_equipment(value: str, pipe_id: str) -> list[Equipment]:
    equipment_items: list[Equipment] = []
    for index, item in enumerate(_split_semicolon_items(value), start=1):
        description, equivalent_length = item.split(":", 1)
        equipment_items.append(
            Equipment(
                equipment_id=f"{pipe_id}_EQ{index}",
                description=description.strip(),
                equivalent_length_m=float(equivalent_length),
            )
        )
    return equipment_items


def _split_semicolon_items(value: str) -> list[str]:
    return [item.strip() for item in value.split(";") if item.strip()]


def _copy_system_nodes_with_conflict_prefix(
    system_network: PipeNetwork,
    plan_network: PipeNetwork,
    merged: PipeNetwork,
) -> dict[str, str]:
    node_ids: dict[str, str] = {}
    for node_id, node in system_network.nodes.items():
        merged_node_id = f"SYS_{node_id}" if node_id in plan_network.nodes else node_id
        node_ids[node_id] = merged_node_id
        copied = deepcopy(node)
        copied.node_id = merged_node_id
        merged.add_node(copied)
    return node_ids


def _copy_plan_nodes(plan_network: PipeNetwork, merged: PipeNetwork) -> None:
    for node in plan_network.nodes.values():
        if node.node_id in merged.nodes:
            raise ValueError(f"Plan node {node.node_id!r} still conflicts after system prefixing.")
        merged.add_node(deepcopy(node))


def _copy_system_pipes(
    system_network: PipeNetwork,
    merged: PipeNetwork,
    system_node_ids: dict[str, str],
) -> None:
    for pipe in system_network.pipes.values():
        copied = deepcopy(pipe)
        copied.from_node = system_node_ids[pipe.from_node]
        copied.to_node = system_node_ids[pipe.to_node]
        merged.add_pipe(copied)


def _copy_plan_pipes(plan_network: PipeNetwork, merged: PipeNetwork) -> None:
    for pipe in plan_network.pipes.values():
        pipe_id = pipe.pipe_id
        copied = deepcopy(pipe)
        if pipe_id in merged.pipes:
            copied.pipe_id = f"PLAN_{pipe_id}"
        merged.add_pipe(copied)


def _copy_plan_nozzles(plan_network: PipeNetwork, merged: PipeNetwork) -> None:
    for nozzle in plan_network.nozzles.values():
        merged.add_nozzle(deepcopy(nozzle))


def _copy_plan_valves(plan_network: PipeNetwork, merged: PipeNetwork) -> None:
    for valve in plan_network.valves.values():
        merged.add_valve(deepcopy(valve))


def _add_connection_pipe(
    network: PipeNetwork,
    system_node_id: str,
    plan_node_id: str,
    index: int,
) -> None:
    system_node = network.nodes[system_node_id]
    plan_node = network.nodes[plan_node_id]
    horizontal = dist((system_node.x, system_node.y), (plan_node.x, plan_node.y))
    rise = plan_node.z - system_node.z
    length = max(horizontal + abs(rise), 0.001)
    network.add_pipe(
        Pipe(
            pipe_id=f"CONN_{index:03d}",
            from_node=system_node_id,
            to_node=plan_node_id,
            diameter_m=_nearest_system_diameter(network, system_node_id),
            length_m=length,
            rise_m=rise,
            c_factor=120.0,
            material=None,
            metadata={"source": "connection_map"},
        )
    )


def _nearest_system_diameter(network: PipeNetwork, node_id: str) -> float:
    for pipe in network.pipes.values():
        if pipe.from_node == node_id or pipe.to_node == node_id:
            return pipe.diameter_m
    return 0.15


def _require_columns(fieldnames: list[str] | None, required: set[str], path: Path) -> None:
    available = set(fieldnames or [])
    missing = sorted(required - available)
    if missing:
        raise ValueError(
            f"CSV file {path} is missing required columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(sorted(available))}"
        )
