"""Tests for network CSV table exports."""

from pathlib import Path

import pandas as pd

from pipenet_converter.export_tables import (
    EQUIPMENT_COLUMNS,
    FITTING_COLUMNS,
    NODE_COLUMNS,
    NOZZLE_COLUMNS,
    PIPE_COLUMNS,
    VALVE_COLUMNS,
    write_network_tables,
)
from pipenet_converter.models import Equipment, Fitting, Node, Nozzle, Pipe, PipeNetwork, Valve


def test_write_network_tables_creates_all_csv_files(tmp_path: Path) -> None:
    output_dir = tmp_path / "tables"

    write_network_tables(_sample_network(), output_dir)

    expected_files = {
        "network_3d_nodes.csv",
        "network_3d_pipes.csv",
        "network_3d_nozzles.csv",
        "network_3d_fittings.csv",
        "network_3d_equipment.csv",
        "network_3d_valves.csv",
    }
    assert {path.name for path in output_dir.iterdir()} == expected_files


def test_exported_csv_files_have_expected_columns_and_row_counts(tmp_path: Path) -> None:
    output_dir = tmp_path / "tables"
    network = _sample_network()

    write_network_tables(network, output_dir)

    nodes = pd.read_csv(output_dir / "network_3d_nodes.csv")
    pipes = pd.read_csv(output_dir / "network_3d_pipes.csv")
    nozzles = pd.read_csv(output_dir / "network_3d_nozzles.csv")
    fittings = pd.read_csv(output_dir / "network_3d_fittings.csv")
    equipment = pd.read_csv(output_dir / "network_3d_equipment.csv")
    valves = pd.read_csv(output_dir / "network_3d_valves.csv")

    assert list(nodes.columns) == NODE_COLUMNS
    assert list(pipes.columns) == PIPE_COLUMNS
    assert list(nozzles.columns) == NOZZLE_COLUMNS
    assert list(fittings.columns) == FITTING_COLUMNS
    assert list(equipment.columns) == EQUIPMENT_COLUMNS
    assert list(valves.columns) == VALVE_COLUMNS

    assert len(nodes) == len(network.nodes)
    assert len(pipes) == len(network.pipes)
    assert len(nozzles) == len(network.nozzles)
    assert len(fittings) == sum(len(pipe.fittings) for pipe in network.pipes.values())
    assert len(equipment) == sum(len(pipe.equipment) for pipe in network.pipes.values())
    assert len(valves) == len(network.valves)
    assert pipes.loc[0, "diameter_label"] == "150A"


def _sample_network() -> PipeNetwork:
    network = PipeNetwork(title="Export sample")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input", source="test"))
    network.add_node(Node("N001", 1.0, 0.0, 0.0, "main_pipe", source="test"))
    network.add_node(Node("@/1", 1.0, 0.0, 0.3, "head_output", source="test"))
    network.add_pipe(
        Pipe(
            "P001",
            "INPUT",
            "N001",
            0.15,
            1.0,
            0.0,
            material="KSD3507",
            fittings=[Fitting("gate", 1), Fitting("check", 1)],
            equipment=[Equipment("EQ1", "AV", 24.0, rel_position=0.5)],
        )
    )
    network.add_nozzle(Nozzle("NZ1", "N001", "@/1", 0.00266666667, status=1))
    network.add_valve(Valve("V1", "INPUT", "N001", "pressure-drop", target_value=0.0))
    return network
