"""Tests for the Remote-N most-remote head extraction module."""

from pathlib import Path

import pandas as pd
import pytest

from pipenet_converter.models import Node, Nozzle, Pipe, PipeNetwork
from pipenet_converter.remote_path import (
    RemotePathOptions,
    extract_remote_path,
    find_nearest_node,
    render_remote_path_png,
    write_remote_path_artifacts,
    write_remote_path_xlsx,
)


def _build_branch_network(branch_count: int = 3, heads_per_branch: int = 12) -> PipeNetwork:
    """Build a tree shaped pipe network with multiple branches off an alarm valve.

    Branch 1 is the farthest branch (longest trunk leg) so the most-remote head
    selection should land entirely inside branch 1.
    """
    network = PipeNetwork(title="branch sample")
    network.add_node(Node("VALVE", 0.0, 0.0, 0.0, "alarm_valve"))

    trunk_lengths = {0: 5.0, 1: 500.0, 2: 30.0}

    nozzle_counter = 1
    for branch in range(branch_count):
        branch_y = float(branch * 100)
        trunk_node = f"T{branch}"
        network.add_node(Node(trunk_node, 0.0, branch_y, 0.0, "junction"))
        network.add_pipe(
            Pipe(
                pipe_id=f"P_T{branch}",
                from_node="VALVE",
                to_node=trunk_node,
                diameter_m=0.1,
                length_m=trunk_lengths.get(branch, 50.0),
                rise_m=0.0,
            )
        )
        prev = trunk_node
        for index in range(heads_per_branch):
            head_node = f"H_{branch}_{index}"
            head_x = float((index + 1) * 5)
            network.add_node(Node(head_node, head_x, branch_y, 0.0, "head"))
            network.add_pipe(
                Pipe(
                    pipe_id=f"P_{branch}_{index}",
                    from_node=prev,
                    to_node=head_node,
                    diameter_m=0.032,
                    length_m=5.0,
                    rise_m=0.0,
                )
            )
            output_node = f"@/{nozzle_counter}"
            network.add_node(Node(output_node, head_x, branch_y, 0.0, "head_output"))
            network.add_nozzle(
                Nozzle(
                    nozzle_id=str(nozzle_counter),
                    input_node=head_node,
                    output_node=output_node,
                    flow_m3s=0.00266,
                    status=1,
                )
            )
            nozzle_counter += 1
            prev = head_node
    return network


def test_farthest_head_is_in_longest_branch() -> None:
    network = _build_branch_network()

    result = extract_remote_path(network, "VALVE", RemotePathOptions(head_count=5))

    assert result.farthest_head_original_id.startswith("H_1_")
    assert result.farthest_head_renamed_id == "H01"


def test_remote_selection_keeps_neighbouring_heads() -> None:
    network = _build_branch_network()

    result = extract_remote_path(network, "VALVE", RemotePathOptions(head_count=10))

    original_ids = set(result.head_renumber_map)
    branches_picked = {hid.split("_")[1] for hid in original_ids}
    assert branches_picked == {"1"}
    assert len(original_ids) == 10


def test_extract_remote_path_renames_nodes_and_pipes() -> None:
    network = _build_branch_network()

    result = extract_remote_path(network, "VALVE", RemotePathOptions(head_count=5))

    assert "ALARM_VALVE" in result.network.nodes
    assert all(new_id.startswith("H") for new_id in result.head_renumber_map.values())
    assert all(pipe_id.startswith("P") for pipe_id in result.pipe_renumber_map.values())
    assert all(pipe.pipe_id.startswith("P") for pipe in result.network.pipes.values())


def test_extract_remote_path_subtree_connects_valve_to_each_head() -> None:
    network = _build_branch_network()

    result = extract_remote_path(network, "VALVE", RemotePathOptions(head_count=4))

    sub = result.network
    import networkx as nx

    graph = nx.Graph()
    for node_id in sub.nodes:
        graph.add_node(node_id)
    for pipe in sub.pipes.values():
        graph.add_edge(pipe.from_node, pipe.to_node, length=pipe.length_m)

    for head_renamed in result.head_renumber_map.values():
        assert nx.has_path(graph, "ALARM_VALVE", head_renamed)


def test_extract_remote_path_rejects_unknown_alarm_valve() -> None:
    network = _build_branch_network()

    with pytest.raises(KeyError):
        extract_remote_path(network, "NOPE")


def test_extract_remote_path_requires_positive_head_count() -> None:
    network = _build_branch_network()

    with pytest.raises(ValueError):
        extract_remote_path(network, "VALVE", RemotePathOptions(head_count=0))


def test_render_remote_path_png_writes_file(tmp_path: Path) -> None:
    network = _build_branch_network()
    result = extract_remote_path(network, "VALVE", RemotePathOptions(head_count=5))

    output_path = tmp_path / "remote.png"
    render_remote_path_png(result, output_path)

    assert output_path.exists()
    assert output_path.stat().st_size > 0


def test_write_remote_path_xlsx_has_two_sheets(tmp_path: Path) -> None:
    network = _build_branch_network()
    result = extract_remote_path(network, "VALVE", RemotePathOptions(head_count=6))

    output_path = tmp_path / "remote.xlsx"
    write_remote_path_xlsx(result, output_path)

    with pd.ExcelFile(output_path) as xlsx:
        assert set(xlsx.sheet_names) == {"Pipe Schedule", "Remote Heads"}
        pipe_df = pd.read_excel(xlsx, sheet_name="Pipe Schedule")
        head_df = pd.read_excel(xlsx, sheet_name="Remote Heads")
    assert len(head_df) == 6
    assert {"pipe_id", "diameter_label", "length_m"}.issubset(pipe_df.columns)
    assert head_df.iloc[0]["head_no"] == "H01"


def test_write_remote_path_artifacts_writes_both_files(tmp_path: Path) -> None:
    network = _build_branch_network()
    result = extract_remote_path(network, "VALVE", RemotePathOptions(head_count=4))

    artifacts = write_remote_path_artifacts(result, tmp_path)

    assert artifacts.isometric_png.exists() and artifacts.isometric_png.stat().st_size > 0
    assert artifacts.pipe_schedule_xlsx.exists() and artifacts.pipe_schedule_xlsx.stat().st_size > 0


def test_find_nearest_node_returns_closest_node() -> None:
    network = _build_branch_network()

    nearest = find_nearest_node(network, (0.4, 0.1))

    assert nearest == "VALVE"
