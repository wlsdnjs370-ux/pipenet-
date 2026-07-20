"""Tests for system/riser CSV integration."""

from pathlib import Path

from pipenet_converter.models import Node, Nozzle, Pipe, PipeNetwork
from pipenet_converter.system_graph import merge_networks, parse_system_edges_csv
from pipenet_converter.validator import validate_network


SYSTEM_EDGES_CSV = """edge_id,from_node,to_node,from_z_m,to_z_m,diameter_m,length_m,rise_m,c_factor,material,fittings,equipment,description
R001,INPUT,B1F_UPPER,0,5.85,0.2,5.85,5.85,120,KSD3507,"gate:1;check:1","",supply to B1F
R002,B1F_UPPER,4F_UPPER_MAIN,5.85,48.55,0.15,42.7,42.7,120,KSD3507,"elbow:1","AV:24",alarm valve to upper warehouse
"""


def test_parse_system_edges_csv_creates_nodes_pipes_fittings_and_equipment(tmp_path: Path) -> None:
    csv_path = tmp_path / "system_edges.csv"
    csv_path.write_text(SYSTEM_EDGES_CSV, encoding="utf-8")

    network = parse_system_edges_csv(csv_path)

    assert set(network.nodes) == {"INPUT", "B1F_UPPER", "4F_UPPER_MAIN"}
    assert network.nodes["INPUT"].z == 0.0
    assert network.nodes["4F_UPPER_MAIN"].z == 48.55
    assert set(network.pipes) == {"R001", "R002"}

    pipe_1 = network.pipes["R001"]
    assert [(fitting.fitting_type, fitting.count) for fitting in pipe_1.fittings] == [
        ("gate", 1),
        ("check", 1),
    ]

    pipe_2 = network.pipes["R002"]
    assert pipe_2.fittings[0].fitting_type == "elbow"
    assert pipe_2.equipment[0].description == "AV"
    assert pipe_2.equipment[0].equivalent_length_m == 24.0
    assert pipe_2.material == "KSD3507"


def test_merge_networks_connects_system_to_plan_network(tmp_path: Path) -> None:
    csv_path = tmp_path / "system_edges.csv"
    csv_path.write_text(SYSTEM_EDGES_CSV, encoding="utf-8")
    system_network = parse_system_edges_csv(csv_path)
    plan_network = _plan_network()

    merged = merge_networks(system_network, plan_network, {"4F_UPPER_MAIN": "N001"})

    assert "INPUT" in merged.nodes
    assert "N001" in merged.nodes
    assert "CONN_001" in merged.pipes
    assert merged.pipes["CONN_001"].from_node == "4F_UPPER_MAIN"
    assert merged.pipes["CONN_001"].to_node == "N001"
    assert "NZ1" in merged.nozzles


def test_active_nozzles_are_reachable_from_input_after_merge(tmp_path: Path) -> None:
    csv_path = tmp_path / "system_edges.csv"
    csv_path.write_text(SYSTEM_EDGES_CSV, encoding="utf-8")
    system_network = parse_system_edges_csv(csv_path)
    plan_network = _plan_network()

    merged = merge_networks(system_network, plan_network, {"4F_UPPER_MAIN": "N001"})
    issues = validate_network(merged, input_node_id="INPUT")

    assert "ACTIVE_NOZZLE_UNREACHABLE" not in {issue.code for issue in issues}


def test_merge_prefixes_conflicting_system_node_ids(tmp_path: Path) -> None:
    csv_path = tmp_path / "system_edges.csv"
    csv_path.write_text(SYSTEM_EDGES_CSV, encoding="utf-8")
    system_network = parse_system_edges_csv(csv_path)
    plan_network = _plan_network()
    plan_network.add_node(Node("INPUT", 100.0, 100.0, 48.55, "plan_input"))

    merged = merge_networks(system_network, plan_network, {"4F_UPPER_MAIN": "N001"})

    assert "SYS_INPUT" in merged.nodes
    assert "INPUT" in merged.nodes
    assert merged.pipes["R001"].from_node == "SYS_INPUT"


def _plan_network() -> PipeNetwork:
    network = PipeNetwork(title="Plan")
    network.add_node(Node("N001", 0.0, 1.0, 48.55, "main_pipe"))
    network.add_node(Node("N002", 10.0, 1.0, 48.55, "branch_pipe"))
    network.add_node(Node("@/1", 10.0, 1.0, 49.15, "head_output"))
    network.add_pipe(Pipe("P001", "N001", "N002", 0.032, 10.0, 0.0))
    network.add_nozzle(Nozzle("NZ1", "N002", "@/1", 0.00266666667, status=1))
    return network
