"""Tests for network validation."""

from pathlib import Path

from pipenet_converter.models import Node, Nozzle, Pipe, PipeNetwork
from pipenet_converter.validator import (
    apply_detected_fittings,
    detect_diameter_change_nodes,
    detect_elbow_nodes,
    detect_tee_nodes,
    orient_network_from_input,
    validate_network,
    write_validation_report,
)


def _issue_codes(network: PipeNetwork, input_node_id: str | None = None) -> set[str]:
    return {issue.code for issue in validate_network(network, input_node_id=input_node_id)}


def test_missing_pipe_node_is_error() -> None:
    network = PipeNetwork(title="Missing node")
    network.add_node(Node("N1", 0.0, 0.0, 0.0, "Input"))
    network.add_pipe(Pipe("P1", "N1", "MISSING", 0.15, 1.0, 0.0))

    issues = validate_network(network)

    assert any(
        issue.severity == "ERROR" and issue.code == "PIPE_TO_NODE_MISSING" for issue in issues
    )


def test_zero_length_pipe_is_error() -> None:
    network = PipeNetwork(title="Zero length")
    network.add_node(Node("N1", 0.0, 0.0, 0.0, "Input"))
    network.add_node(Node("N2", 1.0, 0.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "N1", "N2", 0.15, 0.0, 0.0))

    assert "PIPE_LENGTH_INVALID" in _issue_codes(network)


def test_invalid_diameter_and_c_factor_are_errors() -> None:
    network = PipeNetwork(title="Invalid pipe")
    network.add_node(Node("N1", 0.0, 0.0, 0.0, "Input"))
    network.add_node(Node("N2", 1.0, 0.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "N1", "N2", 0.0, 1.0, 0.0, c_factor=0.0))

    codes = _issue_codes(network)
    assert "PIPE_DIAMETER_INVALID" in codes
    assert "PIPE_C_FACTOR_INVALID" in codes


def test_disconnected_active_nozzle_is_error() -> None:
    network = PipeNetwork(title="Disconnected nozzle")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input"))
    network.add_node(Node("A", 1.0, 0.0, 0.0, "No"))
    network.add_node(Node("B", 10.0, 0.0, 0.0, "No"))
    network.add_node(Node("C", 11.0, 0.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "INPUT", "A", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P2", "B", "C", 0.05, 1.0, 0.0))
    network.add_nozzle(Nozzle("NZ1", "C", "@/1", 0.00266666667, status=1))

    issues = validate_network(network, input_node_id="INPUT")

    assert any(
        issue.severity == "ERROR" and issue.code == "ACTIVE_NOZZLE_UNREACHABLE"
        for issue in issues
    )


def test_tee_candidate_is_detected() -> None:
    network = PipeNetwork(title="Tee")
    network.add_node(Node("C", 0.0, 0.0, 0.0, "No"))
    network.add_node(Node("N", 0.0, 1.0, 0.0, "No"))
    network.add_node(Node("E", 1.0, 0.0, 0.0, "No"))
    network.add_node(Node("S", 0.0, -1.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "C", "N", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P2", "C", "E", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P3", "C", "S", 0.15, 1.0, 0.0))

    assert "TEE_CANDIDATE" in _issue_codes(network)


def test_elbow_candidate_and_diameter_change_are_detected() -> None:
    network = PipeNetwork(title="Elbow")
    network.add_node(Node("A", -1.0, 0.0, 0.0, "No"))
    network.add_node(Node("B", 0.0, 0.0, 0.0, "No"))
    network.add_node(Node("C", 0.0, 1.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "A", "B", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P2", "B", "C", 0.05, 1.0, 0.0))

    codes = _issue_codes(network)

    assert "ELBOW_CANDIDATE" in codes
    assert "DIAMETER_CHANGE_CANDIDATE" in codes


def test_write_validation_report_creates_csv(tmp_path: Path) -> None:
    network = PipeNetwork(title="Report")
    network.add_node(Node("N1", 0.0, 0.0, 0.0, "Input"))
    network.add_pipe(Pipe("P1", "N1", "MISSING", 0.15, 1.0, 0.0))
    issues = validate_network(network)
    output_csv = tmp_path / "validation.csv"

    write_validation_report(issues, output_csv)

    report_text = output_csv.read_text(encoding="utf-8")
    assert "severity,code,message,object_type,object_id" in report_text
    assert "PIPE_TO_NODE_MISSING" in report_text


def test_t_shape_gets_tee_fitting() -> None:
    network = _t_network()

    assert detect_tee_nodes(network) == {"C": 1}

    apply_detected_fittings(network)

    assert [(fitting.fitting_type, fitting.count) for fitting in network.pipes["P1"].fittings] == [
        ("tee", 1)
    ]


def test_l_shape_gets_elbow_fitting() -> None:
    network = _l_network()

    assert detect_elbow_nodes(network) == {"B": 1}

    apply_detected_fittings(network)

    assert [(fitting.fitting_type, fitting.count) for fitting in network.pipes["P1"].fittings] == [
        ("elbow", 1)
    ]


def test_straight_line_does_not_get_elbow() -> None:
    network = PipeNetwork(title="Straight")
    network.add_node(Node("A", -1.0, 0.0, 0.0, "No"))
    network.add_node(Node("B", 0.0, 0.0, 0.0, "No"))
    network.add_node(Node("C", 1.0, 0.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "A", "B", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P2", "B", "C", 0.15, 1.0, 0.0))

    assert detect_elbow_nodes(network) == {}

    apply_detected_fittings(network)

    assert all(not pipe.fittings for pipe in network.pipes.values())


def test_diameter_change_nodes_are_detected_and_marked() -> None:
    network = _l_network()

    changes = detect_diameter_change_nodes(network)

    assert changes == {"B": [0.05, 0.15]}
    assert network.nodes["B"].metadata["diameter_change"] is True
    assert network.nodes["B"].metadata["diameters"] == "0.05,0.15"


def test_fitting_assignment_is_deterministic_without_root() -> None:
    network = _t_network()
    network.add_pipe(Pipe("A_FIRST", "C", "W", 0.15, 1.0, 0.0))
    network.add_node(Node("W", -1.0, -1.0, 0.0, "No"))

    apply_detected_fittings(network)

    assert [(fitting.fitting_type, fitting.count) for fitting in network.pipes["A_FIRST"].fittings] == [
        ("tee", 1)
    ]
    assert all(pipe_id == "A_FIRST" or not pipe.fittings for pipe_id, pipe in network.pipes.items())


def test_orient_network_from_input_flips_upstream_pipes() -> None:
    network = PipeNetwork(title="Orient")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input"))
    network.add_node(Node("A", 1.0, 0.0, 1.0, "No"))
    network.add_node(Node("B", 2.0, 0.0, 2.0, "No"))
    network.add_pipe(Pipe("P1", "A", "INPUT", 0.15, 1.0, -1.0))
    network.add_pipe(Pipe("P2", "B", "A", 0.15, 1.0, -1.0))

    orient_network_from_input(network, "INPUT")

    assert network.pipes["P1"].from_node == "INPUT"
    assert network.pipes["P1"].to_node == "A"
    assert network.pipes["P1"].rise_m == 1.0
    assert network.pipes["P2"].from_node == "A"
    assert network.pipes["P2"].to_node == "B"
    assert network.pipes["P2"].rise_m == 1.0


def _t_network() -> PipeNetwork:
    network = PipeNetwork(title="T")
    network.add_node(Node("C", 0.0, 0.0, 0.0, "No"))
    network.add_node(Node("N", 0.0, 1.0, 0.0, "No"))
    network.add_node(Node("E", 1.0, 0.0, 0.0, "No"))
    network.add_node(Node("S", 0.0, -1.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "C", "N", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P2", "C", "E", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P3", "C", "S", 0.15, 1.0, 0.0))
    return network


def _l_network() -> PipeNetwork:
    network = PipeNetwork(title="L")
    network.add_node(Node("A", -1.0, 0.0, 0.0, "No"))
    network.add_node(Node("B", 0.0, 0.0, 0.0, "No"))
    network.add_node(Node("C", 0.0, 1.0, 0.0, "No"))
    network.add_pipe(Pipe("P1", "A", "B", 0.15, 1.0, 0.0))
    network.add_pipe(Pipe("P2", "B", "C", 0.05, 1.0, 0.0))
    return network
