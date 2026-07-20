"""Tests for network comparison."""

from pathlib import Path

from pipenet_converter.comparator import CompareTolerance, compare_networks, write_comparison_report
from pipenet_converter.models import Fitting, Node, Nozzle, Pipe, PipeNetwork


def test_length_mismatch_detected() -> None:
    reference = _network()
    candidate = _network()
    candidate.pipes["P001"].length_m = 2.0

    issues = compare_networks(reference, candidate, CompareTolerance(length_m=0.1))

    assert "PIPE_LENGTH_MISMATCH" in {issue.issue_type for issue in issues}


def test_diameter_mismatch_detected() -> None:
    reference = _network()
    candidate = _network()
    candidate.pipes["P001"].diameter_m = 0.1

    issues = compare_networks(reference, candidate, CompareTolerance(diameter_m=0.001))

    assert "PIPE_DIAMETER_MISMATCH" in {issue.issue_type for issue in issues}


def test_missing_nozzle_detected() -> None:
    reference = _network()
    candidate = _network()
    del candidate.nozzles["NZ1"]

    issues = compare_networks(reference, candidate, CompareTolerance())

    issue_types = {issue.issue_type for issue in issues}
    assert "NOZZLE_COUNT_MISMATCH" in issue_types
    assert "ACTIVE_NOZZLE_COUNT_MISMATCH" in issue_types
    assert "MISSING_NOZZLE_LABEL" in issue_types


def test_fitting_mismatch_detected() -> None:
    reference = _network()
    candidate = _network()
    candidate.pipes["P001"].fittings = [Fitting("elbow", 1)]

    issues = compare_networks(reference, candidate, CompareTolerance())

    assert "FITTING_MISMATCH" in {issue.issue_type for issue in issues}


def test_elevation_and_nozzle_status_mismatches_detected() -> None:
    reference = _network()
    candidate = _network()
    candidate.nodes["N001"].z = 0.2
    candidate.nozzles["NZ1"].status = 0

    issues = compare_networks(reference, candidate, CompareTolerance(elevation_m=0.05))

    issue_types = {issue.issue_type for issue in issues}
    assert "ELEVATION_MISMATCH" in issue_types
    assert "ACTIVE_NOZZLE_COUNT_MISMATCH" in issue_types
    assert "NOZZLE_STATUS_MISMATCH" in issue_types


def test_write_comparison_report_creates_csv(tmp_path: Path) -> None:
    reference = _network()
    candidate = _network()
    candidate.pipes["P001"].length_m = 2.0
    issues = compare_networks(reference, candidate, CompareTolerance(length_m=0.1))
    output_csv = tmp_path / "compare_report.csv"

    write_comparison_report(issues, output_csv)

    report_text = output_csv.read_text(encoding="utf-8")
    assert "severity,issue_type,reference_id,candidate_id,message,reference_value,candidate_value,delta" in report_text
    assert "PIPE_LENGTH_MISMATCH" in report_text


def _network() -> PipeNetwork:
    network = PipeNetwork(title="Compare")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input"))
    network.add_node(Node("N001", 1.0, 0.0, 0.0, "main_pipe"))
    network.add_node(Node("@/1", 1.0, 0.0, 0.0, "head_output"))
    network.add_pipe(
        Pipe(
            "P001",
            "INPUT",
            "N001",
            0.15,
            1.0,
            0.0,
            fittings=[Fitting("gate", 1)],
        )
    )
    network.add_nozzle(Nozzle("NZ1", "N001", "@/1", 0.00266666667, status=1))
    return network
