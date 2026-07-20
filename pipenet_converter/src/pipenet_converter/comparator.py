"""Comparison utilities for reference and generated PipeNet networks."""

from __future__ import annotations

import csv
from dataclasses import dataclass
from pathlib import Path

from pipenet_converter.models import Pipe, PipeNetwork

FLOW_TOLERANCE_M3S = 1e-9


@dataclass(slots=True)
class CompareTolerance:
    """Numeric tolerances for network comparison."""

    coordinate_m: float = 0.1
    length_m: float = 0.1
    diameter_m: float = 0.001
    elevation_m: float = 0.05


@dataclass(slots=True)
class ComparisonIssue:
    """A discrepancy between a reference and candidate network."""

    severity: str
    issue_type: str
    reference_id: str | None
    candidate_id: str | None
    message: str
    reference_value: str | float | int | None
    candidate_value: str | float | int | None
    delta: float | None


def compare_networks(
    reference: PipeNetwork,
    candidate: PipeNetwork,
    tolerance: CompareTolerance,
) -> list[ComparisonIssue]:
    """Compare two networks by existing labels.

    Future extension point: add spatial/geometric fuzzy matching before or after
    label-based matching for networks generated from CAD without stable labels.
    """
    issues: list[ComparisonIssue] = []
    issues.extend(_compare_counts(reference, candidate))
    issues.extend(_compare_nodes(reference, candidate, tolerance))
    issues.extend(_compare_pipes(reference, candidate, tolerance))
    issues.extend(_compare_nozzles(reference, candidate, tolerance))
    return issues


def write_comparison_report(issues: list[ComparisonIssue], output_csv: str | Path) -> None:
    """Write comparison issues to CSV."""
    output_path = Path(output_csv)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=[
                "severity",
                "issue_type",
                "reference_id",
                "candidate_id",
                "message",
                "reference_value",
                "candidate_value",
                "delta",
            ],
        )
        writer.writeheader()
        for issue in issues:
            writer.writerow(
                {
                    "severity": issue.severity,
                    "issue_type": issue.issue_type,
                    "reference_id": issue.reference_id,
                    "candidate_id": issue.candidate_id,
                    "message": issue.message,
                    "reference_value": issue.reference_value,
                    "candidate_value": issue.candidate_value,
                    "delta": issue.delta,
                }
            )


def _compare_counts(reference: PipeNetwork, candidate: PipeNetwork) -> list[ComparisonIssue]:
    count_checks = [
        ("NODE_COUNT_MISMATCH", "node count", len(reference.nodes), len(candidate.nodes)),
        ("PIPE_COUNT_MISMATCH", "pipe count", len(reference.pipes), len(candidate.pipes)),
        ("NOZZLE_COUNT_MISMATCH", "nozzle count", len(reference.nozzles), len(candidate.nozzles)),
        (
            "ACTIVE_NOZZLE_COUNT_MISMATCH",
            "active nozzle count",
            len(reference.active_nozzles()),
            len(candidate.active_nozzles()),
        ),
    ]
    return [
        ComparisonIssue(
            severity="ERROR",
            issue_type=issue_type,
            reference_id=None,
            candidate_id=None,
            message=f"Reference {label} differs from candidate {label}.",
            reference_value=reference_count,
            candidate_value=candidate_count,
            delta=float(candidate_count - reference_count),
        )
        for issue_type, label, reference_count, candidate_count in count_checks
        if reference_count != candidate_count
    ]


def _compare_nodes(
    reference: PipeNetwork,
    candidate: PipeNetwork,
    tolerance: CompareTolerance,
) -> list[ComparisonIssue]:
    issues: list[ComparisonIssue] = []
    for node_id, reference_node in reference.nodes.items():
        candidate_node = candidate.nodes.get(node_id)
        if candidate_node is None:
            issues.append(
                ComparisonIssue(
                    "ERROR",
                    "MISSING_NODE_LABEL",
                    node_id,
                    None,
                    f"Candidate is missing node {node_id}.",
                    "present",
                    "missing",
                    None,
                )
            )
            continue

        delta_z = candidate_node.z - reference_node.z
        if abs(delta_z) > tolerance.elevation_m:
            issues.append(
                ComparisonIssue(
                    "WARNING",
                    "ELEVATION_MISMATCH",
                    node_id,
                    node_id,
                    f"Node {node_id} elevation differs.",
                    reference_node.z,
                    candidate_node.z,
                    delta_z,
                )
            )

    return issues


def _compare_pipes(
    reference: PipeNetwork,
    candidate: PipeNetwork,
    tolerance: CompareTolerance,
) -> list[ComparisonIssue]:
    issues: list[ComparisonIssue] = []
    for pipe_id, reference_pipe in reference.pipes.items():
        candidate_pipe = candidate.pipes.get(pipe_id)
        if candidate_pipe is None:
            issues.append(
                ComparisonIssue(
                    "ERROR",
                    "MISSING_PIPE_LABEL",
                    pipe_id,
                    None,
                    f"Candidate is missing pipe {pipe_id}.",
                    "present",
                    "missing",
                    None,
                )
            )
            continue

        issues.extend(_compare_pipe_numeric_values(reference_pipe, candidate_pipe, tolerance))
        if _fitting_counts(reference_pipe) != _fitting_counts(candidate_pipe):
            issues.append(
                ComparisonIssue(
                    "WARNING",
                    "FITTING_MISMATCH",
                    pipe_id,
                    pipe_id,
                    f"Pipe {pipe_id} fitting type/count differs.",
                    _fitting_counts_text(reference_pipe),
                    _fitting_counts_text(candidate_pipe),
                    None,
                )
            )

    return issues


def _compare_pipe_numeric_values(
    reference_pipe: Pipe,
    candidate_pipe: Pipe,
    tolerance: CompareTolerance,
) -> list[ComparisonIssue]:
    specs = [
        ("PIPE_LENGTH_MISMATCH", "length", reference_pipe.length_m, candidate_pipe.length_m, tolerance.length_m),
        (
            "PIPE_DIAMETER_MISMATCH",
            "diameter",
            reference_pipe.diameter_m,
            candidate_pipe.diameter_m,
            tolerance.diameter_m,
        ),
        ("PIPE_RISE_MISMATCH", "rise", reference_pipe.rise_m, candidate_pipe.rise_m, tolerance.elevation_m),
    ]
    return [
        ComparisonIssue(
            "WARNING",
            issue_type,
            reference_pipe.pipe_id,
            candidate_pipe.pipe_id,
            f"Pipe {reference_pipe.pipe_id} {label} differs.",
            reference_value,
            candidate_value,
            candidate_value - reference_value,
        )
        for issue_type, label, reference_value, candidate_value, allowed_delta in specs
        if abs(candidate_value - reference_value) > allowed_delta
    ]


def _compare_nozzles(
    reference: PipeNetwork,
    candidate: PipeNetwork,
    tolerance: CompareTolerance,
) -> list[ComparisonIssue]:
    issues: list[ComparisonIssue] = []
    for nozzle_id, reference_nozzle in reference.nozzles.items():
        candidate_nozzle = candidate.nozzles.get(nozzle_id)
        if candidate_nozzle is None:
            issues.append(
                ComparisonIssue(
                    "ERROR",
                    "MISSING_NOZZLE_LABEL",
                    nozzle_id,
                    None,
                    f"Candidate is missing nozzle {nozzle_id}.",
                    "present",
                    "missing",
                    None,
                )
            )
            continue

        flow_delta = candidate_nozzle.flow_m3s - reference_nozzle.flow_m3s
        if abs(flow_delta) > FLOW_TOLERANCE_M3S:
            issues.append(
                ComparisonIssue(
                    "WARNING",
                    "NOZZLE_FLOW_MISMATCH",
                    nozzle_id,
                    nozzle_id,
                    f"Nozzle {nozzle_id} flow differs.",
                    reference_nozzle.flow_m3s,
                    candidate_nozzle.flow_m3s,
                    flow_delta,
                )
            )
        if reference_nozzle.status != candidate_nozzle.status:
            issues.append(
                ComparisonIssue(
                    "WARNING",
                    "NOZZLE_STATUS_MISMATCH",
                    nozzle_id,
                    nozzle_id,
                    f"Nozzle {nozzle_id} status differs.",
                    reference_nozzle.status,
                    candidate_nozzle.status,
                    float(candidate_nozzle.status - reference_nozzle.status),
                )
            )

    return issues


def _fitting_counts(pipe: Pipe) -> dict[str, int]:
    counts: dict[str, int] = {}
    for fitting in pipe.fittings:
        counts[fitting.fitting_type] = counts.get(fitting.fitting_type, 0) + fitting.count
    return counts


def _fitting_counts_text(pipe: Pipe) -> str:
    counts = _fitting_counts(pipe)
    return ";".join(f"{fitting_type}:{count}" for fitting_type, count in sorted(counts.items()))
