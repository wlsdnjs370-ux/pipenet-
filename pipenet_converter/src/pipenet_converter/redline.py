"""Machine-readable redline data for CAD/SDF discrepancy review."""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass
from pathlib import Path

import pandas as pd

from pipenet_converter.comparator import ComparisonIssue
from pipenet_converter.models import Node, PipeNetwork


@dataclass(slots=True)
class RedlineItem:
    """A point annotation describing a CAD/SDF mismatch."""

    redline_id: str
    issue_type: str
    severity: str
    object_type: str
    object_id: str
    x: float
    y: float
    z: float
    message: str
    suggested_action: str


REDLINE_COLUMNS = [
    "redline_id",
    "issue_type",
    "severity",
    "object_type",
    "object_id",
    "x",
    "y",
    "z",
    "message",
    "suggested_action",
]


def create_redline_items(
    issues: list[ComparisonIssue],
    reference: PipeNetwork,
    candidate: PipeNetwork,
) -> list[RedlineItem]:
    """Create redline point annotations from comparison issues."""
    items: list[RedlineItem] = []
    for index, issue in enumerate(issues, start=1):
        object_type, object_id, point = _resolve_issue_location(issue, reference, candidate)
        if point is None or object_id is None:
            continue
        items.append(
            RedlineItem(
                redline_id=f"RL{index:06d}",
                issue_type=issue.issue_type,
                severity=issue.severity,
                object_type=object_type,
                object_id=object_id,
                x=point[0],
                y=point[1],
                z=point[2],
                message=issue.message,
                suggested_action=_suggested_action(issue.issue_type),
            )
        )
    return items


def write_redline_items(items: list[RedlineItem], output_dir: str | Path) -> None:
    """Write redline items to CSV and JSON."""
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)
    rows = [asdict(item) for item in items]
    pd.DataFrame(rows, columns=REDLINE_COLUMNS).to_csv(output_path / "redline_items.csv", index=False)
    (output_path / "redline_items.json").write_text(
        json.dumps(rows, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _resolve_issue_location(
    issue: ComparisonIssue,
    reference: PipeNetwork,
    candidate: PipeNetwork,
) -> tuple[str, str | None, tuple[float, float, float] | None]:
    object_id = issue.candidate_id or issue.reference_id
    if object_id is None:
        return "Network", None, None

    if _is_pipe_issue(issue.issue_type):
        return "Pipe", object_id, _pipe_midpoint(candidate, object_id) or _pipe_midpoint(reference, object_id)
    if _is_nozzle_issue(issue.issue_type):
        return "Nozzle", object_id, _nozzle_point(candidate, object_id) or _nozzle_point(reference, object_id)
    if _is_node_issue(issue.issue_type):
        return "Node", object_id, _node_point(candidate, object_id) or _node_point(reference, object_id)

    return "Object", object_id, _node_point(candidate, object_id) or _node_point(reference, object_id)


def _pipe_midpoint(network: PipeNetwork, pipe_id: str) -> tuple[float, float, float] | None:
    pipe = network.pipes.get(pipe_id)
    if pipe is None:
        return None
    from_node = network.nodes.get(pipe.from_node)
    to_node = network.nodes.get(pipe.to_node)
    if from_node is None or to_node is None:
        return None
    return (
        (from_node.x + to_node.x) / 2,
        (from_node.y + to_node.y) / 2,
        (from_node.z + to_node.z) / 2,
    )


def _nozzle_point(network: PipeNetwork, nozzle_id: str) -> tuple[float, float, float] | None:
    nozzle = network.nozzles.get(nozzle_id)
    if nozzle is None:
        return None
    output_node = network.nodes.get(nozzle.output_node)
    input_node = network.nodes.get(nozzle.input_node)
    node = output_node or input_node
    return _point_from_node(node)


def _node_point(network: PipeNetwork, node_id: str) -> tuple[float, float, float] | None:
    return _point_from_node(network.nodes.get(node_id))


def _point_from_node(node: Node | None) -> tuple[float, float, float] | None:
    if node is None:
        return None
    return (node.x, node.y, node.z)


def _is_pipe_issue(issue_type: str) -> bool:
    return issue_type.startswith("PIPE_") or issue_type in {"FITTING_MISMATCH", "MISSING_PIPE_LABEL"}


def _is_nozzle_issue(issue_type: str) -> bool:
    return issue_type.startswith("NOZZLE_") or issue_type in {"MISSING_NOZZLE_LABEL"}


def _is_node_issue(issue_type: str) -> bool:
    return issue_type.startswith("NODE_") or issue_type in {"MISSING_NODE_LABEL", "ELEVATION_MISMATCH"}


def _suggested_action(issue_type: str) -> str:
    if "LENGTH_MISMATCH" in issue_type:
        return "Check CAD centerline length and SDF pipe length"
    if "DIAMETER_MISMATCH" in issue_type:
        return "Check CAD diameter text and SDF bore"
    if "FITTING_MISMATCH" in issue_type:
        return "Check elbow/tee count"
    if "ELEVATION_MISMATCH" in issue_type or "RISE_MISMATCH" in issue_type:
        return "Check section/system diagram Z rule"
    if "MISSING" in issue_type:
        return "Check extraction mapping or SDF model"
    return "Review CAD-derived network and SDF model"
