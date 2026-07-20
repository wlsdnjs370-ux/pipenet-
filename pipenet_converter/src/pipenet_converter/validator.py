"""Validation utilities for generated sprinkler networks and SDF output."""

from __future__ import annotations

import csv
from math import acos, degrees, sqrt
from pathlib import Path

import networkx as nx

from pipenet_converter.models import Fitting, Node, Pipe, PipeNetwork, ValidationIssue

RISE_TOLERANCE_M = 0.001
ELBOW_ANGLE_THRESHOLD_DEGREES = 10.0


def validate_network(network: PipeNetwork, input_node_id: str | None = None) -> list[ValidationIssue]:
    """Validate a sprinkler network before SDF export."""
    issues: list[ValidationIssue] = []

    for pipe in network.pipes.values():
        issues.extend(_validate_pipe_references_and_values(network, pipe))

    for nozzle in network.nozzles.values():
        if nozzle.input_node not in network.nodes:
            issues.append(
                ValidationIssue(
                    "ERROR",
                    "NOZZLE_INPUT_NODE_MISSING",
                    f"Nozzle {nozzle.nozzle_id} input node {nozzle.input_node} does not exist.",
                    "Nozzle",
                    nozzle.nozzle_id,
                )
            )
        if not nozzle.output_node.startswith("@/") and nozzle.output_node not in network.nodes:
            issues.append(
                ValidationIssue(
                    "ERROR",
                    "NOZZLE_OUTPUT_NODE_MISSING",
                    f"Nozzle {nozzle.nozzle_id} output node {nozzle.output_node} does not exist.",
                    "Nozzle",
                    nozzle.nozzle_id,
                )
            )

    graph = _build_graph(network)
    issues.extend(_detect_orphan_nodes(graph))
    issues.extend(_detect_connectivity_issues(network, graph, input_node_id))
    issues.extend(_detect_fitting_candidates(network, graph))

    return issues


def write_validation_report(issues: list[ValidationIssue], output_csv: str | Path) -> None:
    """Write validation issues to a CSV file."""
    output_path = Path(output_csv)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", newline="", encoding="utf-8") as csv_file:
        writer = csv.DictWriter(
            csv_file,
            fieldnames=["severity", "code", "message", "object_type", "object_id"],
        )
        writer.writeheader()
        for issue in issues:
            writer.writerow(
                {
                    "severity": issue.severity,
                    "code": issue.code,
                    "message": issue.message,
                    "object_type": issue.object_type,
                    "object_id": issue.object_id,
                }
            )


def detect_tee_nodes(network: PipeNetwork) -> dict[str, int]:
    """Return nodes with graph degree >= 3 as tee candidates."""
    graph = _build_graph(network)
    return {str(node_id): 1 for node_id, degree in graph.degree if degree >= 3}


def detect_elbow_nodes(
    network: PipeNetwork,
    angle_threshold_deg: float = ELBOW_ANGLE_THRESHOLD_DEGREES,
) -> dict[str, int]:
    """Return degree-2 nodes with a direction change above ``angle_threshold_deg``."""
    graph = _build_graph(network)
    elbows: dict[str, int] = {}
    for node_id, degree in graph.degree:
        node_id = str(node_id)
        if degree == 2 and _is_elbow_candidate(network, graph, node_id, angle_threshold_deg):
            elbows[node_id] = 1
    return elbows


def detect_diameter_change_nodes(network: PipeNetwork) -> dict[str, list[float]]:
    """Return nodes whose connected pipes have more than one diameter."""
    graph = _build_graph(network)
    diameter_changes: dict[str, list[float]] = {}
    for node_id in graph.nodes:
        connected_pipes = _connected_pipes(network, str(node_id))
        diameters = sorted({pipe.diameter_m for pipe in connected_pipes})
        if len(diameters) > 1:
            diameter_changes[str(node_id)] = diameters
            network.nodes[str(node_id)].metadata["diameter_change"] = True
            network.nodes[str(node_id)].metadata["diameters"] = ",".join(
                format(diameter, ".6g") for diameter in diameters
            )
    return diameter_changes


def orient_network_from_input(network: PipeNetwork, input_node_id: str) -> PipeNetwork:
    """Orient pipe directions by BFS distance from ``input_node_id`` where possible."""
    graph = _build_graph(network)
    if input_node_id not in graph:
        return network

    depths = nx.single_source_shortest_path_length(graph, input_node_id)
    for pipe in network.pipes.values():
        from_depth = depths.get(pipe.from_node)
        to_depth = depths.get(pipe.to_node)
        if from_depth is None or to_depth is None:
            continue
        if from_depth > to_depth:
            pipe.from_node, pipe.to_node = pipe.to_node, pipe.from_node
            if pipe.from_node in network.nodes and pipe.to_node in network.nodes:
                pipe.rise_m = network.nodes[pipe.to_node].z - network.nodes[pipe.from_node].z
    return network


def apply_detected_fittings(network: PipeNetwork, strategy: str = "downstream") -> PipeNetwork:
    """Apply detected tee and elbow fittings to deterministic pipe targets."""
    if strategy != "downstream":
        raise ValueError(f"Unsupported fitting assignment strategy: {strategy!r}")

    input_node_id = _infer_input_node_id(network)
    if input_node_id is not None:
        orient_network_from_input(network, input_node_id)

    graph = _build_graph(network)
    depths = nx.single_source_shortest_path_length(graph, input_node_id) if input_node_id in graph else {}

    for node_id, count in detect_tee_nodes(network).items():
        pipe = _select_fitting_pipe(network, node_id, depths)
        if pipe is not None:
            _add_or_increment_fitting(pipe, "tee", count)

    for node_id, count in detect_elbow_nodes(network).items():
        pipe = _select_fitting_pipe(network, node_id, depths)
        if pipe is not None:
            _add_or_increment_fitting(pipe, "elbow", count)

    detect_diameter_change_nodes(network)
    return network


def _validate_pipe_references_and_values(network: PipeNetwork, pipe: Pipe) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    from_node = network.nodes.get(pipe.from_node)
    to_node = network.nodes.get(pipe.to_node)

    if from_node is None:
        issues.append(
            ValidationIssue(
                "ERROR",
                "PIPE_FROM_NODE_MISSING",
                f"Pipe {pipe.pipe_id} from_node {pipe.from_node} does not exist.",
                "Pipe",
                pipe.pipe_id,
            )
        )
    if to_node is None:
        issues.append(
            ValidationIssue(
                "ERROR",
                "PIPE_TO_NODE_MISSING",
                f"Pipe {pipe.pipe_id} to_node {pipe.to_node} does not exist.",
                "Pipe",
                pipe.pipe_id,
            )
        )
    if pipe.length_m <= 0:
        issues.append(
            ValidationIssue(
                "ERROR",
                "PIPE_LENGTH_INVALID",
                f"Pipe {pipe.pipe_id} length_m must be greater than 0.",
                "Pipe",
                pipe.pipe_id,
            )
        )
    if pipe.diameter_m <= 0:
        issues.append(
            ValidationIssue(
                "ERROR",
                "PIPE_DIAMETER_INVALID",
                f"Pipe {pipe.pipe_id} diameter_m must be greater than 0.",
                "Pipe",
                pipe.pipe_id,
            )
        )
    if pipe.c_factor <= 0:
        issues.append(
            ValidationIssue(
                "ERROR",
                "PIPE_C_FACTOR_INVALID",
                f"Pipe {pipe.pipe_id} c_factor must be greater than 0.",
                "Pipe",
                pipe.pipe_id,
            )
        )

    if from_node is not None and to_node is not None:
        expected_rise = to_node.z - from_node.z
        if abs(pipe.rise_m - expected_rise) > RISE_TOLERANCE_M:
            issues.append(
                ValidationIssue(
                    "WARNING",
                    "PIPE_RISE_MISMATCH",
                    (
                        f"Pipe {pipe.pipe_id} rise_m {pipe.rise_m:g} differs from "
                        f"node elevation delta {expected_rise:g}."
                    ),
                    "Pipe",
                    pipe.pipe_id,
                )
            )

    return issues


def _build_graph(network: PipeNetwork) -> nx.Graph:
    graph = nx.Graph()
    for node_id in network.nodes:
        graph.add_node(node_id)
    for pipe in network.pipes.values():
        if pipe.from_node in network.nodes and pipe.to_node in network.nodes:
            graph.add_edge(pipe.from_node, pipe.to_node, pipe_id=pipe.pipe_id)
    return graph


def _detect_orphan_nodes(graph: nx.Graph) -> list[ValidationIssue]:
    return [
        ValidationIssue(
            "WARNING",
            "ORPHAN_NODE",
            f"Node {node_id} has degree 0.",
            "Node",
            str(node_id),
        )
        for node_id, degree in graph.degree
        if degree == 0
    ]


def _detect_connectivity_issues(
    network: PipeNetwork,
    graph: nx.Graph,
    input_node_id: str | None,
) -> list[ValidationIssue]:
    if input_node_id is None or input_node_id not in graph:
        return []

    reachable_nodes = nx.node_connected_component(graph, input_node_id)
    issues: list[ValidationIssue] = []
    for nozzle in network.nozzles.values():
        if nozzle.status != 1:
            continue
        if nozzle.input_node in network.nodes and nozzle.input_node not in reachable_nodes:
            issues.append(
                ValidationIssue(
                    "ERROR",
                    "ACTIVE_NOZZLE_UNREACHABLE",
                    (
                        f"Active nozzle {nozzle.nozzle_id} input node {nozzle.input_node} "
                        f"is not reachable from {input_node_id}."
                    ),
                    "Nozzle",
                    nozzle.nozzle_id,
                )
            )
    return issues


def _detect_fitting_candidates(network: PipeNetwork, graph: nx.Graph) -> list[ValidationIssue]:
    issues: list[ValidationIssue] = []
    tee_nodes = detect_tee_nodes(network)
    elbow_nodes = detect_elbow_nodes(network)
    diameter_change_nodes = detect_diameter_change_nodes(network)

    for node_id in tee_nodes:
        issues.append(
            ValidationIssue(
                "INFO",
                "TEE_CANDIDATE",
                f"Node {node_id} has graph degree {graph.degree[node_id]} and is a tee candidate.",
                "Node",
                node_id,
            )
        )

    for node_id in elbow_nodes:
        issues.append(
            ValidationIssue(
                "INFO",
                "ELBOW_CANDIDATE",
                f"Node {node_id} has an angle change greater than 10 degrees.",
                "Node",
                node_id,
            )
        )

    for node_id in diameter_change_nodes:
        issues.append(
            ValidationIssue(
                "INFO",
                "DIAMETER_CHANGE_CANDIDATE",
                f"Node {node_id} connects pipes with multiple diameters.",
                "Node",
                node_id,
            )
        )

    return issues


def _connected_pipes(network: PipeNetwork, node_id: str) -> list[Pipe]:
    return [
        pipe
        for pipe in network.pipes.values()
        if pipe.from_node == node_id or pipe.to_node == node_id
    ]


def _is_elbow_candidate(
    network: PipeNetwork,
    graph: nx.Graph,
    node_id: str,
    angle_threshold_deg: float = ELBOW_ANGLE_THRESHOLD_DEGREES,
) -> bool:
    neighbors = list(graph.neighbors(node_id))
    if len(neighbors) != 2:
        return False

    node = network.nodes[node_id]
    neighbor_a = network.nodes[str(neighbors[0])]
    neighbor_b = network.nodes[str(neighbors[1])]
    angle = _angle_between(node, neighbor_a, neighbor_b)
    return abs(180.0 - angle) > angle_threshold_deg


def _infer_input_node_id(network: PipeNetwork) -> str | None:
    if "INPUT" in network.nodes:
        return "INPUT"
    for node in network.nodes.values():
        if node.node_type.lower() in {"input", "supply"}:
            return node.node_id
    return None


def _select_fitting_pipe(
    network: PipeNetwork,
    node_id: str,
    depths: dict[str, int],
) -> Pipe | None:
    connected_pipes = sorted(_connected_pipes(network, node_id), key=lambda pipe: pipe.pipe_id)
    if not connected_pipes:
        return None
    if not depths:
        return connected_pipes[0]

    downstream_candidates: list[tuple[int, str, Pipe]] = []
    for pipe in connected_pipes:
        other_node = pipe.to_node if pipe.from_node == node_id else pipe.from_node
        downstream_candidates.append((depths.get(other_node, -1), pipe.pipe_id, pipe))
    return sorted(downstream_candidates, key=lambda item: (-item[0], item[1]))[0][2]


def _add_or_increment_fitting(pipe: Pipe, fitting_type: str, count: int) -> None:
    for fitting in pipe.fittings:
        if fitting.fitting_type == fitting_type:
            fitting.count += count
            return
    pipe.fittings.append(Fitting(fitting_type, count))


def _angle_between(center: Node, node_a: Node, node_b: Node) -> float:
    vector_a = (node_a.x - center.x, node_a.y - center.y, node_a.z - center.z)
    vector_b = (node_b.x - center.x, node_b.y - center.y, node_b.z - center.z)
    norm_a = _vector_norm(vector_a)
    norm_b = _vector_norm(vector_b)
    if norm_a == 0 or norm_b == 0:
        return 180.0

    dot = sum(a * b for a, b in zip(vector_a, vector_b, strict=True))
    cosine = max(-1.0, min(1.0, dot / (norm_a * norm_b)))
    return degrees(acos(cosine))


def _vector_norm(vector: tuple[float, float, float]) -> float:
    return sqrt(sum(component * component for component in vector))
