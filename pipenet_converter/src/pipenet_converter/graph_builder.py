"""Pipe network graph construction.

This module converts raw DXF-derived pipe centerline geometry into a simple 2D
graph, then into a preliminary ``PipeNetwork``. Version 1 intentionally handles
clean CAD input and leaves advanced edge cases, such as arbitrary line
intersection splitting, for later stages.
"""

from __future__ import annotations

from dataclasses import dataclass
from math import dist
from typing import Iterable

import networkx as nx

from pipenet_converter.dxf_extractor import RawBlock, RawSegment, RawText
from pipenet_converter.models import Node, Nozzle, Pipe, PipeNetwork, diameter_label_to_m


@dataclass(slots=True)
class GraphBuildOptions:
    """Options for converting raw CAD geometry into a pipe graph."""

    snap_tolerance: float = 50.0
    head_attach_tolerance: float = 600.0
    cad_unit_scale_to_m: float = 0.001
    default_diameter_m: float = 0.032
    default_c_factor: float = 120.0


DEFAULT_GRAPH_BUILD_OPTIONS = GraphBuildOptions()


def build_2d_graph_from_segments(
    segments: Iterable[RawSegment],
    snap_tolerance: float,
) -> nx.Graph:
    """Build an undirected 2D graph from raw pipe segments."""
    segment_list = list(segments)
    points = [point for segment in segment_list for point in (segment.start, segment.end)]
    snapped_points = snap_points(points, snap_tolerance)
    graph = nx.Graph()
    point_to_node_id: dict[tuple[float, float], str] = {}

    for snapped_point in dict.fromkeys(snapped_points.values()):
        node_id = f"N{len(point_to_node_id) + 1:06d}"
        point_to_node_id[snapped_point] = node_id
        graph.add_node(
            node_id,
            x=snapped_point[0],
            y=snapped_point[1],
            node_type="pipe",
            head_blocks=[],
        )

    for segment in segment_list:
        start = snapped_points[segment.start]
        end = snapped_points[segment.end]
        start_node = point_to_node_id[start]
        end_node = point_to_node_id[end]
        if start_node == end_node:
            continue

        length_cad = dist(start, end)
        graph.add_edge(
            start_node,
            end_node,
            segment_id=segment.segment_id,
            layer=segment.layer,
            semantic_class=segment.semantic_class,
            length_cad=length_cad,
            length_m=length_cad * DEFAULT_GRAPH_BUILD_OPTIONS.cad_unit_scale_to_m,
        )

    return graph


def snap_points(points: Iterable[tuple[float, float]], tolerance: float) -> dict[tuple[float, float], tuple[float, float]]:
    """Snap points within ``tolerance`` to a shared centroid coordinate."""
    snapped: dict[tuple[float, float], tuple[float, float]] = {}
    clusters: list[list[tuple[float, float]]] = []

    for point in points:
        cluster = _find_cluster(point, clusters, tolerance)
        if cluster is None:
            clusters.append([point])
        else:
            cluster.append(point)

    for cluster in clusters:
        centroid = (
            sum(point[0] for point in cluster) / len(cluster),
            sum(point[1] for point in cluster) / len(cluster),
        )
        for point in cluster:
            snapped[point] = centroid

    return snapped


def attach_blocks_to_graph(
    graph: nx.Graph,
    blocks: Iterable[RawBlock],
    max_distance: float,
) -> dict[str, str]:
    """Attach head blocks to the nearest graph node or create connected head nodes."""
    attachments: dict[str, str] = {}
    for block in blocks:
        if block.semantic_class != "head":
            continue

        nearest_node_id, nearest_distance = _nearest_node(graph, block.insert)
        if nearest_node_id is None:
            continue

        if nearest_distance <= max_distance:
            target_node_id = nearest_node_id
            _mark_head_node(graph, target_node_id, block)
        else:
            target_node_id = f"H_{block.block_id}"
            graph.add_node(
                target_node_id,
                x=block.insert[0],
                y=block.insert[1],
                node_type="head",
                head_blocks=[],
            )
            _mark_head_node(graph, target_node_id, block)
            graph.add_edge(
                nearest_node_id,
                target_node_id,
                segment_id=f"HEAD_{block.block_id}",
                layer=block.layer,
                semantic_class="head_connection",
                length_cad=nearest_distance,
                length_m=nearest_distance * DEFAULT_GRAPH_BUILD_OPTIONS.cad_unit_scale_to_m,
            )

        attachments[block.block_id] = target_node_id

    return attachments


def assign_diameters_to_edges(
    graph: nx.Graph,
    diameter_texts: Iterable[RawText],
    default_diameter_m: float,
) -> None:
    """Assign each graph edge a diameter using the nearest diameter text."""
    text_list = list(diameter_texts)
    for node_a, node_b, data in graph.edges(data=True):
        nearest_text = _nearest_text_to_edge(graph, node_a, node_b, text_list)
        if nearest_text is None:
            diameter_m = default_diameter_m
        else:
            try:
                diameter_m = diameter_label_to_m(nearest_text.text)
            except ValueError as exc:
                raise ValueError(
                    f"Invalid diameter label {nearest_text.text!r} in raw text "
                    f"{nearest_text.text_id!r}."
                ) from exc
        data["diameter_m"] = diameter_m


def graph_to_pipenetwork_2d(graph: nx.Graph, title: str) -> PipeNetwork:
    """Convert a 2D graph into a preliminary ``PipeNetwork`` with z=0."""
    network = PipeNetwork(title=title)
    node_id_map: dict[str, str] = {}

    for index, (graph_node_id, data) in enumerate(graph.nodes(data=True), start=1):
        pipe_node_id = str(graph_node_id)
        node_id_map[str(graph_node_id)] = pipe_node_id
        network.add_node(
            Node(
                node_id=pipe_node_id,
                x=float(data["x"]),
                y=float(data["y"]),
                z=0.0,
                node_type=str(data.get("node_type", "pipe")),
                source="graph_2d",
                metadata={"graph_index": index},
            )
        )

    for index, (node_a, node_b, data) in enumerate(graph.edges(data=True), start=1):
        network.add_pipe(
            Pipe(
                pipe_id=f"P{index:06d}",
                from_node=node_id_map[str(node_a)],
                to_node=node_id_map[str(node_b)],
                diameter_m=float(data.get("diameter_m", DEFAULT_GRAPH_BUILD_OPTIONS.default_diameter_m)),
                length_m=float(
                    data.get(
                        "length_m",
                        data.get("length_cad", 0.0)
                        * DEFAULT_GRAPH_BUILD_OPTIONS.cad_unit_scale_to_m,
                    )
                ),
                rise_m=0.0,
                c_factor=DEFAULT_GRAPH_BUILD_OPTIONS.default_c_factor,
                metadata={
                    "segment_id": str(data.get("segment_id", "")),
                    "layer": str(data.get("layer", "")),
                    "semantic_class": str(data.get("semantic_class", "")),
                },
            )
        )

    nozzle_index = 1
    for graph_node_id, data in graph.nodes(data=True):
        for block in data.get("head_blocks", []):
            nozzle_output_node_id = f"@/{nozzle_index}"
            network.add_node(
                Node(
                    node_id=nozzle_output_node_id,
                    x=float(data["x"]),
                    y=float(data["y"]),
                    z=0.0,
                    node_type="head_output",
                    source="graph_2d",
                    metadata={"block_id": block.block_id},
                )
            )
            network.add_nozzle(
                Nozzle(
                    nozzle_id=str(nozzle_index),
                    input_node=node_id_map[str(graph_node_id)],
                    output_node=nozzle_output_node_id,
                    flow_m3s=0.00266666667,
                    status=1,
                    metadata={"block_id": block.block_id, "block_name": block.block_name},
                )
            )
            nozzle_index += 1

    return network


def _find_cluster(
    point: tuple[float, float],
    clusters: list[list[tuple[float, float]]],
    tolerance: float,
) -> list[tuple[float, float]] | None:
    for cluster in clusters:
        centroid = (
            sum(cluster_point[0] for cluster_point in cluster) / len(cluster),
            sum(cluster_point[1] for cluster_point in cluster) / len(cluster),
        )
        if dist(point, centroid) <= tolerance:
            return cluster
    return None


def _nearest_node(graph: nx.Graph, point: tuple[float, float]) -> tuple[str | None, float]:
    nearest_node_id: str | None = None
    nearest_distance = float("inf")
    for node_id, data in graph.nodes(data=True):
        node_distance = dist(point, (float(data["x"]), float(data["y"])))
        if node_distance < nearest_distance:
            nearest_node_id = str(node_id)
            nearest_distance = node_distance
    return nearest_node_id, nearest_distance


def _mark_head_node(graph: nx.Graph, node_id: str, block: RawBlock) -> None:
    graph.nodes[node_id]["node_type"] = "head"
    graph.nodes[node_id].setdefault("head_blocks", []).append(block)


def _nearest_text_to_edge(
    graph: nx.Graph,
    node_a: str,
    node_b: str,
    diameter_texts: list[RawText],
) -> RawText | None:
    if not diameter_texts:
        return None

    data_a = graph.nodes[node_a]
    data_b = graph.nodes[node_b]
    midpoint = (
        (float(data_a["x"]) + float(data_b["x"])) / 2,
        (float(data_a["y"]) + float(data_b["y"])) / 2,
    )
    return min(diameter_texts, key=lambda text: dist(midpoint, text.insert))
