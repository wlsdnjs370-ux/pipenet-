"""Tests for 2D graph building from raw DXF entities."""

from pipenet_converter.dxf_extractor import RawBlock, RawSegment, RawText
from pipenet_converter.graph_builder import (
    assign_diameters_to_edges,
    attach_blocks_to_graph,
    build_2d_graph_from_segments,
    graph_to_pipenetwork_2d,
    snap_points,
)
import pytest


def test_nearly_touching_endpoints_snap_together() -> None:
    points = [(0.0, 0.0), (10.0, 0.0), (1000.0, 0.0)]

    snapped = snap_points(points, tolerance=50.0)

    assert snapped[(0.0, 0.0)] == (5.0, 0.0)
    assert snapped[(10.0, 0.0)] == (5.0, 0.0)
    assert snapped[(1000.0, 0.0)] == (1000.0, 0.0)


def test_simple_t_shape_creates_degree_three_node() -> None:
    segments = [
        RawSegment("S1", (-1000.0, 0.0), (0.0, 0.0), "F-SP-PIPE", "pipe"),
        RawSegment("S2", (0.0, 0.0), (1000.0, 0.0), "F-SP-PIPE", "pipe"),
        RawSegment("S3", (0.0, 0.0), (0.0, 1000.0), "F-SP-PIPE", "pipe"),
    ]

    graph = build_2d_graph_from_segments(segments, snap_tolerance=50.0)

    assert graph.number_of_nodes() == 4
    assert graph.number_of_edges() == 3
    assert sorted(dict(graph.degree()).values()) == [1, 1, 1, 3]


def test_head_block_attaches_to_nearest_pipe_node() -> None:
    graph = build_2d_graph_from_segments(
        [RawSegment("S1", (0.0, 0.0), (1000.0, 0.0), "F-SP-PIPE", "pipe")],
        snap_tolerance=50.0,
    )
    block = RawBlock("B1", "SP_HEAD", (40.0, 0.0), "F-SP-HEAD", "head", rotation=None)

    attachments = attach_blocks_to_graph(graph, [block], max_distance=600.0)

    attached_node = attachments["B1"]
    assert graph.nodes[attached_node]["node_type"] == "head"
    assert graph.nodes[attached_node]["head_blocks"][0] == block
    assert graph.number_of_nodes() == 2


def test_far_head_block_creates_connected_head_node() -> None:
    graph = build_2d_graph_from_segments(
        [RawSegment("S1", (0.0, 0.0), (1000.0, 0.0), "F-SP-PIPE", "pipe")],
        snap_tolerance=50.0,
    )
    block = RawBlock("B1", "SP_HEAD", (3000.0, 0.0), "F-SP-HEAD", "head", rotation=None)

    attachments = attach_blocks_to_graph(graph, [block], max_distance=600.0)

    assert attachments["B1"] == "H_B1"
    assert graph.nodes["H_B1"]["node_type"] == "head"
    assert graph.has_edge("N000002", "H_B1")


def test_diameter_text_assigns_diameter_to_edges() -> None:
    graph = build_2d_graph_from_segments(
        [RawSegment("S1", (0.0, 0.0), (1000.0, 0.0), "F-SP-PIPE", "pipe")],
        snap_tolerance=50.0,
    )
    diameter_text = RawText("T1", "150A", (500.0, 100.0), "F-SP-TEXT", "diameter_text")

    assign_diameters_to_edges(graph, [diameter_text], default_diameter_m=0.032)

    edge_data = next(iter(graph.edges(data=True)))[2]
    assert edge_data["diameter_m"] == 0.15


def test_invalid_diameter_label_raises_clear_error() -> None:
    graph = build_2d_graph_from_segments(
        [RawSegment("S1", (0.0, 0.0), (1000.0, 0.0), "F-SP-PIPE", "pipe")],
        snap_tolerance=50.0,
    )
    diameter_text = RawText("T_BAD", "300A", (500.0, 0.0), "F-SP-TEXT", "diameter_text")

    with pytest.raises(ValueError, match="Invalid diameter label '300A' in raw text 'T_BAD'"):
        assign_diameters_to_edges(graph, [diameter_text], default_diameter_m=0.032)


def test_graph_to_pipenetwork_2d_creates_pipes_and_nozzles() -> None:
    graph = build_2d_graph_from_segments(
        [
            RawSegment("S1", (0.0, 0.0), (1000.0, 0.0), "F-SP-PIPE", "pipe"),
            RawSegment("S2", (1000.0, 0.0), (1000.0, 1000.0), "F-SP-PIPE", "pipe"),
        ],
        snap_tolerance=50.0,
    )
    attach_blocks_to_graph(
        graph,
        [RawBlock("B1", "SP_HEAD", (1000.0, 1000.0), "F-SP-HEAD", "head", rotation=None)],
        max_distance=600.0,
    )
    assign_diameters_to_edges(
        graph,
        [RawText("T1", "50A", (500.0, 0.0), "F-SP-TEXT", "diameter_text")],
        default_diameter_m=0.032,
    )

    network = graph_to_pipenetwork_2d(graph, title="2D network")

    assert network.title == "2D network"
    assert len(network.pipes) == 2
    assert len(network.nozzles) == 1
    assert len(network.nodes) == 4
    assert {pipe.diameter_m for pipe in network.pipes.values()} == {0.05}
    assert next(iter(network.nozzles.values())).output_node == "@/1"
