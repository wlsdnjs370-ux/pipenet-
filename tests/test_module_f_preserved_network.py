"""Source graph contracts: loops, grids, inactive heads and legacy isolation."""
from __future__ import annotations

import random

import networkx as nx
import pytest

from src.pipenet_converter.graph.flow import build_flow_tree, edge_key
from src.pipenet_converter.graph.network import FlowNetwork, network_mode


def make_network(points, edges, heads, mode="grid", **kwargs):
    tree = build_flow_tree(points, edges, [[h] for h in heads], [0], **kwargs)
    return FlowNetwork.build(tree, mode)


def test_loop_preserves_both_sides_and_prunes_pendant_ring():
    pts = [(0, 0), (1000, 0), (2000, 0), (2000, 1000), (1000, 1000),
           (3000, 1000), (1000, 2000), (0, 2000), (0, 3000), (9000, 0), (9500, 0)]
    ring = {(1, 2), (2, 3), (3, 4), (1, 4)}
    extra = {(4, 6), (6, 7), (7, 8), (6, 8), (9, 10)}
    edges = ring | extra | {(0, 1), (3, 5)}
    net = make_network(pts, edges, [5], "loop", lengths_mm={(1, 2): 1750})
    assert net.edges == ring | {(0, 1), (3, 5)}
    assert net.selected_edges([0]) == net.edges
    assert net.report()["cycle_rank"] == 1
    assert net.report()["max_load_all"] is None
    assert net.excluded[9, 10] == "unreachable"
    assert net.excluded[6, 8] == "no_head_path"
    artifact = net.extraction(pts)
    assert next(p for p in artifact["pipes"] if p["id"] == "E1-2")["plan_length_m"] == 1.75
    assert all(p["bore_m"] is None and p["length_m"] is None for p in artifact["pipes"])
    assert all(n["z_real_m"] is None for n in artifact["nodes"])
    assert edges == ring | extra | {(0, 1), (3, 5)}


def test_grid_keeps_transfer_branch_even_without_operating_head():
    # Two cross mains, three double-ended branches with nozzle stubs.
    pts = [(0, 0), (0, 1000), (0, 2000), (3000, 0), (3000, 1000), (3000, 2000),
           (1500, 0), (1500, 1000), (1500, 2000),
           (1500, -100), (1500, 900), (1500, 1900)]
    transfer = {(0, 1), (1, 2), (3, 4), (4, 5),
                (0, 6), (3, 6), (1, 7), (4, 7), (2, 8), (5, 8)}
    stubs = {(6, 9), (7, 10), (8, 11)}
    net = make_network(pts, transfer | stubs, [9, 10, 11])
    assert net.report()["cycle_rank"] == 2
    assert net.selected_edges([1]) == transfer | {(7, 10)}
    scenario = net.extraction(pts, selected_heads=[1])
    assert scenario["cycle_rank"] == 2
    assert [n["id"] for n in scenario["nozzles"] if n["selected"]] == [1]
    assert len(scenario["pipes"]) == 11


def test_unselected_inline_head_is_not_a_removed_transfer_pipe():
    pts = [(0, 0), (1000, 0), (2000, 0), (2000, 1000), (0, 1000)]
    net = make_network(pts, {(0, 1), (1, 2), (2, 3), (3, 4), (0, 4)}, [1, 3])
    out = net.extraction(pts, selected_heads=[1])
    assert {n["id"]: n["selected"] for n in out["nozzles"]} == {0: False, 1: True}
    assert len(out["pipes"]) == 5


def test_crossing_coordinates_do_not_create_connections_and_unknown_head_fails():
    pts = [(-1, 0), (1, 0), (0, -1), (0, 1)]
    net = make_network(pts, [(0, 1), (2, 3)], [1, 3])
    assert net.edges == {(0, 1)}
    assert net.excluded[2, 3] == "unreachable"
    with pytest.raises(ValueError):
        net.extraction(pts, selected_heads=[1])


def test_path_blocks_match_exhaustive_simple_paths_on_small_graphs():
    rng = random.Random(32)
    for _ in range(70):
        graph = nx.Graph()
        graph.add_nodes_from(range(7))
        graph.add_edges_from((a, b) for a in range(7) for b in range(a + 1, 7)
                             if rng.random() < .32)
        pts = [(n * 100, n * n) for n in range(7)]
        net = make_network(pts, graph.edges, [5, 6])
        expected = set()
        for target in (5, 6):
            for path in nx.all_simple_paths(graph, 0, target):
                expected.update(edge_key(a, b) for a, b in zip(path, path[1:]))
        assert net.edges == expected


def test_large_grid_no_recursion_or_exponential_path_enumeration():
    graph = nx.convert_node_labels_to_integers(nx.grid_2d_graph(50, 50))
    pts = [(n // 50 * 1000, n % 50 * 1000) for n in graph.nodes]
    net = make_network(pts, graph.edges, [2499])
    assert len(net.edges) == graph.number_of_edges()
    assert net.report()["cycle_rank"] == 49 * 49


def test_zero_length_aliases_deterministic_revisions_and_mode_validation():
    pts = [(0, 0), (0, 0), (1, 0), (0, 1)]
    edges = [(0, 1), (1, 2), (2, 3), (0, 3)]
    a = make_network(pts, edges, [2], "loop")
    b = make_network(pts, reversed(edges), [2], "loop")
    assert a.revision == b.revision and a.edges == b.edges
    assert a.revision != make_network(pts, edges, [2], "grid").revision
    with pytest.raises(ValueError):
        network_mode("auto")
