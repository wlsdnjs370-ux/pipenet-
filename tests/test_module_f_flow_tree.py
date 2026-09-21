"""Canonical routes, source preservation, and all-head load invariants."""
from __future__ import annotations

import copy
from pathlib import Path
import sys
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
for path in (ROOT, ROOT / "cad_project_editor_g", ROOT / "core"):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from src.pipenet_converter.graph.flow import build_flow_tree
from services.cad_import.design.flow import flow_for_board, physical_pipe_loads
from services.cad_import.design.worst import worst_k_heads


def test_cycle_dead_end_unreachable_and_single_route():
    pts = [(0, 0), (1000, 0), (1000, 1000), (0, 1000), (2000, 0), (5000, 0), (6000, 0)]
    edges = [(0, 1), (1, 2), (2, 3), (3, 0), (1, 4), (5, 6)]
    original = copy.deepcopy((pts, edges))
    flow = build_flow_tree(pts, edges, [{2}, {5}], [0])
    assert flow.path(2) == [0, 1, 2]
    assert flow.edges == {(0, 1), (1, 2)}
    assert flow.excluded[(1, 4)] == "no_head"
    assert flow.excluded[(2, 3)] == "alternate_route"
    assert flow.excluded[(5, 6)] == "unreachable"
    assert (pts, edges) == original
    assert flow.revision == build_flow_tree(pts, reversed(edges), [{2}, {5}], [0]).revision


def test_zero_length_cycle_is_acyclic_and_stable():
    flow = build_flow_tree([(0, 0)] * 3 + [(10, 0)], [(0, 1), (1, 2), (0, 2), (2, 3)], [{3}], [0])
    assert flow.path(3) == [0, 2, 3]
    assert flow.distance_mm[3] == 10


def test_inferred_cost_is_lexicographic_not_added_length():
    flow = build_flow_tree([(0, 0), (100, 0), (10, 0)], [(0, 1), (1, 2), (0, 2)], [{2}], [0], inferred_edges=[(0, 2)])
    assert flow.path(2) == [0, 1, 2]
    assert flow.distance_mm[2] == 190


def test_declared_length_and_invalid_length():
    flow = build_flow_tree([(0, 0), (100, 0)], [(0, 1)], [{1}], [0], lengths_mm={(0, 1): 2500})
    assert flow.report()["length_m"] == 2.5
    with pytest.raises(ValueError):
        build_flow_tree([(0, 0), (100, 0)], [(0, 1)], [{1}], [0], lengths_mm={(0, 1): -1})


def board_fixture():
    return SimpleNamespace(pts=[(0, 0), (1000, 0), (2000, 0), (1000, 2000)],
        edges={(0, 1), (1, 2), (1, 3)}, hnodes=[{2}, {3}], head_centers=[2, 3],
        disks=[(2000, 0, 50), (1000, 2000, 50)], sources=[0], edge_len_mm={})


def test_worst_zone_and_k_never_rebuild_or_reduce_physical_loads():
    b = board_fixture()
    flow = flow_for_board(b)
    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, 1, only_heads={0},
                     head_xy=b.disks, flow_tree=flow)
    assert w["edges"] == {(0, 1), (1, 2)}
    assert w["loads"][(0, 1)] == 1
    assert w["physical_loads"][(0, 1)] == 2
    assert w["flow_revision"] == flow.revision
    assert flow_for_board(b) is flow
    b.pts[2] = (2200, 0)
    assert flow_for_board(b) is not flow
    changed = flow_for_board(b)
    b.edge_len_mm[(0, 1)] = 9000
    assert flow_for_board(b).revision != changed.revision


def test_requires_one_source_and_rejects_wrong_basis():
    b = board_fixture()
    b.sources.append(3)
    with pytest.raises(ValueError):
        flow_for_board(b)
    f = flow_for_board(b, selected_source="Z2")
    assert f.roots == (3,)
    with pytest.raises(ValueError):
        worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, 1, source_index=0, flow_tree=f)


def test_full_load_through_synthetic_riser_is_not_selected_load():
    net = {"nodes_meta_runtime": {"s": {"io_node": "Input", "type_id": "pump"}, "r": {},
                                 "t": {}, "h": {"type_id": "head"}},
           "pipe_data": {"rise": {"start": "s", "end": "r"},
                         "main": {"start": "r", "end": "t", "head_count_all": 12},
                         "drop": {"start": "t", "end": "h"}}}
    assert physical_pipe_loads(net) == {"drop": 1, "main": 12, "rise": 12}
    net["pipe_data"]["loop"] = {"start": "s", "end": "t"}
    with pytest.raises(ValueError):
        physical_pipe_loads(net)


def test_bore_uses_full_capacity_even_when_reference_exists():
    from services.cad_import.design.bore import decide_bores, nfpc_min_bore_mm
    net = {"pipe_data": {"P1": {}}, "physical_pipe_loads": {"P1": 12}}
    bores = decide_bores(net, {"P1": (0, 1)}, {(0, 1): 1}, [], tree_loads={"P1": 1})
    assert bores["P1"][0] == nfpc_min_bore_mm(12)
    assert bores["P1"][0] > nfpc_min_bore_mm(1)


def test_same_place_alias_with_zone_keeps_global_representative_route():
    pts = [(0, 0), (1000, 0), (0, 1000), (2000, 0), (2000, 0)]
    edges = [(0, 1), (1, 3), (0, 2), (2, 4)]
    hnodes = [{3}, {4}]
    xy = [(2000, 0), (2000, 0)]
    f = build_flow_tree(pts, edges, hnodes, [0], head_xy=xy)
    w = worst_k_heads(pts, edges, hnodes, [0], 1, only_heads={1}, head_xy=xy, flow_tree=f)
    assert w["heads"] == [1]
    assert w["edges"] == set(f.selected_loads([1])) == {(0, 1), (1, 3)}
    from services.cad_import.design.restrict import restrict_to_worst
    b = SimpleNamespace(pts=pts, edges=edges, hnodes=hnodes, sources=[0],
                        disks=[(2000, 0, 30)] * 2, head_centers=[3, 4], disk_kinds=["상향식"] * 2)
    limited = restrict_to_worst({"disk_kinds": b.disk_kinds}, b, {"heads": [1]})
    assert limited["flow_head_nodes"] == [3]
    assert {tuple(e) for e in limited["edges"]} == {(0, 1), (1, 3)}
