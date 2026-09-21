"""Regression: pipe provenance, actual junction identity and no guessed crosses."""
from types import SimpleNamespace

from routes.module_f.common import _boot

_boot()
from services.cad_import.pipeline.stage1 import material_bundles_v2
from services.cad_import.design.restrict import corridor_topology
from services.cad_import.design.fitting import build_fittings


def test_unpicked_layer_with_same_color_is_not_pipe_material():
    world = SimpleNamespace(segs=[(ly, 3, (0, 0), (1000, 0))
                                  for ly in ("pipe", "SHEET", "equipment")])
    spec = {"material_picks": [["pipe", 3]]}
    assert material_bundles_v2(world, spec)[0] == {("pipe", 3)}
    spec["material_scope"] = "color"
    assert len(material_bundles_v2(world, spec)[0]) == 3


def test_explicit_additional_pipe_layer_and_other_fire_separation():
    world = SimpleNamespace(segs=[("pipe", 3, (), ()), ("riser", 3, (), ()),
                                  ("hydrant", 3, (), ())])
    spec = {"material_picks": [["pipe", 3], ["riser", 3]],
            "other_fire": [["hydrant", 3]]}
    assert material_bundles_v2(world, spec)[0] == {("pipe", 3), ("riser", 3)}


def test_nearby_disconnected_elbows_do_not_become_tee_or_cross():
    # Vertices 1 and 4 share a former 50mm display cell, but are not connected.
    pts = [(-1000, 0), (0, 0), (0, 1000), (-1000, 10), (10, 10), (10, 1000)]
    limited = dict(pts=pts, edges=[(0, 1), (1, 2), (3, 4), (4, 5), (1, 0)])
    kfp = dict(nodes_meta_runtime={"n": {}}, pipe_data={
        "a": dict(start="s", end="n"), "b": dict(start="n", end="h")})
    topo = corridor_topology(limited, dict(node_ref={"n": 1},
        edge_ref={"a": [0, 1], "b": [1, 2]}, origin_mm=[0, 0]), kfp)
    assert topo["phys"]["n"] == 2
    fits = build_fittings(kfp, {"s": (-1, 0), "n": (0, 0), "h": (0, 1)},
                          {"a": (25, "test"), "b": (25, "test")},
                          parents={"n": "s", "h": "n"}, phys=topo["phys"])
    assert fits["per_pipe"]["b"]["fittings"] == ["elbow"]


def test_four_port_junction_is_unresolved_not_guessed_cross_or_two_tees():
    net = {"pipe_data": {"a": {"start": "s", "end": "n"},
                         "b": {"start": "n", "end": "h"}}}
    result = build_fittings(net, {"s": (-1, 0), "n": (0, 0), "h": (0, 1)},
                            {"b": (25, "test")}, parents={"n": "s"}, phys={"n": 4})
    assert result["per_pipe"]["b"]["fittings"] == []
    assert result["unresolved_kind"] == 1
    assert result["unresolved_kind_items"][0]["ports"] == 4


def test_upstream_unknown_does_not_assign_branch_tee_loss():
    net = {"pipe_data": {"a": {"start": "s", "end": "n"},
                         "b": {"start": "n", "end": "h"}}}
    result = build_fittings(net, {"s": (-1, 0), "n": (0, 0), "h": (0, 1)},
                            {"b": (25, "test")}, phys={"n": 3})
    assert result["counts"] == {} and result["unresolved_kind"] == 1


def test_hidden_four_port_site_is_not_a_tee_run():
    pts = [(0, 0), (1000, 0), (2000, 0), (1000, 1000), (1000, -1000)]
    topo = corridor_topology(dict(pts=pts, edges=[(0,1),(1,2),(1,3),(1,4)]),
        dict(node_ref={"s": 0, "h": 2}, edge_ref={"p": [0,2]}),
        dict(nodes_meta_runtime={"s": {}, "h": {}},
             pipe_data={"p": dict(start="s",end="h")}))
    assert topo["interior_junctions"] == {}
    assert topo["interior_unresolved"] == {"p": 1}


def test_vertical_through_path_is_run_not_branch_tee():
    net = dict(pipe_data={"a": dict(start="s",end="n"),
                          "b": dict(start="n",end="h"),
                          "c": dict(start="n",end="side")})
    result = build_fittings(net, {"s": (0,0), "n": (0,0), "h": (0,0), "side": (1,0)},
        {p: (25,"test") for p in net["pipe_data"]},
        parents={"n": "s", "h": "n", "side": "n"},
        node_z={"s": 0,"n": 1,"h": 2,"side": 1}, phys={"n": 3})
    assert result["per_pipe"]["b"]["fittings"] == ["tee-run"]
    assert result["per_pipe"]["c"]["fittings"] == ["tee"]


def test_elevated_branch_replaces_original_port_not_a_fourth_port():
    original = dict(pts=[(-1000,0),(0,0),(1000,0),(0,1000)],
                    edges=[(0,1),(1,2),(1,3)])
    kfp = dict(nodes_meta_runtime={"n": {}},pipe_data={
        "in": dict(start="s",end="n"), "run": dict(start="n",end="h"),
        "rise": dict(start="n",end="up")})
    topo = corridor_topology(original, dict(node_ref={"n":1},
        edge_ref={"in":[0,1],"run":[1,2]}), kfp)
    assert topo["phys"]["n"] == 3


def test_missing_edge_reference_is_not_evidence_of_an_extra_vertical_port():
    topo = corridor_topology(dict(pts=[(0,0),(1000,0),(1000,1000)],edges=[(0,1),(1,2)]),
        dict(node_ref={"n":1},edge_ref={}),
        dict(nodes_meta_runtime={"n": {}},pipe_data={
            "a":dict(start="s",end="n"),"b":dict(start="n",end="h")}))
    assert topo["phys"]["n"] == 2


def test_coincident_cad_vertices_do_not_double_elbow_ports():
    # Real-drawing regression: each elbow arm had two distinct vertex IDs at
    # exactly the same XY. Four edges still describe only two physical ports.
    topo = corridor_topology(dict(pts=[(0,0),(-100,0),(0,100),(-100,0),(0,100)],
                                   edges=[(0,1),(0,2),(0,3),(0,4)]),
        dict(node_ref={"n":0}),dict(nodes_meta_runtime={"n":{}},pipe_data={}))
    assert topo["phys"]["n"] == 2
