"""Regressions for a pruned continuation behind a lifted sprinkler head."""
from copy import deepcopy
from types import SimpleNamespace

import pytest

from routes.module_f.common import _boot
from routes.module_f.fitting_inspection import build_inspection
from src.pipenet_converter.graph.fittings import fitting_origins

_boot()
from services.cad_import.convert.engine import _apply_vertical
from services.cad_import.design.fitting import build_fittings
from services.cad_import.design.restrict import corridor_topology
from services.cad_import.design.tables import PipeTablesG


def expanded_head(kind="상향식"):
    """A real plan elbow upstream, and a single retained pipe ending at a head."""
    kfp = dict(nodes_meta_runtime={
        "S": dict(id="S", coords=[-1., -1., 0.], type_id="base"),
        "B": dict(id="B", coords=[-1., 0., 0.], type_id="base"),
        "H": dict(id="H", coords=[0., 0., 0.], type_id="head")},
        pipe_data={"P1": dict(start="S", end="B", length_m=1.),
                   "P2": dict(start="B", end="H", length_m=1.)})
    _apply_vertical(kfp, {"H": kind}, 0., .3, 0., .3, .3, .3, .3, .3,
                    None, None, None, None)
    return kfp


@pytest.mark.parametrize("kind", ["상향식", "하향식", "상하향식"])
def test_vertical_engine_records_head_identity_not_template_identity(kind):
    net = expanded_head(kind)
    assert len(net["head_takeoffs"]) == 1
    connection, head = next(iter(net["head_takeoffs"].items()))
    assert head == "H"
    assert net["pipe_data"]["P2"]["end"] == connection
    # Other generated bends/heads must not inherit this physical junction.
    assert set(net["head_takeoffs"].values()) == {"H"}


@pytest.mark.parametrize("has_continuation,expected", [(True, "tee"), (False, "elbow")])
@pytest.mark.parametrize("kind", ["상향식", "하향식"])
def test_original_continuation_decides_tee_not_pruned_display(has_continuation, expected, kind):
    net = expanded_head(kind)
    connection = next(iter(net["head_takeoffs"]))
    pts = [(-1000, -1000), (-1000, 0), (0, 0), (1000, 0)]
    edges = [(0, 1), (1, 2)] + ([(2, 3)] if has_continuation else [])
    original_refs = {"S": 0, "B": 1, "H": 2}
    before = deepcopy(net)
    topo = corridor_topology(dict(pts=pts, edges=edges), dict(node_ref=original_refs), net)
    assert net == before and connection not in original_refs
    assert topo["fitting_node_ref"][connection] == 2
    assert topo["phys"][connection] == (3 if has_continuation else 2)
    nodes = net["nodes_meta_runtime"]
    fits = build_fittings(net, {n: m["coords"][:2] for n, m in nodes.items()},
        {p: (25, "test") for p in net["pipe_data"]},
        node_z={n: m["coords"][2] for n, m in nodes.items()},
        parents={"B": "S", connection: "B", "H": connection}, phys=topo["phys"])
    head_pipe = next(p for p, v in net["pipe_data"].items() if v["end"] == "H")
    assert fits["per_pipe"][head_pipe]["fittings"] == [expected]
    assert fits["per_pipe"]["P2"]["fittings"] == ["elbow"]


def test_unknown_or_removed_head_does_not_guess_fitting_origin():
    origins = fitting_origins({"head": 7}, {"conn": "head", "bend": "unknown"},
                              {"head", "conn", "bend"})
    assert origins["conn"].source_vertex == 7 and origins["conn"].added_ports == 1
    assert origins["head"].added_ports == 0 and "bend" not in origins
    assert not fitting_origins({"head": 7}, {"conn": "head"}, {"conn"})


def test_head_takeoff_card_and_glyph_keep_omitted_through_arm():
    nodes = [dict(label="1", x=-866., y=-500.), dict(label="2", x=0., y=0.),
             dict(label="3", x=0., y=300.)]
    tbl = PipeTablesG(nodes=nodes, pipes=[
        dict(label="P1", **{"in": "1", "out": "2"}, dia=25),
        dict(label="P2", **{"in": "2", "out": "3"}, dia=25, eq_len=1.5)],
        fittings=[dict(pipe="P2", type="tee", count=1)])
    got = dict(node_ref={"head": 1}, fitting_node_ref={"conn": 1}, phys={"conn": 3})
    board = SimpleNamespace(pts=[(-1000, 0), (0, 0), (1000, 0)], edges=[(0, 1), (1, 2)])
    rows = build_inspection(tbl, got, dict(nid={"2": "conn", "3": "head"}), nodes,
        board=board, transform=dict(iso=True, cos30=3**.5/2, sin30=.5, k=1))["fittings"]
    row = rows[0]
    assert row["source_node"] == 1 and row["original_degree"] == 3
    assert row["current_degree"] == 2 and len(row["arms"]) == 3
    assert not row["symbolic"] and row["kind"] == "tee"
    assert "TEE_BRANCH" in row["eq_source"] and row["eq_m"] == 1.5
    assert [0., 1.] in row["arms"]
    assert any(arm[0] > .8 and arm[1] > .4 for arm in row["arms"])
