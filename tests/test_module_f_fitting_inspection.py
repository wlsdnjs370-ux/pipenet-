"""Read-only inspection preserves physical tee identity and explicit loss sources."""
from copy import deepcopy
from dataclasses import asdict
from types import SimpleNamespace

import pytest
from flask import Flask

from routes.module_f import api_design, jobs
from routes.module_f.common import _boot
from routes.module_f.fitting_inspection import build_inspection


def fixture():
    _boot()
    from services.cad_import.design.tables import PipeTablesG
    board = SimpleNamespace(pts=[(0, 0), (1000, 0), (2000, 0), (1000, 1000), (2000, 1000)],
                            edges=[(0, 1), (1, 2), (1, 3), (3, 4)])
    nodes = [dict(label=str(i+1), x=x, y=y, elevation=0, io_node="Input" if i==0 else "No")
             for i, (x, y) in enumerate([board.pts[i] for i in (0, 1, 3, 4)])]
    pipes = [dict(label=f"P{i+1}", **{"in": str(i+1), "out": str(i+2)},
                  length=1, dia=25, type="KSD 3507", c=120, eq_len=eq)
             for i, eq in enumerate((0, 1.5, .6))]
    tbl = PipeTablesG(nodes=nodes, pipes=pipes,
        fittings=[dict(pipe="P2", **{"in":"2", "out":"3"}, type="tee", count="1"),
                  dict(pipe="P3", **{"in":"3", "out":"4"}, type="elbow", count="1")],
        pipe_labels={f"KP{i}":f"P{i}" for i in (1, 2, 3)})
    got = dict(node_ref={"N1":0, "N2":1, "N3":3, "N4":4},
               phys={"N1":1, "N2":3, "N3":2, "N4":1}, origin_mm=[0, 0],
               edge_ref={"KP1":[0,1], "KP2":[1,3], "KP3":[3,4]})
    got["kfp"] = dict(nodes_meta_runtime={f"N{n['label']}":dict(coords=[n['x']/1000,n['y']/1000,0],
        type_id="base",elevation_m=0) for n in nodes}, pipe_data={f"KP{i+1}":dict(start="N"+p["in"],end="N"+p["out"],length_m=1) for i,p in enumerate(pipes)})
    keys = dict(nid={str(i):f"N{i}" for i in (1,2,3,4)},node={},pipe={})
    return tbl, got, keys, board


def inspect(t, g, k, b, **kw):
    return build_inspection(t,g,k,t.nodes,board=b,transform={"iso":False},**kw)["fittings"]


def test_pruned_tee_is_not_elbow_and_has_three_original_arms():
    t,g,k,b=fixture(); before=deepcopy((asdict(t),g,k,b))
    tee,elbow=inspect(t,g,k,b)
    assert (tee["kind"],tee["shape"],tee["node"]) == ("tee","tee","2")
    assert (tee["original_degree"],tee["current_degree"]) == (3,2)
    assert tee["flow_path"] == ["1", "2", "3"]
    assert len(tee["arms"]) == 3 and tee["symbolic"] is False
    assert elbow["node"] == "3" and elbow["shape"] == "elbow"
    assert "TEE_BRANCH" in tee["eq_source"] and tee["eq_m"] > 0
    assert (asdict(t),g,k,b) == before


def test_tee_run_has_no_glyph_or_properties():
    t,g,k,b=fixture();t.fittings[0]["type"]="tee-run"
    rows=inspect(t,g,k,b)
    assert len(rows)==1 and rows[0]["kind"]=="elbow"
    t.pipes[2]["dia"]=999
    assert inspect(t,g,k,b)[0]["eq_m"] is None


def test_interior_straight_tees_are_absent_even_with_original_branch_evidence():
    t,g,k,b=fixture()
    t.nodes=[dict(label="1",x=0,y=0),dict(label="2",x=2000,y=0)]
    t.pipes=[dict(label="P1",**{"in":"1","out":"2"},dia=25,type="KSD 3507",eq_len=0)]
    t.fittings=[dict(pipe="P1",**{"in":"1","out":"2"},type="tee-run",count=1)]
    g.update(node_ref={"N1":0,"N2":2},edge_ref={"KP1":[0,2]},interior_junctions={"KP1":1})
    assert inspect(t,g,k,b)==[]
    assert inspect(t,g,k,None)==[]


def test_manual_equipment_and_unresolved_locations_remain_visible():
    t,g,k,b=fixture()
    t.equipment=[dict(pipe="P1",desc="45° 엘보",editor_library="ELBOW_45",rel_pos=.4,count=2,eq_len=.6)]
    t.unresolved={"kind_items":[dict(pipe="KP1",pipe_label="P1",node="N1",node_label="1",n=1)]}
    rows=inspect(t,g,k,b)
    assert len(rows)==4
    manual=rows[2]
    assert manual["x"]==400 and manual["eq_m"]==pytest.approx(.3)
    assert manual["shape"]=="elbow" and manual["symbolic"]
    assert rows[3]["kind"]=="unresolved" and rows[3]["node"]=="1"


def test_edited_graph_keeps_type_without_claiming_old_branch_direction():
    t,g,k,b=fixture()
    row=inspect(t,g,k,b,edited=True)[0]
    assert row["kind"]=="tee" and row["symbolic"] and len(row["arms"])==2


def test_unrelated_edit_does_not_discard_original_local_arms():
    t,g,k,b=fixture()
    row=inspect(t,g,k,b,edited=True,base_got=deepcopy(g))[0]
    assert len(row["arms"]) == 3 and not row["symbolic"]
    assert row["glyph_radius"] == 200


@pytest.mark.parametrize("iso",["0","1"])
def test_api_annotations_do_not_change_calculation_tables(tmp_path,iso):
    t,g,k,b=fixture()
    sess=jobs._new_session(key="fitting-inspector-fixture")
    sess.update(design=dict(tables=t,got=g,keys=k),edit=SimpleNamespace(board=b))
    before=deepcopy((asdict(t),g))
    app=Flask(__name__);app.testing=True
    api_design.register(app,UPLOAD_DIR=tmp_path)
    try:
        r=app.test_client().get('/api/module-f/design/preview',query_string=dict(sid=sess['id'],iso=iso))
        assert r.status_code==200,r.json
        rows=r.json['view']['inspection']['fittings']
        assert [r['shape'] for r in rows]==['tee','elbow']
        n=next(n for n in r.json['view']['nodes'] if n['label']=='2')
        assert (rows[0]['x'],rows[0]['y'])==(n['x'],n['y'])
        assert (asdict(t),g)==before
    finally:
        jobs._SESSIONS.pop(sess['id'],None)


def bracket_top_fixture(partner_mm=0.04):
    """괄호 교차 꼭대기(오너 2026-09-22) — 겹친 노드가 0.5 m 세로관으로 선 자리.

    board: 주배관 M0─Mm─M2, Mm 에서 ``partner_mm`` 떨어진 가지 노드 O 가 북쪽(N)·
    남쪽(S) 반쪽을 잇는다. 계산망은 남쪽 반쪽을 가지치기로 잘랐다 — 꼭대기 3 에는
    세로관 P2 와 북쪽 팔 P3 만 남는다. 부속표는 둘 다 분류티다.
    """
    import math
    _boot()
    from services.cad_import.design.tables import PipeTablesG
    ang = math.radians(167.0)
    pts = [(0.0, 0.0), (1000.0, 0.0), (2000.0, 0.0),
           (1000.0 + partner_mm * math.cos(ang), partner_mm * math.sin(ang)),
           (1000.0, 1000.0), (1000.0, -1000.0)]
    board = SimpleNamespace(pts=pts, edges=[(0, 1), (1, 2), (1, 3), (3, 4), (3, 5)])
    plan = {"1": (0, 0, 0.0), "2": (1000, 0, 0.0), "3": (1000, 0, 0.5),
            "4": (1000, 1000, 0.5), "5": (2000, 0, 0.0)}
    c30, s30, lift = math.cos(math.radians(30)), 0.5, 1000.0
    nodes = [dict(label=l, x=x, y=y, elevation=e, io_node="Input" if l == "1" else "No")
             for l, (x, y, e) in plan.items()]
    shown = [dict(label=l, x=(x - y) * c30, y=(x + y) * s30 + e * lift)
             for l, (x, y, e) in plan.items()]
    pipes = [dict(label=pl, length=ln, dia=40, type="KSD 3507", c=120, eq_len=eq, **{"in": a, "out": b})
             for pl, a, b, ln, eq in (("P1", "1", "2", 1.0, 0.0), ("P2", "2", "3", 0.5, 2.4),
                                      ("P3", "3", "4", 1.0, 2.4), ("P4", "2", "5", 1.0, 0.0))]
    tbl = PipeTablesG(nodes=nodes, pipes=pipes,
        fittings=[dict(pipe="P2", type="tee", count="1", **{"in": "2", "out": "3"}),
                  dict(pipe="P3", type="tee", count="1", **{"in": "3", "out": "4"}),
                  dict(pipe="P4", type="tee-run", count="1", **{"in": "2", "out": "5"})],
        pipe_labels={f"KP{i}": f"P{i}" for i in (1, 2, 3, 4)})
    got = dict(node_ref={"N1": 0, "N2": 1, "N3": 3, "N4": 4, "N5": 2},
               phys={"N1": 1, "N2": 3, "N3": 3, "N4": 1, "N5": 1})
    keys = dict(nid={l: f"N{l}" for l in plan}, node={}, pipe={})
    xf = {"iso": True, "cos30": c30, "sin30": s30, "k": 1.0}
    return tbl, got, keys, board, shown, xf


def test_bracket_top_tee_draws_the_pruned_half_not_the_stood_up_joint():
    t, g, k, b, shown, xf = bracket_top_fixture()
    before = deepcopy((asdict(t), g, k))
    rows = build_inspection(t, g, k, shown, board=b, transform=xf)["fittings"]
    top = next(r for r in rows if r["node"] == "3")
    bottom = next(r for r in rows if r["node"] == "2")
    # 부속 종류·등가길이는 부속표 그대로 — 기호의 팔만 셋이 된다(ㄱ자 → T자).
    assert (top["kind"], top["name"], top["pipe"]) == ("tee", "분류티", "P3")
    assert (top["original_degree"], top["current_degree"]) == (3, 2)
    assert len(top["arms"]) == 3 and top["symbolic"] is False
    down = [u for u in top["arms"] if u[1] < -0.99]
    south = [u for u in top["arms"] if u[0] > 0.8 and -0.6 < u[1] < -0.4]
    assert len(down) == 1 and len(south) == 1          # 세로관 한 번 + 잘린 남쪽 반쪽
    assert len(bottom["arms"]) == 3 and bottom["symbolic"] is False
    assert (asdict(t), g, k) == before


def test_bracket_rule_needs_an_overlapping_joint():
    """겹친 노드가 아니면(50 mm) 종전 규칙 그대로 — 모호하면 그리지 않는다."""
    t, g, k, b, shown, xf = bracket_top_fixture(partner_mm=50.0)
    rows = build_inspection(t, g, k, shown, board=b, transform=xf)["fittings"]
    top = next(r for r in rows if r["node"] == "3")
    assert top["kind"] == "tee" and len(top["arms"]) == 2 and top["symbolic"] is True
