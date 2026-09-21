"""Merged glyphs use combined coordinates and original (unpruned) plan ports."""
from copy import deepcopy

import pytest
from flask import Flask

from routes.module_f import api_merge, jobs
from routes.module_f.merge import merge_network, bake_combined_iso
from routes.module_f.fitting_inspection import build_merged_inspection
from test_module_f_fitting_inspection import fixture
from test_module_f_merge_view import _riser


def merged_fixture():
    t,g,k,b=fixture()
    plan=dict(tables=t,got=g,keys=k,method="manual")
    merged=merge_network(t,riser=_riser(),mode="lsp_gravity")
    return plan,merged,b


@pytest.mark.parametrize("iso",[False,True])
def test_shifted_nodes_original_arms_and_unchanged_hydraulics(iso):
    plan,merged,board=merged_fixture()
    c=merged["combined"];before=deepcopy(c)
    nodes=bake_combined_iso(merged)[0] if iso else c.nodes
    rows=build_merged_inspection(c,plan,nodes,offset=9,board=board,
        transform=dict(k=1,iso=iso,cos30=3**.5/2,sin30=.5))["fittings"]
    tee=next(r for r in rows if r["kind"]=="tee")
    n=next(n for n in nodes if n["label"]=="11")
    assert (tee["node"],tee["x"],tee["y"])==("11",n["x"],n["y"])
    assert len(tee["arms"])==3 and tee["original_degree"]==3
    assert tee["current_degree"]==2 and tee["glyph_radius"]<=200
    assert tee["flow_path"]==["10","11","12"]
    assert c==before


def test_merged_move_does_not_reuse_old_omitted_arm():
    plan,merged,board=merged_fixture()
    before=deepcopy(merged)
    c=merged["combined"]
    next(n for n in c.nodes if n["label"]=="11")["x"]+=500
    rows=build_merged_inspection(c,plan,c.nodes,offset=9,board=board,
        transform=dict(k=1,iso=False),merge_editor=dict(base_object=before))["fittings"]
    tee=next(r for r in rows if r["kind"]=="tee")
    assert len(tee["arms"])==2 and tee["symbolic"]


def test_preview_inactive_plan_slot_has_fitting_cards(tmp_path):
    plan,merged,board=merged_fixture()
    from types import SimpleNamespace
    sess=jobs._new_session(key="merged-fitting-test")
    sess.update(active="system",slots={"plan":dict(design=plan,edit=SimpleNamespace(board=board))},merged=merged)
    app=Flask(__name__);app.testing=True
    api_merge.register(app,UPLOAD_DIR=tmp_path)
    try:
        result=app.test_client().get('/api/module-f/merge/preview',query_string=dict(sid=sess['id'],iso=1))
        assert result.status_code==200,result.json
        rows=result.json["view"]["inspection"]["fittings"]
        tee=next(r for r in rows if r["kind"]=="tee")
        assert tee["node"]=="11" and len(tee["arms"])==3
    finally:
        jobs._SESSIONS.pop(sess["id"],None)


def test_pipe_name_collision_keeps_loss_override_evidence_on_plan_pipe():
    plan,merged,board=merged_fixture()
    c=merged["combined"]
    plan["tables"].unresolved={"applied":[dict(what="eq_len",pipe_label="P2",kind="tee",dia=25,m=4.2,note="confirmed")]}
    next(p for p in c.pipes if p["label"]=="P2")["label"]="P2_2"
    for f in c.fittings:
        if f["pipe"]=="P2":
            f["pipe"]="P2_2"
    rows=build_merged_inspection(c,plan,c.nodes,offset=9,board=board,
        transform=dict(k=1,iso=False))["fittings"]
    tee=next(r for r in rows if r["kind"]=="tee")
    assert (tee["node"],tee["pipe"],tee["eq_m"])==("11","P2_2",4.2)
    assert "confirmed" in tee["eq_source"]


def test_zero_offset_auto_identity_does_not_add_nine():
    t,g,k,b=fixture()
    rows=build_merged_inspection(t,dict(tables=t,got=g,keys=k),t.nodes,
        offset=0,board=b,transform=dict(k=1,iso=False))["fittings"]
    tee=next(r for r in rows if r["kind"]=="tee")
    assert tee["node"]=="2" and len(tee["arms"])==3
