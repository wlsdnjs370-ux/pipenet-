"""F policy excludes run-tee records, without changing physical classification."""
from copy import deepcopy

from src.pipenet_converter.graph.fitting_policy import CALCULATION_FITTINGS
from test_prune_fittings import _net, _LINE, _XYD, _BORES, _PARENT, build_fittings
from test_module_f_fitting_inspection import fixture
from routes.module_f.merge import to_head_tables


def test_policy_keeps_branch_tee_and_elbow_but_not_endpoint_or_interior_runs():
    net = _net(_LINE)
    kw = dict(parents=_PARENT, phys={"W":1,"E":3,"H1":2,"C":3,"N":1},
              interior_junctions={"P2":2})
    before = deepcopy(net)
    legacy = build_fittings(net, _XYD, _BORES, **kw)
    current = build_fittings(net, _XYD, _BORES, allowed_kinds=CALCULATION_FITTINGS, **kw)
    assert legacy["counts"]["tee-run"] == 3
    assert current["counts"] == {"tee":1}
    assert current["per_pipe"]["P4"]["fittings"] == ["tee"]
    assert not current["per_pipe"]["P2"]["fittings"]
    assert {p:r["equivalent_length"] for p,r in current["per_pipe"].items()} == {
        p:r["equivalent_length"] for p,r in legacy["per_pipe"].items()}
    assert net == before


def test_merge_discards_legacy_runs_without_touching_equipment_or_pipe_values():
    tbl,_,_,_ = fixture()
    tbl.fittings.append(dict(pipe="P1",type="tee-run",count=1,**{"in":"1","out":"2"}))
    tbl.fittings[0]["node"]="2"
    tbl.equipment=[dict(pipe="P1",desc="A/V",eq_len=9.5,**{"in":"1","out":"2"})]
    before=deepcopy(tbl)
    merged=to_head_tables(tbl)
    assert [f["type"] for f in merged.fittings] == ["tee","elbow"]
    assert merged.fittings[0]["node"]=="11"
    assert merged.equipment[0]["eq_len"]==9.5
    assert [p["eq_len"] for p in merged.pipes]==[p["eq_len"] for p in tbl.pipes]
    assert tbl==before


def test_policy_keeps_explicit_no_fitting_override_evidence():
    # A shallow bend is unresolved by the angle rule, but the user's confirmed
    # "no fitting" answer is still an applied answer, not a silently missed edit.
    from math import cos, sin, radians
    net=_net({"P1":("W","E"),"P2":("E","H1")})
    xy={"W":(0,0),"E":(1,0),"H1":(1+cos(radians(10)),sin(radians(10)))}
    result=build_fittings(net,xy,_BORES,parents={"E":"W","H1":"E"},
        allowed_kinds=CALCULATION_FITTINGS,
        overrides={"kind":[dict(node="E",pipe="P2",kind="none",note="confirmed") ]})
    assert result["unresolved_kind"]==0
    assert any(r["kind"]=="none" for r in result["applied_overrides"])


def test_straight_tee_cannot_be_reintroduced_by_manual_override(tmp_path):
    from flask import Flask
    from routes.module_f import api_design, jobs
    app=Flask(__name__);app.testing=True
    api_design.register(app,UPLOAD_DIR=tmp_path)
    sess=jobs._new_session(key="fitting-policy-input")
    try:
        client=app.test_client()
        response=client.post('/api/module-f/design/fitting-override',json=dict(
            sid=sess['id'],kind=[dict(node="N2",pipe="P2",kind="tee-run")]))
        assert response.status_code==400
        assert not sess.get('fitting_overrides')
        kinds=client.get('/api/module-f/design/fitting-override',query_string=dict(sid=sess['id'])).json['kinds']
        assert {r['value'] for r in kinds}=={'none','tee','elbow','elbow-45'}
    finally:
        jobs._SESSIONS.pop(sess['id'],None)
