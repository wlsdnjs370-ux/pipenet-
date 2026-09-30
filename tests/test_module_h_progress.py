"""Read-only live observations preserve exact F parsing/graph results."""
from __future__ import annotations

from copy import deepcopy
from threading import Event
import uuid

from flask import Flask, jsonify

from routes.module_f import jobs
from routes.module_f.common import _boot
from routes.module_f.world import _world_payload
from routes.module_h_progress import PreviewStore, install
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.progress import ProgressEvent, current_reporter, report, reporting
from test_module_f_world_display import drawing
from test_module_f_cancel import wait_until


def test_flow_observations_are_actual_order_and_preserve_cycles_and_inputs():
    points=[(0,0),(100,0),(100,100),(0,100),(1000,1000),(1100,1000)]
    edges=[(0,1),(1,2),(2,3),(3,0),(4,5)]
    before=deepcopy((points,edges))
    expected=build_flow_tree(points,edges,[{2}],[0])
    events=[]
    with reporting(events.append):
        actual=build_flow_tree(points,edges,[{2}],[0])
    assert actual==expected and (points,edges)==before
    assert events[0].kind=="reset" and events[0].data["roots"]==[[0,0]]
    segments=[row for e in events if e.kind=="graph" for row in e.data["segments"]]
    assert segments[0]==[0,0,100,0]
    # All four cycle edges are explored, but the unreachable component is not fabricated.
    assert len(segments)==4
    assert all(max(row)<=100 for row in segments)
    assert current_reporter() is None
    def broken(event):
        raise RuntimeError("preview disconnected")
    with reporting(broken):
        assert build_flow_tree(points,edges,[{2}],[0])==expected


def test_decoded_geometry_arrives_incrementally_and_payload_is_identical():
    _boot()
    from services.cad_import.pipeline.stage1 import explode
    ents=[{"t":"LINE","8":f"layer-{i%3}","62":7,"10":i*10,"20":0,"11":i*10,"21":100} for i in range(300)]
    expected,_=explode({},ents,{})
    events=[]
    with reporting(events.append):
        actual,_=explode({},ents,{})
    assert actual.__dict__==expected.__dict__
    chunks=[e.data for e in events if e.kind=="world"]
    assert chunks[0]["count"]<chunks[-1]["count"]==300
    assert all(len(e["items"])<=256 for e in chunks)
    assert {r["layer"] for e in chunks for r in e["items"]}=={"layer-0","layer-1","layer-2"}
    raw=drawing();before=deepcopy(raw.__dict__)
    baseline=_world_payload(raw);events=[]
    with reporting(events.append):
        assert _world_payload(raw)==baseline
    assert raw.__dict__==before
    assert any(r["layer"]=="late-loop" for e in events if e.kind=="world" for r in e.data["items"])


def test_store_bounds_cursor_gaps_and_detached_geometry():
    store=PreviewStore();op=uuid.uuid4().hex
    data={"segments":[[0,0,100,100]]}
    store.publish(op,ProgressEvent("graph","trace",data))
    data["segments"][0][0]=999
    assert store.read(op,0)["events"][0]["data"]["segments"][0][0]==0
    for i in range(250):
        store.publish(op,ProgressEvent("graph","trace",{"count":i}))
    got=store.read(op,0)
    assert got["gap"] and len(got["events"])==48
    assert not store.read(op,got["cursor"])["gap"]
    assert len(store.channels[op].events)==192
    for _ in range(15):
        store.read(uuid.uuid4().hex,0)
    assert len(store.channels)==12


def test_h_origin_only_scope_and_real_worker_propagation():
    app=Flask(__name__);app.testing=True;install(app)
    sess=jobs._new_session()
    entered=Event();release=Event()
    @app.post("/api/module-f/test-preview")
    def run():
        def work():
            report("reset","stage",mode="world")
            entered.set();release.wait(3)
            return {"same":"result"}
        jobs._run_job(sess,"preview-test",work)
        return jsonify(ok=True)
    try:
        client=app.test_client();op=uuid.uuid4().hex
        client.post("/api/module-f/test-preview",headers={"Referer":"http://localhost/module-h","X-Module-F-Operation":op})
        assert entered.wait(3)
        events=client.get(f"/api/module-h/progress?operation={op}").json["events"]
        assert events and sess["job"]["state"]=="run"  # Actual worker is still processing.
        assert current_reporter() is None
        release.set();wait_until(lambda:sess["job"]["state"]!="run")
        assert sess["job"]["result"]=={"same":"result"}
        entered.clear();op2=uuid.uuid4().hex
        client.post("/api/module-f/test-preview",headers={"Referer":"http://localhost/module-f","X-Module-F-Operation":op2})
        assert entered.wait(3)
        wait_until(lambda:sess["job"]["state"]!="run")
        assert client.get(f"/api/module-h/progress?operation={op2}").json["events"]==[]
        assert client.get("/api/module-h/progress?operation=invalid").status_code==400
    finally:
        release.set();jobs._SESSIONS.pop(sess["id"],None)
