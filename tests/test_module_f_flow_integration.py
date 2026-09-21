"""Real HTTP/edit/expansion/UI contracts for the canonical water-flow tree."""
from __future__ import annotations

import copy
import json

import pytest
from flask import Flask

from routes.module_f.common import _boot
from routes.module_f import api_edit, api_design, jobs

_boot()
from services.cad_import.edit.board import EditBoard
from services.cad_import.edit.session import EditSession
from services.cad_import.design.flow import flow_for_board
from services.cad_import.design.restrict import expand_worst, restrict_to_worst


@pytest.fixture
def flow_client(tmp_path, monkeypatch):
    b = EditBoard("flow-fixture", [(0, 0), (1000, 0), (2000, 0), (3000, 0),
                  (2000, 2000), (3000, 2000), (4000, 0), (2000, -1000)],
                  {(0, 1), (1, 2), (2, 3), (3, 5), (5, 4), (2, 4), (3, 6), (2, 7)},
                  [(2000, 2000, 30), (3000, 2000, 30)])
    for d in b.disks:
        b.set_head_kind(d, "상향식")
    b.sources = [0]
    b.valves = [0]
    es = EditSession(b, key="flow-fixture", out_dir=str(tmp_path))
    sess = jobs._new_session(key=es.key)
    sess["edit"] = es
    app = Flask(__name__)
    app.testing = True
    api_edit.register(app)
    api_design.register(app, UPLOAD_DIR=tmp_path)
    def run(session, phase, fn):
        result = fn()
        session["job"] = {"state": "done", "result": result}
    monkeypatch.setattr(api_design, "_run_job", run)
    monkeypatch.setattr(api_design, "_dia_texts", lambda sess: [])
    yield app.test_client(), sess
    jobs._SESSIONS.pop(sess["id"], None)


def test_flow_worst_design_report_share_routes_and_source_is_untouched(flow_client):
    c, sess = flow_client
    b = sess["edit"].board
    original = copy.deepcopy((b.pts, b.edges))
    flow = c.post("/api/module-f/edit/flow", json={"sid": sess["id"]})
    assert flow.status_code == 200, flow.get_json()
    report = flow.get_json()["state"]["flow_report"]
    assert report["heads"] == 2 and report["excluded"]["alternate_route"] > 0
    tree = flow_for_board(b)
    worst = c.post("/api/module-f/edit/worst", json={"sid": sess["id"], "k": 1})
    assert worst.status_code == 200, worst.get_json()
    assert sess["worst"]["flow_revision"] == report["revision"]
    assert sess["worst"]["edges"] <= tree.edges
    c.post("/api/module-f/design/build", json={"sid": sess["id"], "k": 1})
    assert sess["job"]["result"]["ok"], sess["job"]["result"]
    got = sess["design"]["got"]
    assert got["worst"]["edges"] == sess["worst"]["edges"]
    assert max(got["physical_pipe_loads"].values()) == 2
    assert max(got["tree_loads"].values()) == 1
    assert any(p["head_count_all"] == 2 and p["head_count_selected"] == 1
               for p in sess["design"]["tables"].pipes)
    data = c.get("/api/module-f/edit/flow/report", query_string={"sid": sess["id"]})
    assert data.status_code == 200 and "attachment" in data.headers["Content-Disposition"]
    assert data.get_json()["summary"]["revision"] == report["revision"]
    assert (b.pts, b.edges) == original
    # A source tee stays physical even after pruning one calculation arm.
    assert max(got["phys"].values()) >= 3
    b.pts[1] = (1100, 0)  # Same number of nodes and edges, different basis.
    c.post("/api/module-f/design/build", json={"sid": sess["id"], "k": 1})
    assert not sess["job"]["result"]["ok"]
    assert c.get("/api/module-f/edit/flow/report", query_string={"sid": sess["id"]}).status_code == 400


def test_source_change_and_edit_invalidate_visible_tree(flow_client):
    c, sess = flow_client
    b = sess["edit"].board
    b.sources = [0, 6]
    assert c.post("/api/module-f/edit/flow", json={"sid": sess["id"]}).status_code == 400
    old = c.post("/api/module-f/edit/flow", json={"sid": sess["id"], "source": "Z1"}).get_json()
    new = c.post("/api/module-f/edit/flow", json={"sid": sess["id"], "source": "Z2"}).get_json()
    assert old["state"]["flow_report"]["revision"] != new["state"]["flow_report"]["revision"]
    assert "wet_pipes" not in new["state"].get("keep", [])
    api_edit._note_edit(sess)
    state = c.get("/api/module-f/edit/state", query_string={"sid": sess["id"]}).get_json()["state"]
    assert not state["flowed"] and state["flow_report"] is None and not state["wet_pipes"]


def test_full_conversion_uses_same_tree(flow_client):
    from services.cad_import.convert.engine import ensure_planar
    c, sess = flow_client
    es = sess["edit"]
    tree = flow_for_board(es.board)
    payload = restrict_to_worst(es.convert_payload(), es.board, {"heads": list(tree.head_node)})
    assert {tuple(e) for e in payload["edges"]} == tree.edges
    built = ensure_planar(payload)
    net = built["kfp"]
    assert len(net["pipe_data"]) == len(net["nodes_meta_runtime"]) - 1
    assert sum(p["length_m"] for p in net["pipe_data"].values()) == pytest.approx(
        sum(tree.lengths_mm[e] for e in tree.edges) / 1000)


def test_full_conversion_retains_declared_length_after_straight_merge(flow_client):
    from services.cad_import.convert.engine import ensure_planar
    _, sess = flow_client
    es = sess["edit"]
    es.board.edge_len_mm = {(0, 1): 4500}
    tree = flow_for_board(es.board)
    payload = restrict_to_worst(es.convert_payload(), es.board, {"heads": list(tree.head_node)})
    built = ensure_planar(payload)
    assert built.get("kfp"), built.get("_planar_error")
    assert sum(p["length_m"] for p in built["kfp"]["pipe_data"].values()) == pytest.approx(
        sum(tree.lengths_mm[e] for e in tree.edges) / 1000)


def test_browser_flow_button_report_and_dimmed_original(flow_client, tmp_path):
    from test_module_f_viewport_browser import install_page, settle, pw_api, ROOT
    from urllib.parse import urlparse
    c, sess = flow_client
    source = (ROOT / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();',
        '  window.__flowTest={setEdit,renderEdit};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={"width": 1440, "height": 1000})
        errors = install_page(page, source)
        def route(req):
            url = urlparse(req.request.url)
            r = c.open(url.path + ("?" + url.query if url.query else ""), method=req.request.method,
                       data=req.request.post_data, content_type="application/json")
            if r.status_code == 404:
                req.fulfill(json={"ok": True, "rows": [], "items": [], "fields": {}})
            else:
                req.fulfill(status=r.status_code, body=r.data, content_type=r.content_type)
        page.route("http://module-f.test/api/module-f/**", route)
        state = c.get("/api/module-f/edit/state", query_string={"sid": sess["id"]}).get_json()["state"]
        page.evaluate("""({sid,state}) => {
          Object.assign(__mf,{sid,slot:'plan',method:'manual'});
          __flowTest.setEdit(state);__viewTest.setStage('edit');__flowTest.renderEdit();
        }""", {"sid": sess["id"], "state": state})
        page.click("#ed-flow")
        page.wait_for_function("!!__mf.edit.flow_report")
        assert "전체 부하" in page.locator("#ed-info").text_content()
        assert "단일 경로 확정" in page.locator("#ed-flow-summary").inner_text()
        strokes = page.evaluate("""async () => {
          const ctx=document.querySelector('#cv').getContext('2d'), old=ctx.stroke, out=[];
          ctx.stroke=function(){out.push({alpha:this.globalAlpha,width:this.lineWidth});return old.call(this);};
          try{__viewTest.draw();await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));}
          finally{ctx.stroke=old;} return out;
        }""")
        assert any(s["alpha"] == pytest.approx(.22) for s in strokes)
        assert any(s["alpha"] == 1 and s["width"] == pytest.approx(2.6) for s in strokes), strokes
        assert not errors, errors
        page.click('[data-fold="ed-status-body"]')
        page.screenshot(path=str(ROOT / "data/module_f_flow_tree_ui.png"))
        browser.close()
