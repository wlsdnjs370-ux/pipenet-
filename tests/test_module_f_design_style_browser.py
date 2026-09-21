"""Calculation preview reuses the edit corridor style without changing data."""
from __future__ import annotations

import json
import os
from copy import deepcopy
from dataclasses import asdict

import pytest
from flask import Flask

from test_module_f_viewport_browser import EDIT, ROOT, install_page, pw_api, settle
from test_module_f_network_editor import design
from routes.module_f import api_design, jobs


VIEW = {
    "nodes": [
        dict(label="1", x=0, y=0, input=True, valve=True),
        dict(label="2", x=120, y=70),
        dict(label="3", x=240, y=0),
        dict(label="4", x=240, y=140, head=True, up=True),
        dict(label="5", x=360, y=70, head=True, up=False),
        dict(label="6", x=480, y=140, head=True, up=False, worst_head=True),
    ],
    "pipes": [
        dict(label="P1", a="1", b="2", load=3, src="text", len_m=4, dia=40),
        dict(label="P2", a="2", b="3", load=2, src="nfpc_min", len_m=4, dia=32),
        dict(label="P3", a="2", b="4", load=1, src="nfpc_fallback", len_m=4, dia=25),
        dict(label="P4", a="3", b="5", load=2, src="text", len_m=4, dia=32),
        dict(label="P5", a="5", b="6", load=1, src="text", len_m=4, dia=25),
    ],
    "worst_path": ["1", "2", "3", "5", "6"],
}


@pytest.fixture()
def page():
    """Render only isolated fixtures, never a live user's session."""
    source = (ROOT / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();',
        '  window.__styleTest = {drawDesign, drawEdit, withDim};\n'
        '  setStage("open");\n  loadSaved();')
    edit = deepcopy(EDIT)
    at = {n["label"]: n for n in VIEW["nodes"]}
    edit["worst"] = dict(max_load=3, heads=[], worst_path=[], corridor=[
        [at[p["a"]]["x"], at[p["a"]]["y"], at[p["b"]]["x"], at[p["b"]]["y"], p["load"]]
        for p in VIEW["pipes"]])
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        p = browser.new_page(viewport={"width": 1440, "height": 1000})
        errors = install_page(p, source)
        p.evaluate("""({view, edit}) => {
          Object.assign(__mf, {sid:'style-test', key:'style-fixture', method:'manual', edit,
            design:{view, tables:{nodes:view.nodes,pipes:view.pipes}, hilite:new Set(),
              gaps:new Map(), marks:{}, sel:null}});
          document.querySelector('#dg-plan').checked=false;
          document.querySelector('#dg-iso').checked=true;
          document.querySelector('#ed-worst-view').value='only';
          __viewTest.setStage('design');
        }""", dict(view=VIEW, edit=edit))
        settle(p)
        yield p
        assert not errors, errors
        browser.close()


def capture(page, renderer="drawDesign", dim=False):
    """Record actual canvas commands/styles while still drawing normally."""
    return page.evaluate("""({renderer, dim}) => {
      const c=document.querySelector('#cv').getContext('2d'), calls=[], old={};
      let path=[];
      for(const name of ['beginPath','moveTo','lineTo','arc','rect','closePath','stroke','fill']) {
        old[name]=c[name];
        c[name]=function(...args) {
          if(name==='beginPath') path=[];
          else if(name==='stroke'||name==='fill') calls.push({op:name,path:path.map(p=>p.slice()),
            color:name==='stroke'?this.strokeStyle:this.fillStyle,alpha:this.globalAlpha,
            width:this.lineWidth,dash:this.getLineDash(),cap:this.lineCap});
          else path.push([name,...args]);
          return old[name].apply(this,args);
        };
      }
      try {
        if(dim) __styleTest.withDim(__styleTest[renderer]);
        else __styleTest[renderer]();
      } finally { for(const [name,fn] of Object.entries(old)) c[name]=fn; }
      return calls;
    }""", dict(renderer=renderer, dim=dim))


def strokes(calls):
    return [c for c in calls if c["op"] == "stroke"]


def test_default_pipes_match_edit_corridor(page):
    assert not page.locator("#dg-bore-color").is_checked()
    assert page.evaluate("__mf.boreColor") is False
    expected = strokes(capture(page, "drawEdit"))[:len(VIEW["pipes"])]
    actual = strokes(capture(page))[:len(VIEW["pipes"])]
    for a, e in zip(actual, expected):
        for key in ("color", "width", "alpha", "dash", "cap"):
            assert a[key] == e[key], (key, a, e)


def test_heads_use_only_directional_triangles_and_red_anchor(page):
    calls = strokes(capture(page))
    triangles = [c for c in calls if any(p[0] == "closePath" for p in c["path"])]
    assert len(triangles) == 3
    for triangle, up in zip(triangles, [True, False, False]):
        points = [p for p in triangle["path"] if p[0] in ("moveTo", "lineTo")]
        assert len(points) == 3
        assert (points[0][2] < points[1][2]) is up  # Screen Y increases downward.
    assert [t["color"] for t in triangles] == ["#ffffff", "#ffffff", "#ff3b3b"]
    assert not any(p[0] == "arc" for c in calls for p in c["path"])


def test_diagnostic_colors_remain_optional_and_do_not_change_data(page):
    before = page.evaluate("JSON.stringify([__mf.design.view,__mf.design.tables])")
    page.locator('[data-fold="dg-evidence-body"]').click()
    page.check("#dg-bore-color")
    settle(page)
    pipes = strokes(capture(page))[:len(VIEW["pipes"])]
    assert [p["color"] for p in pipes[:3]] == ["#38bdf8", "#facc15", "#64748b"]
    assert pipes[2]["dash"] == [6, 4]
    page.uncheck("#dg-bore-color")
    settle(page)
    assert all(p["color"] == "#ffffff" for p in strokes(capture(page))[:5])
    assert page.evaluate("JSON.stringify([__mf.design.view,__mf.design.tables])") == before


def test_selection_dimming_keeps_table_highlight_and_path_hierarchy(page):
    page.evaluate("__mf.design.sel={kind:'pipe',label:'P2'};__mf.design.hilite.add('P1')")
    calls = strokes(capture(page, dim=True))
    assert calls[0]["color"] == "#f97316" and calls[0]["alpha"] == 1
    assert 0 < calls[1]["alpha"] <= .35
    red_path = next(c for c in calls if c["color"] == "#ff3b3b" and c["dash"])
    assert red_path["alpha"] == pytest.approx(.85 * .35)
    assert red_path["width"] == pytest.approx(2.2) and red_path["dash"] == [9, 5]


def test_render_preserves_graph_coordinates_and_optional_screenshot(page):
    before = page.evaluate("JSON.stringify([__mf.design.view,__mf.design.tables])")
    for _ in range(3):
        capture(page)
    assert json.loads(before)[0] == VIEW
    assert page.evaluate("JSON.stringify([__mf.design.view,__mf.design.tables])") == before
    if os.environ.get("MODULE_F_DESIGN_STYLE_SCREENSHOT"):
        page.locator("#cv").screenshot(path=os.environ["MODULE_F_DESIGN_STYLE_SCREENSHOT"])


def test_crossing_gaps_still_break_the_lower_pipe_only(page):
    page.evaluate("__mf.design.gaps=new Map([[0,[.5]]])")
    pipes = strokes(capture(page))[:5]
    assert [p[0] for p in pipes[0]["path"]] == ["moveTo", "lineTo", "moveTo", "lineTo"]
    assert [p[0] for p in pipes[1]["path"]] == ["moveTo", "lineTo"]


@pytest.mark.parametrize("iso", ["0", "1"])
def test_preview_head_direction_comes_from_elevation_not_projection(tmp_path, iso):
    """Both glyph orientations use the existing API contract, without rebaking data."""
    sess = jobs._new_session(key="direction-style-test")
    sess["design"] = design()
    tbl = sess["design"]["tables"]
    tbl.nodes[2]["elevation"] = -.5
    tbl.nodes.append(dict(label="4", x=93100, y=39154, elevation=.5, io_node="No"))
    tbl.pipes.append(dict(label="P3", **{"in": "2", "out": "4"},
                         length=3, dia=25, type="KSD 3507", c=120, eq_len=0))
    tbl.nozzles.append(dict(label="H2", **{"in": "4", "out": "@/4"},
                           flow_lmin=80, flow_m3s=80/60000, lib="SP-HEAD", status="1"))
    before = deepcopy(asdict(tbl))
    app = Flask(__name__)
    app.testing = True
    api_design.register(app, UPLOAD_DIR=tmp_path)
    try:
        response = app.test_client().get("/api/module-f/design/preview",
            query_string=dict(sid=sess["id"], iso=iso))
        assert response.status_code == 200, response.json
        heads = {n["label"]: n for n in response.json["view"]["nodes"] if n.get("head")}
        assert heads["3"]["up"] is False
        assert heads["4"]["up"] is True
        assert asdict(tbl) == before
    finally:
        jobs._SESSIONS.pop(sess["id"], None)
