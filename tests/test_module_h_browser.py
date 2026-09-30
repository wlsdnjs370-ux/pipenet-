"""Real H/F browser handlers against isolated engine fixtures, never live data."""
from pathlib import Path
import xml.etree.ElementTree as ET
from threading import Thread
import time
import re
import json
from queue import Queue, Empty

import pytest
from flask import Flask, Response, jsonify, request
from werkzeug.serving import make_server

from routes.module_h import register
from routes.module_f import api_convert, api_design, api_edit, api_merge, api_network_edit, api_sizing, api_slot, jobs, network_edit as ne
from test_module_f_network_editor import design
from test_module_f_viewport_browser import WORLD, EDIT, pw_api, settle

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture()
def h_ui(tmp_path, monkeypatch):
    monkeypatch.setattr(ne, "HISTORY_DIR", tmp_path / "history")
    monkeypatch.syspath_prepend(str(ROOT / "core"))
    # Each isolated app needs the G engine path even when a previous test's
    # monkeypatch restored sys.path after common._boot() had already run.
    monkeypatch.syspath_prepend(str(ROOT / "cad_project_editor_g"))
    def run_now(session, phase, function):
        session["job"] = {"state":"done","phase":phase,"started":time.time(),
            "ended":time.time(),"error":None,"result":function()}
    monkeypatch.setattr(api_sizing, "_run_job", run_now)
    monkeypatch.setattr(api_merge, "_run_job", run_now)
    sess = jobs._new_session(key="h-browser-isolated")
    sess["design"] = design()
    app = Flask(__name__, template_folder=str(ROOT / "templates"), static_folder=str(ROOT / "static"))
    app.testing = True
    register(app)
    api_slot.register(app, _save_upload=lambda *_args, **_kwargs: tmp_path/'unused.dxf')
    api_edit.register(app)
    api_design.register(app, UPLOAD_DIR=tmp_path)
    api_merge.register(app, UPLOAD_DIR=tmp_path)
    api_network_edit.register(app)
    api_convert.register(app, UPLOAD_DIR=tmp_path)
    source = (ROOT / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();',
        '  window.__hTest={setStage,setEdit,renderEdit,draw,busy,watch,stopWatch,say,post,loadSubGraph,subExtract,sysDefaultHidden,mergeState:loadMergeState,preview:designPreview,select:insSelect,merged:loadMergeView,table:renderDesignTable};\n  setStage("open");\n  loadSaved();')
    requests = []
    errors = []
    @app.before_request
    def record_request():
        requests.append((request.method, request.path))

    @app.get("/static/module_f.js")
    def engine_with_test_hooks():
        return Response(source, mimetype="application/javascript")

    @app.get("/api/module-f/job")
    def completed_job():
        return jsonify(sess.get("_h_job_fixture") or dict(ok=True,state="done",lines=[],elapsed=0))

    @app.get("/api/module-f/job/stream")
    def no_stream():
        stream = sess.get("_h_stream_fixture")
        if stream is None:
            return "",404
        def events():
            while True:
                try:
                    event, value = stream.get(timeout=5)
                except Empty:
                    return
                yield f"event: {event}\ndata: {json.dumps(value)}\n\n"
                if event == "state" and value.get("state") in {"done", "error", "cancelled"}:
                    return
        return Response(events(), mimetype="text/event-stream")

    @app.post("/api/module-f/job/cancel")
    def cancel_fixture_job():
        sess["_h_cancel_body"] = request.get_json()
        return jsonify(ok=True, stopped=True, rollback_errors=[])

    @app.post("/api/module-f/test-preview-flow")
    def real_observed_flow():
        from src.pipenet_converter.graph.flow import build_flow_tree
        flow=build_flow_tree([(0,0),(100,0),(100,100)],[(0,1),(1,2)],[{2}],[0])
        return jsonify(ok=True,revision=flow.revision)

    @app.errorhandler(404)
    def optional_bootstrap_lists(error):
        if request.path.startswith("/api/module-f/"):
            return jsonify(ok=True,items=[],rows=[],fields={})
        return "Not found",404

    server = make_server("127.0.0.1",0,app,threaded=True)
    thread = Thread(target=server.serve_forever,daemon=True)
    thread.start()
    try:
        with pw_api.sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page(viewport={"width":1500,"height":960})
            page._h_preview_store = app.extensions["module_h_preview"]
            page._h_app = app
            page.set_default_timeout(10000)
            page.on("pageerror", lambda error: errors.append(error.stack))

            page.goto(f"http://127.0.0.1:{server.server_port}/module-h")
            page.wait_for_function("!!window.__hTest")
            assert page.locator("#h-layout-error").count() == 0
            assert page.locator("body").evaluate("el=>el.classList.contains('h-ready')")
            yield page, sess, requests, tmp_path
            assert not errors, errors
            browser.close()
    finally:
        server.shutdown()
        thread.join(timeout=2)
        server.server_close()
        jobs._SESSIONS.pop(sess["id"], None)


def seed(page, sess, stage="edit"):
    page.evaluate("""({sid,world,edit,stage})=>{
      Object.assign(__mf,{sid,slot:'plan',method:'manual',world,edit});
      __hTest.setStage(stage);
    }""", {"sid":sess["id"],"world":WORLD,"edit":EDIT,"stage":stage})
    settle(page)


def menu(page, control):
    page.click("#h-more")
    page.click(control)


def test_area_selection_controls_and_real_request_have_no_quota(h_ui):
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import EditSession
    page,sess,requests,_ = h_ui
    board = EditBoard('h-area-browser',[(0,0),(1000,0),(2000,0),(1000,2000),(2000,2000)],
                      {(0,1),(1,2),(1,3),(2,4)},[(1000,2000,30),(2000,2000,30)])
    board.sources = [0]
    for disk in board.disks:
        board.set_head_kind(disk,'상향식')
    sess['edit'] = EditSession(board,key=board.key)
    state = page.request.get(page.url.replace('/module-h','/api/module-f/edit/state'),
                             params={'sid':sess['id']}).json()['state']
    page.evaluate('''({sid,state,world})=>{
      Object.assign(__mf,{sid,world,slot:'plan',method:'manual'});
      __hTest.setEdit(state);__hTest.setStage('edit');__hTest.renderEdit();
    }''',{'sid':sess['id'],'state':state,'world':WORLD})
    settle(page)
    assert page.locator('#h-primary').inner_text() == '영역 지정'
    page.click('#h-primary')
    assert page.locator('#ed-zone-arm').is_checked()
    for control in ['ed-k','ed-k-preset','ed-sheet-wrap','ed-sheets','adv-auto']:
        assert page.locator('#'+control).is_hidden()
    assert not any(path.endswith('/worst/reference-counts') for _,path in requests)
    page.click('#ed-worst')
    assert '영역을 먼저 지정' in page.locator('#ed-worst-err').inner_text()
    assert not sess.get('worst')
    # Deliberately poison the dormant legacy count; the actual request must omit it.
    page.evaluate("__mf.zones=[[900,1900,2100,2100]];document.querySelector('#ed-k').value='1';__hTest.renderEdit()")
    # Static files may refresh before Python restarts. Never send a no-K request
    # to an older server that would silently use its default reference count.
    page.route('**/api/module-h/selection-policy',lambda route: route.fulfill(status=404,body='old server'))
    before = len([p for method,p in requests if method=='POST' and p.endswith('/edit/worst')])
    page.click('#ed-worst')
    page.wait_for_function("document.querySelector('#ed-worst-err').textContent.includes('서버 재시작')")
    assert len([p for method,p in requests if method=='POST' and p.endswith('/edit/worst')]) == before
    page.unroute('**/api/module-h/selection-policy')
    with page.expect_request('**/api/module-f/edit/worst') as sent:
        page.click('#ed-worst')
    payload = sent.value.post_data_json
    assert payload['selection_mode'] == 'area_all' and 'k' not in payload and 'sheet' not in payload
    page.wait_for_function("__mf.edit.worst?.selection_mode==='area_all'")
    page.wait_for_function("document.querySelector('#h-activity').dataset.state==='done'")
    assert page.evaluate('__mf.edit.worst.k') == 2
    assert set(sess['worst']['heads']) == {0,1}
    assert '헤드 2개 전체 추출' in page.locator('#status').text_content()
    output = ROOT/'data/module_h_workspace_20260928'
    output.mkdir(exist_ok=True)
    page.screenshot(path=str(output/'area-all-controls.png'))
    page.click('#ed-zone-clear')
    page.wait_for_function('!__mf.edit.worst')
    assert not sess.get('worst')
    assert page.locator('#h-primary').inner_text() == '영역 지정'


def test_landing_has_one_workspace_and_preserves_controls(h_ui):
    page, sess, requests, tmp_path = h_ui
    assert page.locator("#h-welcome").is_visible()
    assert page.locator(".h-conversion-mark .h-format").all_text_contents() == ["DXF", "SDF"]
    assert page.locator(".h-conversion-mark [data-unit]").count() == 5
    assert page.locator(".h-conversion-mark rect").count() == 4
    assert page.locator("#h-workflow").count() == 0
    assert page.locator("#steps").is_hidden()
    assert page.locator(".side").is_hidden()
    assert page.locator("#h-primary").inner_text() == "Access"
    for control in re.findall(r'\bid="([^"]+)"', (ROOT / "templates/module_f.html").read_text(encoding="utf-8")):
        assert page.locator(f'[id="{control}"]').count() == 1, control
    assert page.evaluate("getComputedStyle(document.body).getPropertyValue('--accent').trim()") == "#5B8DEF"
    page.screenshot(path=str(tmp_path / "module-h-landing.png"))
    before = len([r for r in requests if r[0] == "POST"])
    page.click("#h-primary")
    page.click('#h-access-plan')
    page.wait_for_function('!document.querySelector("#h-drawer").hidden')
    assert page.locator("#dxf").is_visible()
    page.keyboard.press("Escape")
    page.click("#h-history")
    assert page.locator("#steps").is_visible()
    assert len([r for r in requests if r[0] == "POST"]) == before
    seed(page,sess)
    page.click("#h-history")
    page.locator("#steps div",has_text="객체 정의").click()
    page.wait_for_function("__mf.stage==='pick'")
    page.click("#h-primary")
    assert page.locator("#pk-crop-pen").is_visible()
    page.click("#pk-crop-pen")
    assert page.locator("#pk-crop-pen").get_attribute("aria-pressed") == "true"
    page.keyboard.press("Escape")
    page.click("#h-history")
    page.locator("#steps div",has_text="배관 네트워크 정의").focus()
    page.keyboard.press("Enter")
    page.wait_for_function("__mf.stage==='edit'")
    page.click("#h-primary")
    assert page.locator("#ed-network-mode").is_visible()
    assert page.locator("#ed-aj-scan").is_hidden()
    page.get_by_text("끊긴 배관 자동 이음",exact=True).click()
    assert page.locator("#ed-aj-scan").is_visible()


def test_primary_never_skips_prerequisites_or_runs_from_layout(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess)
    assert page.locator("#h-primary").inner_text() == "연결점 지정하기"
    page.evaluate("""()=>{
      window.__actionCount=0;
      document.querySelector('#ed-worst').onclick=()=>++window.__actionCount;
      __mf.edit.sources=[[0,0]];__hTest.setStage('edit');
    }""")
    settle(page)
    assert page.locator("#h-primary").inner_text() == "영역 지정"
    page.click("#h-primary")
    assert page.locator("#ed-zone-arm").is_checked()
    assert page.evaluate("window.__actionCount") == 0
    page.keyboard.press("Escape")
    page.evaluate("__mf.zones=[[0,0,100,100]];__hTest.setStage('edit')")
    settle(page)
    assert page.locator("#h-primary").inner_text() == "영역 배관망 추출"
    page.click("#h-primary")  # settings, not calculation
    assert page.evaluate("window.__actionCount") == 0
    page.click("#h-primary")
    assert page.evaluate("window.__actionCount") == 1
    page.evaluate("document.querySelector('#ed-worst').disabled=true")
    settle(page)
    assert page.locator("#h-primary").is_disabled()
    page.evaluate("document.querySelector('#ed-worst').disabled=false;document.querySelector('#busy').classList.remove('hidden')")
    settle(page)
    assert page.locator("#h-primary").is_disabled()
    page.evaluate("document.querySelector('#busy').classList.add('hidden')")
    seed(page,sess,"sub")
    page.evaluate("__mf.slot='system';__hTest.setStage('sub')")
    settle(page)
    assert page.locator("#h-primary").inner_text() == "경로 추출"
    page.click("#h-more")
    assert page.locator("#h-table").is_disabled()
    page.keyboard.press("Escape")
    page.evaluate("__mf.slot='plan';__mf.method='auto';__hTest.setStage('auto')")
    settle(page)
    assert page.locator("#h-primary").inner_text() == "도면 다시 업로드"
    page.click("#h-primary")
    page.wait_for_function("__mf.stage==='open'")


def test_view_table_and_stale_warnings_preserve_data(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,"design")
    page.evaluate("async()=>{await __hTest.preview();}")
    before=page.evaluate("JSON.stringify(__mf.design.view)")
    menu(page,"#h-view")
    assert page.locator("#dg-plan").is_visible()
    page.uncheck("#dg-plan")
    settle(page)
    assert page.evaluate("JSON.stringify(__mf.design.view)") == before
    assert page.locator("#dg-fx").is_hidden()
    page.keyboard.press("Escape")
    assert page.locator("#h-more").evaluate("el=>el===document.activeElement")
    menu(page,"#h-table")
    page.wait_for_selector('#h-properties:not([hidden]) #hp-rows tr')
    assert page.locator("#h-properties").bounding_box()["width"] > 700
    assert page.evaluate("JSON.stringify(__mf.design.view)") == before
    page.keyboard.press("Escape")
    assert page.locator('#h-properties').is_hidden()
    page.evaluate("""()=>{
      const el=document.querySelector('#dg-stale');
      el.textContent='표를 다시 확정해야 합니다';el.classList.remove('hidden');
    }""")
    settle(page)
    assert page.locator("#h-alert").is_visible()
    assert page.locator("#h-primary").inner_text() == "입력값 반영"
    page.click("#h-alert-open")
    assert page.locator("#dg-stale").is_visible()
    page.evaluate("document.querySelector('#dg-build').disabled=true")
    settle(page)
    assert page.locator("#h-primary").is_disabled()


def test_sizing_runs_and_emits_separate_valid_sdf_without_changing_source(h_ui):
    from test_module_f_sizing import tables, merged
    from routes.module_f.hydraulic_sizing import fingerprint
    page, sess, requests, tmp_path = h_ui
    sess["merged"] = merged(tables())
    sess['merge_summary']={'merged':True}
    before = fingerprint(sess["merged"]["combined"])
    seed(page,sess,"merge")
    page.evaluate('async()=>{await __hTest.mergeState();await __hTest.merged();}')
    menu(page,"#h-sizing")
    page.click("#sz-load")
    page.wait_for_function("document.querySelector('#sz-basis').textContent.includes('배관 5개')")
    page.check("#sz-confirm")
    page.click("#sz-run")
    page.wait_for_function("document.querySelector('#sz-result').textContent.includes('조건 충족')",timeout=20000)
    assert sess["sizing_report"]["feasible"]
    assert page.locator("#sz-emit").is_enabled()
    assert fingerprint(sess["merged"]["combined"]) == before
    # The primary merged-output action must emit the proposal, not the original
    # tables (which may intentionally still contain unresolved drawing bores).
    page.wait_for_function("document.querySelector('#h-output-basis').value==='proposal'")
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden') && !document.querySelector('#mg-emit').disabled")
    page.evaluate("document.querySelector('#mg-emit').click()")
    page.wait_for_function("!document.querySelector('#sz-download').disabled",timeout=20000)
    sdf = Path(sess["sizing_files"]["sdf_iso"])
    assert sdf.is_relative_to(tmp_path)
    assert ET.parse(sdf).getroot() is not None
    assert fingerprint(sess["merged"]["combined"]) == before
    assert ('POST','/api/module-f/merge/sizing/emit') in requests
    assert ('POST','/api/module-f/merge/emit') not in requests
    with page.expect_download() as download:
        page.click("#sz-download")
    assert download.value.suggested_filename.endswith(".sdf")
    page.screenshot(path=str(tmp_path / "module-h-sizing-result.png"))
    page.fill("#sz-flow","90")
    assert page.locator("#sz-emit").is_disabled()
    assert page.locator("#sz-download").is_disabled()
    assert "다시 계산" in page.locator("#sz-result").inner_text()


def test_review_edit_and_undo_still_call_real_engine(h_ui):
    page, sess, requests, _ = h_ui
    seed(page,sess,"design")
    menu(page,"#h-view")
    page.uncheck("#dg-plan")
    page.click("#h-drawer-close")
    page.evaluate("async()=>{await __hTest.preview();__hTest.select('pipe','P2');}")
    menu(page,"#h-table")
    page.wait_for_selector('#hp-rows tr[data-label="P2"]')
    page.locator('#hp-rows tr[data-label="P2"] [data-field=dn]').select_option('32')
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===32 && document.querySelector('#busy').classList.contains('hidden')")
    assert sess["design"]["tables"].pipes[1]["dia"] == 32
    page.click("#hp-undo")
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===25 && document.querySelector('#busy').classList.contains('hidden')")
    assert ("POST","/api/module-f/network-editor") in requests


def test_merged_drawers_sizing_guard_and_viewport(h_ui):
    page, sess, requests, tmp_path = h_ui
    from test_module_f_merge import _riser
    sess["slots"]["system"]["riser"] = _riser()
    machine=_riser(2)
    machine['nodes'][0]['label']='M1';machine['nodes'][1]['label']='M2'
    machine['pipes'][0].update(label='MP1',**{'in':'M1','out':'M2'})
    machine.update(source_node_label='M1',conn_node_label='M2')
    sess['slots']['machineroom']['machineroom']=machine
    sess["supply_mode"] = "lsp_gravity"
    api_merge.rebuild_merged(sess)
    # The existing _riser fixture tests joins only; supply explicit SDF fields.
    for pipe in sess["merged"]["combined"].pipes:
        pipe.setdefault("c",120)
    seed(page,sess,"design")
    menu(page,"#h-merge")
    page.wait_for_function("__mf.stage==='merge'")
    page.wait_for_function("!!__mf.mergeView?.pipes.length")
    before = page.evaluate("JSON.stringify(__mf.mergeView)")
    menu(page,"#h-sizing")
    assert page.locator("#sz-flow").input_value() == "80"
    assert page.locator("#sz-branch").input_value() == "6"
    assert page.locator("#sz-other").input_value() == "10"
    assert page.locator("#sz-suction").is_hidden()
    assert page.locator("#sz-confirm").is_visible()
    page.click("#sz-run")
    page.wait_for_function("document.querySelector('#h-drawer-status').textContent.includes('오류')")
    assert "먼저 불러오세요" in page.locator("#sz-result").inner_text()
    assert not any(method=="POST" and path.endswith("sizing/run") for method,path in requests)
    assert page.evaluate("JSON.stringify(__mf.mergeView)") == before
    page.click("#h-drawer-close")
    menu(page,"#h-export")
    assert page.locator("#mg-emit").is_visible()
    assert page.locator("#mg-dl-sdf").is_visible()
    assert page.locator("#mg-dl-kfp").is_hidden()
    assert page.locator("#sz-emit").is_hidden()
    page.get_by_text("다른 형식 · 편집용 사본",exact=True).last.click()
    assert page.locator("#mg-dl-kfp").is_visible()
    page.click("#mg-emit")
    page.wait_for_function("!document.querySelector('#mg-dl-sdf').disabled",timeout=20000)
    assert ET.parse(sess["merge_files"]["sdf_iso"]).getroot() is not None
    with page.expect_download() as download:
        page.click("#mg-dl-sdf")
    assert download.value.suggested_filename.endswith(".sdf")
    page.screenshot(path=str(tmp_path / "module-h-merged.png"))
    page.set_viewport_size({"width":1100,"height":760})
    settle(page)
    assert page.locator("#cv").bounding_box()["width"] > 500
    drawer = page.locator("#h-drawer").bounding_box()
    assert drawer["x"] >= 0 and drawer["x"]+drawer["width"]<=1100
    page.click("#h-drawer-close")
    page.screenshot(path=str(tmp_path / "module-h-compact.png"))


def test_table_columns_precision_keyboard_and_conversion_readonly(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,"design")
    page.evaluate("async()=>{await __hTest.preview();Object.assign(__mf.design.tables.pipes[0],{dia_src:'review_default',flow_direction:'solver_reference'});__hTest.table();}")
    before=page.evaluate("JSON.stringify(__mf.design.tables)")
    menu(page,"#h-table")
    page.wait_for_selector('#hp-rows tr')
    assert page.locator('#hp-head th').all_text_contents() == ['배관','시작','끝','재질','호칭경 (mm)','내경 (mm)','길이 (m)','C','근거']
    assert page.locator('#hp-rows [data-field=length]').first.input_value() == '1.0603'
    assert page.locator('#hp-rows [data-field=length]').first.get_attribute('title') == str(page.evaluate('__mf.design.tables.pipes[0].length'))
    page.locator('#hp-rows tr[data-label="P2"]').focus()
    page.keyboard.press('Enter')
    assert page.locator('#dg-ins').is_hidden()
    assert page.evaluate('__mf.design.sel') == {'kind':'pipe','label':'P2'}
    assert page.locator('#hp-reference, #hp-source-details').count() == 0
    assert page.evaluate("JSON.stringify(__mf.design.tables)") == before
    page.click('#hp-close')
    page.evaluate("__hTest.setStage('conv')")
    settle(page)
    menu(page,'#h-table')
    assert page.locator('#dg-grid th:visible').all_text_contents() == ['번호','시작 노드','끝 노드','호칭경 (mm)','길이 (m)','종류 / 규격','관경 근거']
    page.click('#h-all-columns')
    assert page.locator('#dg-grid td[data-h-column="length"]').first.inner_text() == str(page.evaluate('__mf.design.tables.pipes[0].length'))
    for table in ['nodes','nozzles','fittings','equipment','pipes']:
        page.select_option('#dg-table',table)
        settle(page)
        assert page.locator('#dg-grid table').count() == 1
    assert page.locator('#h-table-edit').is_hidden()
    page.locator('#dg-grid tr[data-label="P2"]').click()
    assert page.locator('.bore-edit').is_disabled()


def test_workspace_canvas_coordinates_and_screenshots(h_ui):
    page,sess,requests,_=h_ui
    out=ROOT/'data'/'module_h_workspace_20260928'
    out.mkdir(exist_ok=True)
    page.screenshot(path=str(out/'landing.png'))
    seed(page,sess,'design')
    page.evaluate('async()=>{await __hTest.preview();}')
    menu(page,'#h-view')
    page.uncheck('#dg-plan')
    page.keyboard.press('Escape')
    before=len([r for r in requests if r[0]=='POST'])
    for size in [(1500,960),(1100,760)]:
        page.set_viewport_size({'width':size[0],'height':size[1]})
        settle(page)
        cv=page.locator('#cv').bounding_box()
        stage=page.locator('#stage').bounding_box()
        assert cv==stage  # Drawing, pointer hit tests, and gizmo share one origin.
        assert cv['width']==size[0]
        page.screenshot(path=str(out/f'workspace-{size[0]}.png'))
        menu(page,'#h-table')
        assert page.locator('#h-properties').bounding_box()['x']>=0
        page.screenshot(path=str(out/f'table-{size[0]}.png'))
        page.keyboard.press('Escape')
    assert len([r for r in requests if r[0]=='POST'])==before


def test_unconfirmed_loop_and_stale_output_are_not_reported_as_validated(h_ui):
    page,sess,_,_=h_ui
    seed(page,sess,'design')
    page.evaluate("async()=>{await __hTest.preview();__mf.edit.network_mode='loop';__hTest.setStage('design');}")
    settle(page)
    assert '미확정' in page.locator('#h-alert-open').inner_text()
    assert page.locator('#h-action-title').inner_text() == '입력표 준비'
    assert page.locator('#h-action-help').is_hidden()
    page.evaluate("__hTest.setStage('conv');document.querySelector('#dg-stale').textContent='후보 변경';document.querySelector('#dg-stale').classList.remove('hidden');document.querySelector('#btn-download-design').disabled=false")
    settle(page)
    assert page.locator('#h-primary').get_attribute('data-native-action')=='btn-back-design'


def test_changed_calculation_form_requires_rebuild_and_rebuild_stays_accessible(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,'design')
    page.evaluate('async()=>{await __hTest.preview();}')
    menu(page,'#h-task')
    assert page.locator('#dg-build').is_visible()
    before=page.evaluate('JSON.stringify(__mf.design.tables)')
    page.select_option('#dg-sched','KSD 3562')
    settle(page)
    assert '변경사항 반영 필요' in page.locator('#h-alert-open').inner_text()
    assert page.locator('#h-primary').get_attribute('data-native-action')=='dg-build'
    assert page.evaluate('JSON.stringify(__mf.design.tables)')==before
    page.evaluate("__hTest.setStage('conv')")
    settle(page)
    assert page.locator('#h-primary').get_attribute('data-native-action')=='btn-back-design'
