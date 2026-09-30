"""Five-stage F presentation with real handlers and isolated engine sessions."""
from pathlib import Path
from urllib.parse import urlsplit
import re
import xml.etree.ElementTree as ET

import pytest

from test_module_h_browser import h_ui, seed  # noqa: F401 - shared isolated HTTP server
from test_module_f_viewport_browser import ROOT, settle


@pytest.fixture()
def f_ui(h_ui):
    page, sess, requests, tmp_path = h_ui
    base = urlsplit(page.url)
    origin = f"{base.scheme}://{base.netloc}"
    source = (ROOT / "static/module_f.js").read_text(encoding="utf-8")
    source = source.replace('  setStage("open");\n  loadSaved();', '''
  window.__hTest={setStage,draw,preview:designPreview,select:insSelect,merged:loadMergeView};
  window.__layoutTest={table:renderDesignTable,stale:renderStale,settings:designSettings,sync:syncDesignForMethod};
  window.__nativeControls=Object.fromEntries(['ed-flow','ed-worst','ed-save','dg-build','btn-convert','mg-build','mg-emit','pk-next'].map(id=>[id,[$(id),$(id).onclick]]));
  window.__nativeSettings=JSON.stringify(designSettings());
  setStage("open");
  loadSaved();''')
    page.route(origin + "/static/module_f.js*", lambda route: route.fulfill(body=source, content_type="application/javascript"))
    page.route(origin + "/module-f", lambda route: route.fulfill(
        body=(ROOT / "templates/module_f.html").read_text(encoding="utf-8"), content_type="text/html"))
    page.goto(origin + "/module-f")
    page.wait_for_function("document.body.classList.contains('mf-clean')")
    assert page.locator("#mf-layout-error").count() == 0
    settle(page)
    yield page, sess, requests, tmp_path


def model(page):
    return page.evaluate("JSON.stringify([__mf.world,__mf.edit,__mf.design?.tables,__mf.design?.view])")


def test_five_steps_and_all_native_controls_preserved(f_ui):
    page, _, _, _ = f_ui
    assert page.locator("#steps > div").all_text_contents() == ["도면 열기", "찍기", "손질", "수리계산", "수리계산 입력 변환"]
    assert page.locator('#steps [aria-disabled="true"]').count() == 4
    assert page.evaluate("JSON.stringify(__layoutTest.settings())===__nativeSettings")
    assert page.evaluate("Object.entries(__nativeControls).every(([id,[node,handler]])=>node===document.getElementById(id)&&node.onclick===handler)")
    for control in re.findall(r'\bid="([^"]+)"', (ROOT / "templates/module_f.html").read_text(encoding="utf-8")):
        assert page.locator(f'[id="{control}"]').count() == 1, control


def test_edit_advanced_keeps_values_graph_and_action_handlers(f_ui):
    page, sess, requests, _ = f_ui
    seed(page, sess)
    before = model(page)
    assert page.locator("#ed-flow").is_visible()
    assert page.locator("#ed-save").is_visible()
    assert page.locator("#ed-worst").is_visible()
    assert page.locator("#ed-network-mode").is_hidden()
    assert page.locator("#ed-aj-scan").is_hidden()
    start = len(requests)
    page.locator("#mf-advanced-edit > summary").click()
    assert page.locator("#ed-network-mode").is_visible()
    page.get_by_text("끊긴 배관 자동 이음", exact=True).click()
    assert page.locator("#ed-aj-scan").is_visible()
    page.locator("#mf-advanced-edit > summary").click()
    assert model(page) == before
    assert not any(method == "POST" for method, _ in requests[start:])
    assert page.evaluate("Object.entries(__nativeControls).every(([id,[node,handler]])=>node===document.getElementById(id)&&node.onclick===handler)")


def test_wide_table_readability_fields_and_original_row_selection(f_ui):
    page, sess, _, _ = f_ui
    seed(page, sess, "design")
    page.evaluate("async()=>{await __hTest.preview();}")
    page.evaluate("()=>{Object.assign(__mf.design.tables.pipes[0],{flow_direction:'solver_reference',group:'Unset',dia_src:'review_default'});__layoutTest.table();}")
    settle(page)
    before = model(page)
    page.click("#mf-table")
    assert page.locator('#dg-grid td[data-mf-column="length"]').first.inner_text() == '1.06'
    assert page.locator('#mf-table-warning').is_hidden()
    assert page.locator(".side").is_hidden()
    assert page.locator("#dg-grid").bounding_box()["width"] > 1100
    assert page.locator('#dg-grid th[data-mf-column="flow_direction"]').is_hidden()
    assert page.locator('#dg-grid th[data-mf-column="group"]').is_hidden()
    assert "검토용 · 미확정" in page.locator("#dg-grid").inner_text()
    assert page.locator('#dg-grid th:visible').all_text_contents() == ["번호", "시작 노드", "끝 노드", "호칭경 (mm)", "길이 (m)", "종류 / 규격", "관경 근거"]
    assert page.locator('#dg-grid td[data-mf-column="length"]').first.evaluate("n=>getComputedStyle(n).whiteSpace") == "nowrap"
    page.click("#mf-all-columns")
    assert page.locator('#dg-grid td[data-mf-column="length"]').first.inner_text() == str(page.evaluate('__mf.design.tables.pipes[0].length'))
    assert page.locator('#dg-grid th[data-mf-column="flow_direction"]').is_visible()
    assert "기준 방향 · 유향 미확정" in page.locator("#dg-grid").inner_text()
    page.click("#mf-all-columns")
    page.locator('#dg-grid tr[data-label="P2"]').click()
    assert page.locator("#dg-ins").is_visible()
    assert page.evaluate("__mf.design.hilite.has('P2')")
    assert model(page) == before
    page.click("#dg-ins-close")
    page.click("#mf-network")
    assert page.locator(".side").is_visible()
    # Table is available in conversion too; changing the view never generates files.
    page.evaluate("__hTest.setStage('conv')")
    settle(page)
    page.click("#mf-table")
    assert page.locator("#dg-table").is_visible()
    assert page.locator("#mf-table-build").is_hidden()
    assert page.locator("#dg-bore-row").is_hidden()
    page.locator('#dg-grid tr[data-label="P2"]').click()
    assert page.locator("#dg-ins").is_visible()
    assert page.locator('.bore-edit').is_disabled()
    assert page.locator('.ov-save:visible').count() == 0


def test_required_review_inputs_stale_warning_and_five_table_types(f_ui):
    page, sess, _, _ = f_ui
    seed(page, sess, "design")
    page.evaluate("async()=>{await __hTest.preview(); __mf.edit.network_mode='loop';__layoutTest.sync();__layoutTest.stale({why:['후보 변경']});}")
    settle(page)
    assert page.locator("#dg-review-dia").is_visible()
    assert page.locator("#dg-review-datum").is_visible()
    assert page.locator("#dg-build").is_visible()
    assert "undefined" not in page.locator("#dg-stale").text_content()
    assert page.locator("#dg-stale").is_hidden()  # Full explanation is folded, not removed.
    assert page.get_by_text("입력표 갱신 필요 · 표 확정", exact=True).is_visible()
    page.click("#mf-table")
    assert "입력표 갱신 필요" in page.locator("#mf-table-warning").inner_text()
    for table in ["nodes", "nozzles", "fittings", "equipment", "pipes"]:
        page.select_option("#dg-table", table)
        settle(page)
        assert page.locator("#dg-table").input_value() == table
    page.evaluate("document.querySelector('#dg-build').disabled=true")
    settle(page)
    assert page.locator("#mf-table-build").is_disabled()


def test_large_table_scroll_keeps_headers_and_precision_without_model_change(f_ui):
    page, sess, _, _ = f_ui
    seed(page, sess, "design")
    page.evaluate("""async()=>{
      await __hTest.preview();
      __mf.design.tables.pipes=Array.from({length:200},(_,i)=>({label:`T${i}`,in:String(i),out:String(i+1),dia:65,length:6.23456789,type:'KSD 3507',dia_src:'review_default',c:120}));
      __layoutTest.table();
    }""")
    settle(page)
    before=model(page)
    page.click('#mf-table')
    assert page.locator('#dg-grid tbody tr').count()==200
    assert page.locator('#dg-grid td[data-mf-column="length"]').first.inner_text()=='6.235'
    assert page.locator('#dg-grid td[data-mf-column="length"]').first.get_attribute('title')=='원래 값: 6.23456789'
    grid=page.locator('#dg-grid')
    header=page.locator('#dg-grid th:visible').first
    y=header.bounding_box()['y']
    grid.evaluate('n=>n.scrollTop=1200')
    settle(page)
    assert header.bounding_box()['y']==pytest.approx(y,abs=1)
    page.click('#mf-all-columns')
    assert model(page)==before


def test_table_open_is_closed_on_stage_or_slot_change(f_ui):
    page, sess, _, _ = f_ui
    seed(page, sess, "design")
    page.click("#mf-table")
    page.evaluate("__hTest.setStage('edit')")
    settle(page)
    assert page.locator("#mf-table-pane").is_hidden()
    assert page.locator("#ed-flow").is_visible()
    page.evaluate("__mf.slot='system';__hTest.setStage('sub')")
    settle(page)
    assert page.locator("#steps > div").all_text_contents() == ["도면 열기", "경로 추출"]
    assert page.locator("#sub-extract").is_visible()
    assert page.locator("#mf-table").is_hidden()


def test_real_edit_undo_and_merged_sdf_survive_layout(f_ui):
    page, sess, requests, tmp_path = f_ui
    seed(page, sess, "design")
    page.uncheck("#dg-plan")
    page.evaluate("async()=>{await __hTest.preview();}")
    page.click("#mf-table")
    page.locator('#dg-grid tr[data-label="P2"]').click()
    page.wait_for_function("document.querySelector('#ne-material').options.length>0")
    page.click(".bore-edit")
    page.wait_for_function("document.querySelector('#ne-op').value==='pipe'")
    page.select_option("#ne-dn", "32")
    page.fill("#ne-note", "F 화면 정리 회귀 검증")
    page.click("#ne-apply")
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===32 && document.querySelector('#busy').classList.contains('hidden')")
    assert sess["design"]["tables"].pipes[1]["dia"] == 32
    # Editing toolbar is on the drawing, so return to that view to undo.
    page.click("#mf-network")
    page.click("#ne-undo")
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===25 && document.querySelector('#busy').classList.contains('hidden')")
    assert ("POST", "/api/module-f/network-editor") in requests
    from test_module_f_merge import _riser
    from routes.module_f import api_merge
    sess["slots"]["system"]["riser"] = _riser()
    sess["supply_mode"] = "lsp_gravity"
    api_merge.rebuild_merged(sess)
    for pipe in sess["merged"]["combined"].pipes:
        pipe.setdefault("c", 120)
    page.click("#btn-merge")
    page.wait_for_function("__mf.stage==='merge' && !!__mf.mergeView?.pipes.length")
    page.click("#mg-emit")
    page.wait_for_function("!document.querySelector('#mg-dl-sdf').disabled", timeout=20000)
    assert ET.parse(sess["merge_files"]["sdf_iso"]).getroot() is not None
    assert Path(sess["merge_files"]["sdf_iso"]).is_relative_to(tmp_path.parent / "module_f_merged" / sess["id"])
    with page.expect_download() as download:
        page.click("#mg-dl-sdf")
    assert download.value.suggested_filename.endswith(".sdf")


def test_isolated_screenshots_and_smaller_desktop(f_ui):
    page, sess, _, _ = f_ui
    out = ROOT / "data" / "module_f_ui_20260928"
    out.mkdir(exist_ok=True)
    seed(page, sess)
    page.screenshot(path=str(out / "edit.png"))
    seed(page, sess, "design")
    page.evaluate("async()=>{await __hTest.preview();}")
    page.click("#mf-table")
    page.screenshot(path=str(out / "table.png"))
    page.set_viewport_size({"width":1100,"height":760})
    settle(page)
    assert page.locator("#dg-grid").bounding_box()["width"] > 1000
    page.click("#mf-network")
    page.click("#mf-settings")
    assert page.locator("#dg-fx").is_visible()
    page.screenshot(path=str(out / "advanced.png"))
