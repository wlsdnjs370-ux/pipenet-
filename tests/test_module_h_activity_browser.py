"""H's real native progress handlers with isolated, controlled job notifications."""
from queue import Queue

from test_module_h_browser import ROOT, h_ui, seed, menu, settle
from src.pipenet_converter.progress import ProgressEvent


def begin(page, phase="평면도 읽기"):
    page.evaluate("phase=>__hTest.busy(true,phase)", phase)
    page.wait_for_function("document.body.classList.contains('h-processing')")


def finish(page):
    page.evaluate("__hTest.stopWatch();__hTest.busy(false);__hTest.say('입력표 준비 완료','ok')")
    page.wait_for_function("document.querySelector('#h-activity-state').textContent==='처리 완료'")


def test_welcome_hides_at_processing_start_before_world_or_next_layout_frame(h_ui):
    page,sess,_,_=h_ui
    assert page.locator("#h-welcome").is_visible()
    # Before Access is chosen, busy still hides the initial architecture.
    assert page.locator("#h-welcome").is_visible()
    visible=page.evaluate("""()=>{
      __hTest.busy(true,'업로드 준비 중');
      return getComputedStyle(document.querySelector('#h-welcome')).display;
    }""")
    assert visible=="none"  # Synchronous CSS guard, without waiting for a frame.
    page.wait_for_function("document.querySelector('#h-welcome').hidden")
    assert page.evaluate("__mf.world") is None
    assert page.locator("#h-activity").is_visible()
    page.evaluate("__hTest.busy(true,'평면도 읽기')")
    settle(page)
    assert page.locator("#h-welcome").is_hidden()
    # A failed/cancelled first upload can be retried from the empty screen.
    page.evaluate("__hTest.busy(false);__hTest.say('읽기 실패','err')")
    page.wait_for_function("!document.querySelector('#h-welcome').hidden")
    assert page.locator("#h-welcome").is_visible()
    begin(page)
    seed(page,sess,"pick")
    finish(page)
    assert page.locator("#h-welcome").is_hidden()


def test_nonmodal_camera_and_menus_work_but_graph_edits_do_not(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,"design")
    page.evaluate("async()=>{await __hTest.preview();}")
    before=page.evaluate("JSON.stringify([__mf.edit,__mf.design.tables])")
    begin(page)
    panel=page.locator("#h-activity").bounding_box()
    assert panel["x"]>1100 and panel["width"]<=360 and panel["y"]==64
    assert page.locator("#busy").evaluate("el=>getComputedStyle(el).position")=="static"
    old_view=page.evaluate("JSON.stringify(__mf.view)")
    page.mouse.move(500,400)
    page.mouse.wheel(0,-300)
    page.wait_for_function("old=>JSON.stringify(__mf.view)!==old",arg=old_view)
    old_view=page.evaluate("JSON.stringify(__mf.view)")
    page.mouse.down(button="right")
    page.mouse.move(550,440,steps=5)
    page.mouse.up(button="right")
    assert page.evaluate("JSON.stringify(__mf.view)")!=old_view
    start=len(requests)
    page.mouse.click(500,400)
    page.keyboard.press("Delete")
    page.keyboard.press("Control+Shift+z")
    menu(page,"#h-task")
    assert page.locator("#dg-sched").get_attribute("aria-disabled")=="true"
    value=page.locator("#dg-sched").input_value()
    page.evaluate("()=>{const n=document.querySelector('#dg-sched');n.value='KSD 3562';n.dispatchEvent(new Event('change',{bubbles:true}));document.querySelector('#dg-build').click();}")
    assert page.locator("#dg-sched").input_value()==value
    page.keyboard.press("Escape")
    menu(page,"#h-table")
    page.select_option("#hp-kind","node")
    page.select_option("#hp-kind","pipe")
    page.locator('#hp-rows tr[data-label="P2"] td').first.click()
    assert page.locator("#h-properties").is_visible()
    assert page.locator("#dg-ins").is_hidden()
    page.click("#hp-close")
    page.keyboard.press("Escape")
    page.click("#h-history")
    page.evaluate("document.querySelector('#steps [data-h-stage=pick]').click()")
    assert page.evaluate("__mf.stage")=="design"
    assert page.evaluate("JSON.stringify([__mf.edit,__mf.design.tables])")==before
    assert not [r for r in requests[start:] if r[0]=="POST"]
    # Native engine remains in charge of disabled properties even after H unlocks.
    page.evaluate("document.querySelector('#dg-build').disabled=true")
    finish(page)
    assert page.locator("#dg-build").is_disabled()
    assert page.locator("#dg-sched").get_attribute("data-h-busy-locked") is None


def test_measured_progress_collapse_and_fresh_next_job(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess)
    begin(page,"도면 압축")
    page.evaluate("__hTest.busy(true,'도면 전송',{pct:0.375})")
    page.wait_for_function("document.querySelector('#busy .track').getAttribute('aria-valuenow')==='37.5'")
    assert "38%" in page.locator("#h-activity-measure").inner_text()
    page.click("#h-activity-toggle")
    assert page.locator("#h-activity-body").is_hidden()
    assert not page.locator("#busy").evaluate("el=>el.classList.contains('hidden')")
    first=page.locator("#h-activity-elapsed").inner_text()
    page.wait_for_function("s=>document.querySelector('#h-activity-elapsed').textContent!==s",arg=first)
    page.evaluate("__hTest.busy(true,'배관 연결 분석');document.querySelector('#log').textContent='배관 420개 분석';")
    page.wait_for_function("document.querySelector('#h-activity-latest').textContent==='배관 420개 분석'")
    assert page.locator("#busy .track").get_attribute("aria-valuenow") is None
    page.click("#h-activity-toggle")
    assert page.locator("#h-activity-measure").inner_text()==""
    assert not [r for r in requests if r==("POST","/api/module-f/job/cancel")]
    finish(page)
    page.click("#h-activity-dismiss")
    begin(page,"두 번째 작업")
    assert "420" not in page.locator("#h-activity-latest").inner_text()
    assert page.locator("#h-activity-log").is_hidden()
    assert "두 번째 작업" in page.locator("#h-activity-events").text_content()
    finish(page)


def test_native_sse_updates_before_completion_without_extra_connections(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess)
    stream=sess["_h_stream_fixture"]=Queue()
    begin(page)
    page.evaluate("window.__done=0;__hTest.watch(()=>{++window.__done;__hTest.say('완료','ok')})")
    page.wait_for_function("!!__mf.es")
    stream.put(("state",dict(state="run",phase="연결 분석",elapsed=1,queued=True)))
    page.wait_for_function("document.querySelector('#h-activity-state').textContent==='순서 대기 중'")
    stream.put(("line","헤드 연결 12개 확인"))
    stream.put(("state",dict(state="run",phase="이음 복원",elapsed=2,queued=False)))
    page.wait_for_function("document.querySelector('#h-activity-latest').textContent==='헤드 연결 12개 확인'")
    page.wait_for_function("document.querySelector('#h-activity-state').textContent==='처리 중'")
    assert page.evaluate("__done")==0
    assert page.locator("#busy").is_visible()
    page.click("#h-more")
    assert page.locator("#h-drawer").is_visible()
    stream.put(("state",dict(state="done",phase="이음 복원",elapsed=3)))
    page.wait_for_function("document.querySelector('#h-activity-state').textContent==='처리 완료'")
    assert page.evaluate("__done")==1
    assert requests.count(("GET","/api/module-f/job/stream"))==1
    assert ("GET","/api/module-f/job") not in requests


def test_polling_fallback_error_is_not_success(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess)
    sess["_h_job_fixture"]=dict(ok=True,state="run",phase="배관 처리",elapsed=1,lines=["배관 7개 처리"],queued=False)
    begin(page)
    page.evaluate("window.__done=0;__hTest.watch(()=>++window.__done)")
    page.wait_for_function("document.querySelector('#h-activity-latest').textContent==='배관 7개 처리'")
    assert page.locator("#h-activity-channel").inner_text()=="주기 확인"
    sess["_h_job_fixture"]=dict(ok=True,state="error",phase="배관 처리",elapsed=2,lines=["연결 데이터 오류"],error="연결 데이터 오류")
    page.wait_for_function("document.querySelector('#h-activity-state').textContent==='처리 실패'")
    assert "연결 데이터 오류" in page.locator("#h-activity-result").inner_text()
    assert page.evaluate("__done")==0
    assert page.locator("#h-activity-details").is_hidden()
    assert not page.locator("body").evaluate("el=>el.classList.contains('h-processing')")


def test_native_cancel_and_new_tab_remain_available(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess)
    begin(page)
    sid=page.evaluate("__mf.sid")
    with page.expect_popup() as popup:
        page.click("#h-other-work")
    other=popup.value
    other.wait_for_function("document.body.classList.contains('h-ready')")
    assert other.evaluate("__mf.sid")!=sid
    assert other.evaluate("window.opener===null")
    other.close()
    assert page.evaluate("__mf.sid")==sid
    page.click("#busy-cancel")
    page.wait_for_function("document.querySelector('#h-activity-state').textContent==='중지됨'")
    assert sess["_h_cancel_body"]["sid"]==sid
    assert sess["_h_cancel_body"]["operation"]
    assert requests.count(("POST","/api/module-f/job/cancel"))==1


def test_h_names_and_compact_notices_keep_validation_details(h_ui):
    page,sess,_,_=h_ui
    assert page.locator("#h-open").inner_text()=="Access"
    seed(page,sess,"design")
    page.evaluate("async()=>{await __hTest.preview();__mf.edit.network_mode='loop';const n=document.querySelector('#dg-review-notice');n.textContent='검토용 루프·그리드 입력입니다. 유량 방향은 미확정입니다.';n.classList.remove('hidden');}")
    settle(page)
    assert page.locator("#h-alert-open").inner_text()=="검토용 · 미확정 ›"
    assert page.locator("#dg-review-notice").is_hidden()
    page.click("#h-history")
    assert page.locator("#steps div").all_text_contents()==["도면 업로드","객체 정의","배관 네트워크 정의","속성 정의","수리계산 입력 변환"]
    page.keyboard.press("Escape")
    page.click("#h-alert-open")
    assert page.locator("#dg-review-notice").is_hidden()
    assert page.locator("#h-action-help").is_hidden()
    assert not page.locator(".h-prose:visible").count()


def test_activity_visuals_on_wide_and_compact_viewports(h_ui):
    page,sess,_,_=h_ui
    seed(page,sess,"design")
    page.evaluate("async()=>{await __hTest.preview();}")
    begin(page,"배관 연결 분석")
    page.evaluate("document.querySelector('#log').textContent='헤드 13개 · 배관 218개\u2014연결 확인 중'")
    settle(page)
    output=ROOT/"data/module_h_workspace_20260928"
    output.mkdir(exist_ok=True)
    for width,height in [(1500,960),(1100,760)]:
        page.set_viewport_size(dict(width=width,height=height))
        settle(page)
        panel=page.locator("#h-activity").bounding_box()
        assert panel["x"]>0 and panel["x"]+panel["width"]<=width
        assert panel["y"]+panel["height"]<height-60
        page.screenshot(path=str(output/f"activity-{width}.png"))
    page.click("#h-activity-toggle")
    menu(page,"#h-table")
    page.screenshot(path=str(output/"activity-table-collapsed.png"))
    finish(page)


def test_live_canvas_batches_are_real_received_data_not_model_mutations(h_ui):
    page,sess,_,_=h_ui
    # The isolated app publishes controlled observations through the actual HTTP route.
    seed(page,sess,"edit")
    before=page.evaluate("JSON.stringify([__mf.world,__mf.edit])")
    begin(page,"연결 탐색")
    op=page.evaluate("__mf.activityOperation")
    # Expose the test server's store only to the Python fixture, never to production JS.
    store=page._h_preview_store
    store.publish(op,ProgressEvent("reset","연결 탐색",{"mode":"graph","roots":[[0,0]]}))
    store.publish(op,ProgressEvent("graph","연결 탐색",{"segments":[[0,0,100,0]],"cursor":[100,0],"count":1}))
    page.wait_for_function("document.querySelector('#h-progress-canvas').dataset.items==='1'")
    assert page.locator("#h-progress-canvas").is_visible()
    assert "미확정" in page.locator("#h-progress-label").inner_text()
    store.publish(op,ProgressEvent("graph","연결 탐색",{"segments":[[100,0,100,100]],"count":2}))
    page.wait_for_function("document.querySelector('#h-progress-canvas').dataset.items==='2'")
    assert page.evaluate("JSON.stringify([__mf.world,__mf.edit])")==before
    store.publish(op,ProgressEvent("reset","레이어 표시 준비",{"mode":"world"}))
    store.publish(op,ProgressEvent("world","레이어 표시 준비",{"items":[{"layer":"ALT-test","kind":"segs","xy":[0,0,100,100]}]}))
    page.wait_for_function("document.querySelector('#h-progress-canvas').dataset.mode==='world'")
    page.screenshot(path=str(ROOT/"data/module_h_workspace_20260928/live-canvas-preview.png"))
    finish(page)
    assert page.locator("#h-progress-canvas").is_hidden()
    # Old-operation updates cannot leak into a new job.
    begin(page,"다음 작업")
    store.publish(op,ProgressEvent("graph","old",{"segments":[[0,0,999,999]]}))
    page.wait_for_timeout(220)
    assert page.locator("#h-progress-canvas").is_hidden()
    finish(page)


def test_browser_request_enables_real_core_observations_only_for_h(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,"edit")
    begin(page,"연결 탐색")
    page.evaluate("async()=>{await __hTest.post('/api/module-f/test-preview-flow',{});}")
    page.wait_for_function("document.querySelector('#h-progress-canvas').dataset.items==='2'")
    assert page.locator("#h-progress-canvas").is_visible()
    assert "연결 탐색" in page.locator("#h-progress-label").inner_text()
    finish(page)
    assert page.locator("#h-progress-canvas").is_hidden()


def test_windows_drag_clamp_reset_and_functional_landing(h_ui):
    page,sess,requests,_=h_ui
    assert page.locator(".h-conversion-mark").is_visible()
    assert page.locator("#h-welcome h2,#h-welcome p").count()==0
    assert "도면을 열고" not in page.locator("#h-welcome").inner_text()
    page.click("#h-primary")
    page.click('#h-access-plan')
    page.wait_for_function('!document.querySelector("#h-drawer").hidden')
    original=page.locator("#h-drawer").bounding_box()
    head=page.locator(".h-drawer-head").bounding_box()
    page.mouse.move(head["x"]+90,head["y"]+20)
    page.mouse.down();page.mouse.move(head["x"]+390,head["y"]+70,steps=6);page.mouse.up()
    moved=page.locator("#h-drawer").bounding_box()
    assert moved["x"]>original["x"]+250
    page.set_viewport_size({"width":800,"height":640})
    settle(page)
    box=page.locator("#h-drawer").bounding_box()
    assert box["x"]>=0 and box["x"]+box["width"]<=800
    assert box["y"]>=0 and box["y"]+box["height"]<=640
    page.locator(".h-drawer-head").dblclick(position={"x":70,"y":20})
    assert page.locator("#h-drawer").get_attribute("data-h-dragged") is None
    page.keyboard.press("Escape")
    seed(page,sess)
    begin(page)
    head=page.locator(".h-activity-head")
    head.focus();page.keyboard.press("ArrowLeft")
    assert page.locator("#h-activity").get_attribute("data-h-dragged")=="true"
    assert not [r for r in requests if r[0]=="POST"]
    finish(page)


def test_property_table_editor_and_native_dialogs_move_without_data_changes(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,"design")
    menu(page,"#h-view")
    page.uncheck("#dg-plan")
    page.click("#h-drawer-close")
    page.evaluate("async()=>{await __hTest.preview();__hTest.select('pipe','P2');}")
    menu(page,"#h-table")
    page.wait_for_selector('#h-properties:not([hidden]) #hp-rows tr')
    assert page.locator('#dg-ins').is_hidden()
    before=page.evaluate("JSON.stringify([__mf.edit,__mf.design.tables])")
    start=len(requests)
    for panel,head,show in [
        ("#h-properties","#h-properties > header",None),
        ("#ne-panel",".ne-pop-head","document.querySelector('#ne-panel').classList.remove('hidden')"),
        ("#ne-history-dialog","#ne-history-dialog > h3","document.querySelector('#ne-history-dialog').showModal()"),
        ("#conv-modal .sheet","#conv-modal .sheet > header","document.querySelector('#conv-modal').classList.remove('hidden')"),
    ]:
        if show:
            page.evaluate(show)
        assert page.locator(panel).is_visible()
        old=page.locator(panel).bounding_box()
        handle=page.locator(head)
        handle.focus();page.keyboard.press("ArrowLeft");page.keyboard.press("ArrowDown")
        assert page.locator(panel).get_attribute("data-h-dragged")=="true"
        new=page.locator(panel).bounding_box()
        assert old!=new
        assert new["x"]>=0 and new["y"]>=0
        if panel=="#h-properties":
            page.click('#hp-close')
        elif panel=="#ne-history-dialog":
            page.locator("#ne-history-dialog button").click()
        elif panel=="#conv-modal .sheet":
            page.locator("#conv-cancel").click()
        else:
            page.evaluate("selector=>document.querySelector(selector).classList.add('hidden')",panel)
    assert page.evaluate("JSON.stringify([__mf.edit,__mf.design.tables])")==before
    assert not [r for r in requests[start:] if r[0]=="POST"]
