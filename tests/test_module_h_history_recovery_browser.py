"""The failed input-preview path provides a safe, visible history recovery."""
from copy import deepcopy

from test_module_h_browser import h_ui, ROOT
from test_module_h_bottom_properties_browser import prepare
from test_module_h_undo_browser import wait_pipe
from routes.module_f import network_edit as ne, api_network_edit


def test_preview_conflict_recovery_cancel_then_accept_preserves_history(h_ui, monkeypatch):
    page, sess, tmp = prepare(h_ui, monkeypatch)
    api_network_edit.install_guards(page._h_app)
    page.evaluate('ModuleHProperties.open()')
    page.wait_for_selector('#hp-rows tr')
    original = page.evaluate('moduleFNetworkEditor.getState().pipes.find(r=>r.dia===40)')
    label = original['label']
    page.locator(f'#hp-rows tr[data-label="{label}"] [data-field=dn]').select_option('80')
    wait_pipe(page, label, 80, original['c'])
    page.click('#hp-close')
    path = ne._path(sess, 'design')
    old = path.read_bytes()
    sess['design'] = deepcopy(ne.ensure(sess, 'design')['base_object'])
    sess['design']['tables'].pipes[0]['length'] += .1
    ne.accept_rebuilt(sess)
    assert ne.ensure(sess, 'design')['conflict']
    # The guarded preview fails as in the report, but now exposes recovery even
    # with the table closed. No obsolete edit-toolbar controls are needed.
    page.evaluate("__hTest.busy(true,'입력값 반영 중')")
    page.wait_for_function("document.body.classList.contains('h-processing')")
    result = page.evaluate('async()=>{try{await __hTest.preview();return null;}catch(e){__hTest.say(e.message,"err");return e.message;}finally{__hTest.busy(false);}}')
    assert '이전 편집' in result
    page.wait_for_selector('#ha-accept-basis:visible')
    page.wait_for_function("document.querySelector('#h-activity').dataset.state==='review'")
    assert page.locator('#status').get_attribute('class') == 'warn'
    assert page.locator('#h-activity-state').inner_text() == '이전 편집 확인 필요'
    # Regression: the user opened processing history to investigate the error;
    # that drawer used to hide the ONLY recovery button with visibility:hidden.
    page.click('#h-history')
    page.wait_for_function("document.body.classList.contains('h-drawer-open')")
    assert page.locator('#ha-accept-basis').is_visible()
    assert page.locator('#h-primary').inner_text() == '편집 기록 확인'
    start = len(h_ui[2])
    page.click('#h-primary')
    assert page.locator('#ha-accept-basis').evaluate('e=>document.activeElement===e')
    # A stale old frontend snapshot must not cause repeated rebuilds while the
    # server holds a manual edit. UI navigation/inspection must not archive it.
    page.evaluate('__mf.ovDirty=true')
    page.wait_for_timeout(800)
    assert not [p for method,p in h_ui[2][start:] if method=='POST']
    assert path.read_bytes() == old
    page.evaluate('__mf.ovDirty=false')
    page.click('#h-history')
    assert page.locator('#ha-accept-basis').is_visible()
    page.screenshot(path=str(tmp / 'history-basis-recovery.png'))
    output = ROOT/'outputs/h_history_recovery_review'
    output.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(output/'recovery-with-history-open.png'))
    # The same conflict must never hide an unrelated computation/transport error.
    page.evaluate("__hTest.say('실제 서버 처리 실패','err')")
    page.wait_for_function("document.querySelector('#h-activity').dataset.state==='error'")
    assert page.locator('#status').get_attribute('class') == 'err'
    page.once('dialog', lambda dialog: dialog.dismiss())
    page.click('#ha-accept-basis')
    page.wait_for_function("!document.querySelector('#ha-accept-basis').disabled")
    assert path.read_bytes() == old
    assert ne.ensure(sess, 'design')['conflict']
    page.once('dialog', lambda dialog: dialog.accept())
    page.click('#ha-accept-basis')
    page.wait_for_function("!moduleFNetworkEditor.getState()?.conflict && document.querySelector('#ha-history-conflict').hidden && document.querySelector('#busy').classList.contains('hidden')")
    assert page.evaluate('__mf.stage') == 'design'
    assert any(p.read_bytes() == old for p, _ in ne._history_versions(sess, 'design'))
    assert ne.ensure(sess, 'design')['current'].pipes[label].row['dia'] == 40
    page.evaluate('ModuleHProperties.open()')
    page.wait_for_selector('#hp-rows tr')
    page.locator(f'#hp-rows tr[data-label="{label}"] [data-field=dn]').select_option('80')
    wait_pipe(page, label, 80, original['c'])
    page.evaluate('document.activeElement?.blur()')
    page.keyboard.press('Control+z')
    wait_pipe(page, label, 40, original['c'])


def test_history_recovery_does_not_leak_to_other_drawing_or_stage(h_ui, monkeypatch):
    page,sess,_=prepare(h_ui,monkeypatch)
    # A stale response cached by the editor cannot hold a different drawing.
    page.evaluate('''()=>{
      const old=moduleFNetworkEditor.getState;
      window.__restoreEditorState=()=>{moduleFNetworkEditor.getState=old;};
      moduleFNetworkEditor.getState=()=>({sid:'other-drawing',scope:'design',conflict:'old',undo:2,can_accept_basis:true});
      __mf.ovDirty=false;
    }''')
    page.wait_for_timeout(400)
    assert page.locator('#ha-history-conflict').is_hidden()
    assert page.evaluate('ModuleHAttributes.historyConflict()') is None
    page.evaluate("moduleFNetworkEditor.getState=()=>({sid:__mf.sid,scope:'design',conflict:'old',undo:2,can_accept_basis:true})")
    page.wait_for_selector('#ha-accept-basis:visible')
    page.evaluate("moduleFNetworkEditor.getState=()=>({sid:__mf.sid,scope:'design',conflict:'old',undo:2,can_accept_basis:false,basis_stale:{why:['선정 변경']}})")
    page.wait_for_selector('#ha-rebuild-basis:visible')
    assert page.locator('#ha-accept-basis').is_disabled()
    assert '현재 선택으로 기준망' in page.locator('#ha-recovery-message').inner_text()
    page.evaluate("__hTest.setStage('edit')")
    page.wait_for_function("document.querySelector('#ha-history-conflict').hidden")
    page.evaluate('__restoreEditorState()')
