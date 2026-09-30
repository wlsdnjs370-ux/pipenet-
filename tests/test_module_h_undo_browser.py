"""Ctrl+Z undoes edits, independently of H's process-navigation button."""
import pytest

from test_module_h_browser import h_ui, seed, settle
from test_module_h_bottom_properties_browser import prepare
from test_module_f_merge_editor_followup import combined


def unfocus(page):
    """Send application shortcuts outside browser-owned form controls."""
    page.evaluate('document.activeElement?.blur()')


def wait_pipe(page, label, diameter, c):
    """Wait for the canonical editor and its transaction to finish refreshing."""
    page.wait_for_function('''({label,diameter,c})=>{
      const pipe=moduleFNetworkEditor.getState()?.pipes.find(r=>r.label===label);
      return pipe?.dia===diameter && pipe?.c===c && document.querySelector('#busy').classList.contains('hidden');
    }''', arg=dict(label=label, diameter=diameter, c=c))


@pytest.mark.parametrize('scope', ['design', 'merge'])
def test_keyboard_undo_restores_one_property_at_a_time_without_leaving_stage(h_ui, monkeypatch, scope):
    page, sess, _ = prepare(h_ui, monkeypatch)
    if scope == 'merge':
        combined(sess)
        seed(page, sess, 'merge')
        page.evaluate('async()=>{await __hTest.mergeState();await __hTest.merged();await moduleFNetworkEditor.refresh();}')
    page.evaluate('ModuleHProperties.open()')
    page.wait_for_selector('#hp-rows tr')
    original = page.evaluate('moduleFNetworkEditor.getState().pipes.find(r=>r.dia===40)')
    label, diameter, c = original['label'], original['dia'], original['c']
    row = page.locator(f'#hp-rows tr[data-label="{label}"]')
    actions = []
    page.on('request', lambda req: actions.append(req.post_data_json['action'])
            if req.method == 'POST' and req.url.endswith('/api/module-f/network-editor') else None)
    row.locator('[data-field=dn]').select_option('80')
    wait_pipe(page, label, 80, c)
    row.locator('[data-field=c]').fill('135')
    row.locator('[data-field=c]').press('Tab')
    wait_pipe(page, label, 80, 135)
    # Selection is not an edit: undo the last change even after another row is selected.
    page.locator(f'#hp-rows tr:not([data-label="{label}"]) td').first.click()
    unfocus(page)
    page.keyboard.press('Control+z')
    wait_pipe(page, label, 80, c)
    page.wait_for_function('''({label,c})=>Number(document.querySelector(
      `#hp-rows tr[data-label="${label}"] [data-field=c]`)?.value)===c''', arg=dict(label=label, c=c))
    assert page.evaluate('__mf.stage') == scope
    # H's property dock need not be open for the same editing history to work.
    page.click('#hp-close')
    unfocus(page)
    page.keyboard.press('Control+z')
    wait_pipe(page, label, diameter, c)
    assert page.evaluate('__mf.stage') == scope
    assert page.locator('#h-properties').is_hidden()
    assert actions == ['preview', 'apply', 'preview', 'apply', 'undo', 'undo']
    before = list(actions)
    page.keyboard.press('Control+z')
    page.wait_for_function("document.querySelector('#status').textContent.includes('되돌릴 것이 없습니다')")
    assert actions == before
    assert page.evaluate('__mf.stage') == scope
    # Existing network redo still replays exactly one change.
    page.keyboard.press('Control+Shift+z')
    wait_pipe(page, label, 80, c)
    assert actions == before + ['redo']
    assert page.evaluate('__mf.stage') == scope
    actual = sess['design']['tables'] if scope == 'design' else sess['merged']['combined']
    pipe = next(r for r in actual.pipes if str(r['label']) == label)
    assert pipe['dia'] == 80 and pipe['c'] == c


@pytest.mark.parametrize(('stage', 'undo_id'), [('pick', 'pk-undo'), ('edit', 'ed-undo')])
def test_shortcut_dispatches_once_to_native_edit_undo(h_ui, stage, undo_id):
    page, sess, requests, _ = h_ui
    seed(page, sess, stage)
    # Observe dispatch only; these bare fixtures do not contain a pick/edit history.
    page.evaluate('''id=>{window.undoCalls=0;const button=document.getElementById(id);
      button.disabled=false;button.onclick=()=>window.undoCalls++;}''', undo_id)
    unfocus(page)
    for key in ('Control+z', 'Meta+z'):
        page.keyboard.press(key)
    assert page.evaluate('undoCalls') == 2
    assert page.evaluate('__mf.stage') == stage
    assert not [path for method, path in requests if method == 'POST']


def test_region_undo_uses_region_history_not_process_navigation(h_ui):
    page, sess, _, _ = h_ui
    seed(page, sess, 'edit')
    page.evaluate('''()=>{
      __mf.zones=[[10,10,50,50]];document.querySelector('#ed-zone-arm').checked=true;
      window.zoneUndoCalls=0;window.editUndoCalls=0;
      const zone=document.querySelector('#ed-zone-undo');zone.disabled=false;
      zone.onclick=()=>window.zoneUndoCalls++;
      document.querySelector('#ed-undo').onclick=()=>window.editUndoCalls++;
    }''')
    unfocus(page)
    page.keyboard.press('Control+z')
    assert page.evaluate('[zoneUndoCalls,editUndoCalls,__mf.stage]') == [1, 0, 'edit']


def test_subdrawing_point_undo_keeps_current_process(h_ui):
    page, sess, requests, _ = h_ui
    seed(page, sess, 'edit')
    page.evaluate("__mf.slot='system';__hTest.setStage('sub');__mf.sub.arm=0")
    box = page.locator('#cv').bounding_box()
    page.mouse.click(box['x'] + box['width'] / 2, box['y'] + box['height'] / 2)
    assert page.evaluate('__mf.sub.picks[0]') is not None
    assert page.evaluate('__mf.undo.length') == 1
    unfocus(page)
    page.keyboard.press('Control+z')
    assert page.evaluate('[__mf.stage,__mf.sub.picks,__mf.undo.length]') == ['sub', [None, None], 0]
    page.keyboard.press('Control+z')
    assert page.evaluate('__mf.stage') == 'sub'
    assert not [path for method, path in requests if method == 'POST']


def test_repeat_busy_and_modal_keys_do_not_consume_edit_history(h_ui, monkeypatch):
    page, _, _ = prepare(h_ui, monkeypatch)
    page.evaluate('ModuleHProperties.open()')
    page.wait_for_selector('#hp-rows tr')
    original = page.evaluate('moduleFNetworkEditor.getState().pipes.find(r=>r.dia===40)')
    label = original['label']
    page.locator(f'#hp-rows tr[data-label="{label}"] [data-field=dn]').select_option('80')
    wait_pipe(page, label, 80, original['c'])
    unfocus(page)
    before = page.evaluate('moduleFNetworkEditor.getState().revision')
    page.evaluate("document.body.dispatchEvent(new KeyboardEvent('keydown',{key:'z',ctrlKey:true,repeat:true,bubbles:true,cancelable:true}))")
    page.evaluate("__hTest.busy(true,'테스트 처리 중')")
    page.keyboard.press('Control+z')
    page.evaluate('__hTest.busy(false)')
    page.evaluate('''()=>{const dialog=document.createElement('dialog');dialog.id='undo-modal';
      dialog.innerHTML='<button>확인</button>';document.body.append(dialog);dialog.showModal();}''')
    page.keyboard.press('Control+z')
    settle(page)
    assert page.evaluate('moduleFNetworkEditor.getState().revision') == before
    assert page.evaluate('__mf.stage') == 'design'
    page.evaluate("document.querySelector('#undo-modal').close()")
    unfocus(page)
    page.keyboard.press('Control+z')
    wait_pipe(page, label, 40, original['c'])
    assert page.evaluate('__mf.stage') == 'design'


@pytest.mark.parametrize('stage', ['open', 'conv', 'auto'])
def test_empty_history_never_returns_to_previous_process(h_ui, stage):
    page, sess, requests, _ = h_ui
    seed(page, sess, stage)
    unfocus(page)
    page.keyboard.press('Control+z')
    page.wait_for_function("document.querySelector('#status').textContent.includes('되돌릴 것이 없습니다')")
    assert page.evaluate('__mf.stage') == stage
    assert not [path for method, path in requests if method == 'POST']
