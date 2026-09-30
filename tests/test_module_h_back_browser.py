"""H process-back navigation is separate from graph or native text undo."""
from copy import deepcopy

import pytest

from test_module_h_browser import h_ui, seed, settle


def focus_canvas_page(page):
    page.evaluate('document.activeElement?.blur()')


def test_button_position_and_plan_process_sequence_without_edits(h_ui):
    page,sess,requests,tmp=h_ui
    assert page.locator('#h-back').is_disabled()
    assert page.locator('#h-back').evaluate("e=>e.nextElementSibling.id")=='h-history'
    assert page.locator('#h-back').get_attribute('aria-keyshortcuts') is None
    seed(page,sess,'edit')
    before=page.evaluate('JSON.stringify({edit:__mf.edit,world:__mf.world,zones:__mf.zones})')
    posts=[r for r in requests if r[0]=='POST']
    assert '객체 정의' in page.locator('#h-back').get_attribute('title')
    assert 'Ctrl+Z' not in page.locator('#h-back').get_attribute('title')
    page.screenshot(path=str(tmp/'module-h-back.png'))
    page.click('#h-back')
    page.wait_for_function("__mf.stage==='pick'")
    page.click('#h-back')
    page.wait_for_function("__mf.stage==='open'")
    assert page.locator('#h-back').is_disabled()
    page.keyboard.press('Control+z');settle(page)
    assert page.evaluate('__mf.stage')=='open'
    assert before==page.evaluate('JSON.stringify({edit:__mf.edit,world:__mf.world,zones:__mf.zones})')
    assert posts==[r for r in requests if r[0]=='POST']


def test_output_back_to_attributes_then_network_preserves_tables(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,'conv');before=deepcopy(sess['design']['tables'])
    page.click('#h-back')
    page.wait_for_function("__mf.stage==='design' && !document.querySelector('#h-back').disabled")
    page.click('#h-back');page.wait_for_function("__mf.stage==='edit'")
    assert sess['design']['tables']==before
    assert not [path for method,path in requests if method=='POST']


@pytest.mark.parametrize('entry',['edit','conv'])
def test_merge_returns_to_its_entry_not_a_graph_undo(h_ui,entry):
    page,sess,requests,_=h_ui
    seed(page,sess,entry)
    page.evaluate("__hTest.setStage('merge')");settle(page)
    page.click('#h-back')
    page.wait_for_function("entry=>__mf.stage===entry",arg=entry)
    assert not [path for method,path in requests if method=='POST']


@pytest.mark.parametrize('slot',['system','system2','machineroom'])
def test_subdrawing_back_is_local_to_the_current_slot(h_ui,slot):
    page,sess,requests,_=h_ui
    seed(page,sess,'edit')
    # Slot changes must discard a previous plan/merge return destination.
    page.evaluate("__hTest.setStage('merge')");settle(page)
    page.evaluate("slot=>{__mf.slot=slot;__hTest.setStage('sub');}",slot);settle(page)
    page.click('#h-back');page.wait_for_function("__mf.stage==='open'")
    assert page.evaluate('__mf.slot')==slot
    assert not [path for method,path in requests if method=='POST']


def test_text_undo_and_busy_repeat_guards(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,'edit')
    page.evaluate('''()=>{
      const input=document.createElement('textarea');input.id='text-undo-probe';
      input.style.cssText='position:fixed;top:80px;left:20px;z-index:9999';
      document.body.append(input);
      window.addEventListener('keydown',e=>{if(e.key==='z')window.__textUndoBlocked=e.defaultPrevented;});
    }''')
    page.locator('#text-undo-probe').fill('before')
    page.locator('#text-undo-probe').press('End')
    page.locator('#text-undo-probe').press_sequentially(' after')
    page.keyboard.press('Control+z')
    assert page.evaluate('__mf.stage')=='edit'
    assert page.evaluate('__textUndoBlocked') is False
    # Chromium can coalesce successive keystrokes into different undo chunks.
    # Verify native editing occurred, not a browser-specific chunk boundary.
    undone=page.locator('#text-undo-probe').input_value()
    assert undone!='before after' and 'before after'.startswith(undone)
    focus_canvas_page(page)
    page.evaluate("document.body.dispatchEvent(new KeyboardEvent('keydown',{key:'z',ctrlKey:true,repeat:true,bubbles:true,cancelable:true}))")
    assert page.evaluate('__mf.stage')=='edit'
    page.evaluate("__hTest.busy(true,'테스트 처리 중')");settle(page)
    assert page.locator('#h-back').is_disabled()
    page.keyboard.press('Control+z')
    assert page.evaluate('__mf.stage')=='edit'
    page.evaluate('__hTest.busy(false)');settle(page)
    page.click('#h-back');page.wait_for_function("__mf.stage==='pick'")
    assert not [path for method,path in requests if method=='POST']


def test_auto_flow_returns_to_upload_without_touching_graph(h_ui):
    page,sess,requests,_=h_ui
    seed(page,sess,'edit')
    page.evaluate("__mf.method='auto';__hTest.setStage('auto')");settle(page)
    page.click('#h-back');page.wait_for_function("__mf.stage==='open'")
    assert page.evaluate('__mf.method')=='auto'
    assert not [path for method,path in requests if method=='POST']


@pytest.mark.parametrize('modal',['native','conversion'])
def test_ctrl_z_does_not_navigate_under_a_modal(h_ui,modal):
    page,sess,requests,_=h_ui
    seed(page,sess,'edit')
    page.evaluate('''kind=>{
      if(kind==='native') {
        const dialog=document.createElement('dialog');dialog.id='back-modal-probe';
        dialog.innerHTML='<button>확인</button>';document.body.append(dialog);dialog.showModal();
      } else {
        document.querySelector('#conv-modal').classList.remove('hidden');
        document.querySelector('#conv-cancel').focus();
      }
    }''',modal)
    page.keyboard.press('Control+z');settle(page)
    assert page.evaluate('__mf.stage')=='edit'
    assert not [path for method,path in requests if method=='POST']
