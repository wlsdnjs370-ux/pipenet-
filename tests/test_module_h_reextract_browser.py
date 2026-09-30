"""H removes only the duplicate re-extraction row, not the extraction action."""
from test_module_h_browser import h_ui, seed, settle


def test_changed_selection_keeps_single_extract_action(h_ui):
    page, sess, requests, tmp = h_ui
    seed(page, sess, 'edit')
    page.evaluate('''()=>{
      __mf.zones=[[10,10,50,50]];__mf.edit.sources=[0];
      __mf.edit.worst={selection_mode:'area_all',zones:__mf.zones,corridor:[],heads:[]};
      __mf.edit.edits_since_worst=3;__hTest.renderEdit();
    }''')
    settle(page)
    page.click('#h-primary')
    assert page.locator('#h-drawer').is_visible()
    assert page.locator('#ed-worst').is_visible()
    assert page.locator('#ed-worst').inner_text() == '영역 배관망 추출'
    assert page.locator('#ed-recalc-row').is_hidden()
    assert page.locator('#ed-edits').is_hidden()
    assert page.locator('#h-primary').get_attribute('data-native-action') == 'ed-worst'
    # Native rerenders may remove `hidden`; H's removal remains in force.
    page.evaluate("document.querySelector('#ed-recalc-row').classList.remove('hidden');__hTest.renderEdit()")
    assert page.locator('#ed-recalc-row').is_hidden()
    page.screenshot(path=str(tmp / 'single-area-extraction.png'))
    assert not [path for method, path in requests if method == 'POST']
