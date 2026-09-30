"""H integration presentation reuses the actual merge/sizing/export controls."""
from copy import deepcopy

import pytest

from routes.module_f import api_merge
from test_module_h_browser import h_ui, seed
from test_module_f_merge import _riser


def prepare(h_ui):
    """Provide real joinable input tables without touching a user's drawing."""
    page, sess, requests, tmp = h_ui
    sess['slots']['system']['riser'] = _riser()
    machine = _riser(2)
    machine['nodes'][0]['label'] = 'M1'
    machine['nodes'][1]['label'] = 'M2'
    machine['pipes'][0].update(label='MP1', **{'in': 'M1', 'out': 'M2'})
    machine.update(source_node_label='M1', conn_node_label='M2')
    sess['slots']['machineroom']['machineroom'] = machine
    for graph in [sess['slots']['system']['riser'], machine]:
        for pipe in graph['pipes']:
            pipe.setdefault('c', 120)
    sess['supply_mode'] = 'lsp_gravity'
    api_merge.rebuild_merged(sess)
    for pipe in sess['merged']['combined'].pipes:
        pipe.setdefault('c', 120)
    page.wait_for_function('!!window.ModuleHIntegration')
    seed(page, sess, 'edit')
    page.evaluate('ModuleHAccess.refresh()')
    page.wait_for_function('ModuleHAccess.canMerge')
    page.evaluate('ModuleHAccess.merge()')
    page.wait_for_function("__mf.stage==='merge' && !!__mf.mergeView?.pipes.length")
    page.wait_for_selector('#hm-workspace:not([hidden])')
    page.wait_for_function("document.querySelector('#h-drawer').dataset.pane==='task'")
    return page, sess, requests, tmp


def test_console_real_merge_readiness_native_handlers_and_no_automatic_writes(h_ui):
    page, sess, requests, tmp = prepare(h_ui)
    original = deepcopy(sess['merged']['combined'].pipes)
    posts = [r for r in requests if r[0] == 'POST']
    assert page.locator('#hm-state').inner_text() == '결합 완료 · 검토 필요'
    assert page.locator('#hm-sources li[data-ready="true"]').count() == 3
    assert page.locator('#hm-pipes').inner_text() == str(len(original))
    assert page.locator('#mg-build').is_visible()
    assert page.locator('#mg-summary').is_hidden()
    page.click('#hm-evidence > summary')
    assert page.locator('#mg-summary').is_visible()
    page.click('#hm-evidence > summary')
    page.locator('#h-drawer .h-drawer-scroll').evaluate('el=>el.scrollTop=0')
    page.screenshot(path=str(tmp / 'integration-connections.png'))
    page.click('#hm-output')
    assert page.locator('#h-output-basis').input_value() == 'original'
    assert page.locator('#mg-dl-coord').input_value() == '_iso'
    assert page.locator('#mg-emit').is_visible()
    assert '원본 배관망 출력' in page.locator('#hm-output-note').inner_text()
    page.click('#hm-sizing')
    assert page.locator('#sz-flow').input_value() == '80'
    assert page.locator('#sz-branch').input_value() == '6'
    assert page.locator('#ha-sizing-scope').is_visible()
    page.select_option('#sz-mode', 'fixed_supply')
    page.wait_for_selector('#sz-pressure', state='visible')
    assert page.locator('#sz-pressure').is_visible()
    assert page.locator('#sz-ceiling').is_hidden()
    assert page.locator('#sz-source').is_hidden()
    assert not page.locator('#sz-confirm').is_checked()
    page.click('#hm-view')
    assert page.locator('#mg-iso').is_checked()
    page.click('#hm-table')
    page.wait_for_selector('#h-properties:not([hidden]) #hp-rows tr')
    assert [r for r in requests if r[0] == 'POST'] == posts
    assert sess['merged']['combined'].pipes == original
    # View changes and repeated layout updates cannot duplicate engine IDs.
    assert page.evaluate("""()=>{const ids=[...document.querySelectorAll('[id]')].map(x=>x.id);return ids.length===new Set(ids).size;}""")
    page.click('#hp-close')
    page.click('#hm-connect')
    with page.expect_request('**/api/module-f/merge/build'):
        page.click('#mg-build')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')")
    assert sess['merged']['combined'].pipes == original


@pytest.mark.parametrize('size', [(1500, 960), (1024, 768), (390, 844), (844, 390)])
def test_console_responsive_and_keyboard(h_ui, size):
    page, sess, requests, tmp = prepare(h_ui)
    page.set_viewport_size({'width': size[0], 'height': size[1]})
    page.wait_for_timeout(100)
    for id in ['hm-workspace', 'h-drawer', 'hm-connect', 'hm-output']:
        box = page.locator('#' + id).bounding_box()
        assert box['x'] >= 0 and box['x'] + box['width'] <= size[0] + 1
        assert box['y'] >= 0 and box['y'] + box['height'] <= size[1] + 1
    page.locator('#hm-output').focus()
    page.keyboard.press('Enter')
    page.wait_for_function("document.querySelector('#hm-output').getAttribute('aria-pressed')==='true'")
    assert page.locator('#mg-emit').is_visible()
    assert page.locator('#h-drawer').get_attribute('data-pane') == 'export'
    page.screenshot(path=str(tmp / f'integration-output-{size[0]}.png'))
    page.keyboard.press('Escape')
    assert page.locator('#h-drawer').is_hidden()
    page.evaluate("__hTest.setStage('edit')")
    page.wait_for_function("!document.body.classList.contains('h-integration')")
    assert page.locator('#hm-workspace').is_hidden()


def test_stale_busy_split_and_proposal_status_are_never_success(h_ui):
    page, sess, requests, tmp = prepare(h_ui)
    page.evaluate("""()=>{const e=document.querySelector('#mg-stale');e.textContent='기준망이 바뀌었습니다';e.classList.remove('hidden');}""")
    page.wait_for_function("document.querySelector('#hm-state').textContent==='변경사항 반영 필요'")
    page.click('#hm-output')
    assert '연결을 갱신' in page.locator('#hm-output-note').inner_text()
    page.evaluate("__hTest.busy(true, '통합 검증')")
    page.wait_for_function("document.querySelector('#hm-connect').disabled")
    assert page.locator('#hm-output').is_disabled()
    assert page.locator('#hm-state').inner_text() == '처리 중'
    page.evaluate("""()=>{__hTest.busy(false);document.querySelector('#mg-stale').classList.add('hidden');const e=document.querySelector('#mg-split');e.textContent='배관망이 나뉘었습니다';e.classList.remove('hidden');}""")
    page.wait_for_function("document.querySelector('#hm-state').textContent==='연결 확인 필요'")
    assert page.locator('#mg-split').is_visible()
    page.evaluate("""()=>{const s=document.querySelector('#h-output-basis');s.options[1].disabled=false;s.value='proposal';s.dispatchEvent(new Event('change',{bubbles:true}));}""")
    page.wait_for_function("document.querySelector('#hm-output-note').textContent.includes('다시 계산')")
    assert '다시 계산' in page.locator('#hm-output-note').inner_text()
    assert not [r for r in requests if r[0] == 'POST']


def test_multiple_system_cards_and_absent_counts_are_not_invented(h_ui):
    page, sess, requests, tmp = prepare(h_ui)
    page.evaluate("""()=>{__mf.merge={order:['plan','system','system2','machineroom'],labels:{plan:'평면도',system:'계통도 1',system2:'계통도 2',machineroom:'기계실'},ready:{plan:true,system:true,system2:false,machineroom:true}};ModuleHUI.queue();}""")
    page.wait_for_function("document.querySelectorAll('#hm-sources li').length===4")
    assert page.locator('#hm-sources li').nth(2).inner_text().endswith('미준비')
    assert page.locator('#hm-pipes').inner_text() == '—'
    assert page.locator('#hm-state').inner_text() == '결합 전'
