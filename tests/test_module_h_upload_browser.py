"""One file selection starts the shared upload pipeline in H, never in F."""
from copy import deepcopy
from urllib.parse import urlparse

import pytest

from test_module_h_browser import h_ui
from test_module_f_viewport_browser import WORLD, ui


FILE = {'name': 'selected.dxf', 'mimeType': 'application/dxf', 'buffer': b'0\nSECTION\n2\nENTITIES\n0\nENDSEC\n0\nEOF\n'}


def upload_responses(page, sid: str, kind: str, outcome: dict):
    """Stub HTTP boundaries, retaining actual file/XHR/job/stage UI handlers."""
    seen = []
    world = deepcopy(WORLD)
    world['bundles'][0].update(n_seg=2, n_circle=1, n_arc_all=1, cat='PIPE')
    world.update(counts={'segs': 2, 'circles': 1, 'arcs': 1}, dropped={'segs': 0, 'circles': 0, 'arcs': 0})
    pick = dict(materials=[], heads=[], counts={}, mat_done=False, mode='재료', armed=False,
                highlight=dict(pipe_segs=[], tri_segs=[], head_circles=[]))

    def respond(route):
        path = urlparse(route.request.url).path
        if path.endswith('/slot/open'):
            seen.append(route.request.post_data_buffer)
            if outcome.get('fail'):
                route.fulfill(status=400, json={'ok': False, 'error': 'fixture DXF 읽기 실패'})
            else:
                route.fulfill(json={'ok': True, 'sid': sid, 'filename': FILE['name'], 'needs_method': kind == 'plan'})
        elif path.endswith('/world'):
            route.fulfill(json={'ok': True, 'key': 'upload-fixture', 'world': world, 'state': pick})
        elif path.endswith('/recon'):
            route.fulfill(json={'ok': True, 'recon': {'state': 'none'}})
        elif path.endswith('/slot/read'):
            route.fulfill(json={'ok': True, 'method': 'manual'})
        elif path.endswith('/slot/state'):
            route.fulfill(json={'ok': True, 'slots': [{'kind': kind, 'active': True, 'key': 'upload-fixture', 'label': kind}], 'active': kind})
        elif path.endswith('/sub/graph'):
            route.fulfill(json={'ok': True, 'nodes': [], 'edges': [], 'layers': [], 'chosen': [], 'forced': 0, 'components': 0})
        else:
            route.continue_()

    page.route('**/api/module-f/**', respond)
    page.evaluate('kind=>{__mf.slot=kind;__hTest.setStage("open")}', kind)
    return seen


@pytest.mark.parametrize('kind,next_stage', [('plan', 'pick'), ('system', 'sub'), ('machineroom', 'sub')])
def test_file_selection_uploads_once_and_advances_without_load_button(h_ui, kind, next_stage):
    page, sess, _, _ = h_ui
    seen = upload_responses(page, sess['id'], kind, {})
    sess['_h_job_fixture'] = dict(ok=True, state='run', lines=[], elapsed=0)
    page.click('#h-welcome-open')
    page.click(f'#h-access-{kind}')
    page.wait_for_function('!document.querySelector("#h-drawer").hidden')
    assert page.locator('#btn-open').is_hidden()
    assert page.locator('#btn-open').locator('..').is_hidden()
    page.set_input_files('#dxf', FILE)
    page.wait_for_function('!!__mf.sid && !document.querySelector("#busy").classList.contains("hidden")')
    page.wait_for_function('document.querySelector("#dxf").disabled')
    assert page.locator('#h-drawer').is_hidden()
    assert page.evaluate('__mf.stage') == 'open'  # No pretend success before parsing.
    assert page.locator('#h-primary').inner_text() != '불러오기'
    page.evaluate('document.querySelector("#dxf").dispatchEvent(new Event("change",{bubbles:true}))')
    sess['_h_job_fixture'] = dict(ok=True, state='done', lines=[], elapsed=0)
    page.wait_for_function('stage=>(__mf.stage===stage || document.querySelector("#status").classList.contains("err")) && document.querySelector("#busy").classList.contains("hidden")', arg=next_stage)
    assert page.evaluate('__mf.stage') == next_stage, page.locator('#status').inner_text()
    page.wait_for_function('!document.querySelector("#dxf").disabled')
    assert len(seen) == 1
    assert b'filename="selected.dxf"' in seen[0] and FILE['buffer'] in seen[0]
    assert f'\r\n{kind}\r\n'.encode() in seen[0]
    assert page.locator('#dxf').input_value() == ''
    assert page.locator('#h-primary').inner_text() != '불러오기'


def test_cancel_no_upload_failure_no_advance_same_file_can_retry(h_ui):
    page, sess, _, _ = h_ui
    outcome = {'fail': True}
    seen = upload_responses(page, sess['id'], 'plan', outcome)
    page.click('#h-welcome-open')
    page.click('#h-access-plan')
    page.wait_for_function('!document.querySelector("#h-drawer").hidden')
    page.set_input_files('#dxf', [])
    assert not seen
    page.set_input_files('#dxf', FILE)
    page.wait_for_function('document.querySelector("#status").classList.contains("err")')
    assert page.evaluate('__mf.stage') == 'open'
    assert len(seen) == 1
    assert page.locator('#dxf').input_value() == ''
    outcome['fail'] = False
    page.click('#h-primary')
    page.set_input_files('#dxf', FILE)
    page.wait_for_function('__mf.stage==="pick"')
    assert len(seen) == 2


def test_native_f_retains_explicit_upload_confirmation(ui):
    posts = []
    ui.on('request', lambda r: posts.append(r.url) if r.method == 'POST' else None)
    ui.set_input_files('#dxf', FILE)
    assert ui.locator('#btn-open').is_visible()
    assert ui.locator('#dxf').input_value().endswith(FILE['name'])
    assert ui.evaluate('__mf.stage') == 'open'
    assert not posts


def test_upload_card_default_is_compact_resizable_and_separate_from_other_panes(h_ui):
    page, _, _, _ = h_ui
    page.click('#h-welcome-open');page.click('#h-access-plan')
    panel=page.locator('#h-drawer');page.wait_for_function('!document.querySelector("#h-drawer").hidden')
    box=panel.bounding_box()
    assert box['width'] == pytest.approx(360, abs=2)
    assert box['height'] == pytest.approx(205, abs=2)
    field=page.locator('#dxf').bounding_box()
    assert field['y'] + field['height'] < box['y'] + box['height']
    handle=panel.locator('[data-edge=se]').bounding_box()
    x,y=handle['x']+5,handle['y']+5
    page.mouse.move(x,y);page.mouse.down();page.mouse.move(x+100,y+90,steps=5);page.mouse.up()
    bigger=panel.bounding_box()
    assert bigger['width'] > box['width']+80 and bigger['height'] > box['height']+70
    page.click('#h-drawer-close');page.click('#h-history')
    assert panel.bounding_box()['height'] > 400
    page.click('#h-drawer-close');page.click('#h-primary')
    assert panel.bounding_box()['height'] == pytest.approx(bigger['height'], abs=2)
    page.locator('.h-drawer-head h2').dblclick()
    assert panel.bounding_box()['height'] == pytest.approx(205, abs=2)
    page.set_viewport_size(dict(width=390,height=640))
    box=panel.bounding_box()
    assert box['x'] >= 0 and box['x']+box['width'] <= 390
