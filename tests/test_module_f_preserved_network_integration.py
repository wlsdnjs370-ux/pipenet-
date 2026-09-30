"""HTTP, persistence, isolated-draft and browser checks for preserved networks."""
from __future__ import annotations

import copy
from urllib.parse import urlparse

import pytest

from test_module_f_flow_integration import flow_client
from routes.module_f import api_convert, api_merge
from routes.module_f.cancellation import install
from services.cad_import.design.flow import network_for_board
from services.cad_import.design.restrict import restrict_to_worst
from services.cad_import.edit.io import load_edits, write_edits


def switch(c, sess, mode):
    result = c.post('/api/module-f/edit/network-mode',
                    json={'sid': sess['id'], 'network_mode': mode})
    assert result.status_code == 200, result.json
    return result.json


@pytest.mark.parametrize('mode', ['loop', 'grid'])
def test_mode_flow_selection_download_preserve_cycles_and_tree_unchanged(flow_client, mode):
    c, sess = flow_client
    install(c.application)
    b = sess['edit'].board
    original = copy.deepcopy((b.pts, b.edges))
    tree_report = c.post('/api/module-f/edit/flow', json={'sid': sess['id']}).json['state']['flow_report']
    switch(c, sess, mode)
    assert sess['flow_report'] is None and sess['worst'] is None
    assert not sess['edit']._flowed
    assert sess['edit'].board is not b
    result = c.post('/api/module-f/edit/flow', json={'sid': sess['id']})
    assert result.status_code == 200, result.json
    report = result.json['state']['flow_report']
    assert report['network_mode'] == mode and report['cycle_rank'] == 1
    assert report['max_load_all'] is None
    picked = c.post('/api/module-f/edit/worst', json={'sid': sess['id'], 'k': 1})
    assert picked.status_code == 200, picked.json
    w = sess['worst']
    assert len(w['heads']) == 1 and len(w['edges']) > len(w['worst_path']) - 1
    assert not w['loads'] and w['max_load'] is None
    assert picked.json['summary']['rank_invariant']['ok']
    assert len(picked.json['state']['worst']['corridor']) == len(w['edges'])
    audit = c.get('/api/module-f/edit/flow/report', query_string={'sid': sess['id']})
    assert audit.status_code == 200, audit.json
    assert 'module-f-network.json' in audit.headers['Content-Disposition']
    assert audit.json['network']['cycle_rank'] == 1
    assert audit.json['scenario']['cycle_rank'] == 1
    assert sum(h['selected'] for h in audit.json['scenario']['nozzles']) == 1
    assert not audit.json['network']['calculation_ready']
    assert (sess['edit'].board.pts, sess['edit'].board.edges) == original
    switch(c, sess, 'tree')
    after = c.post('/api/module-f/edit/flow', json={'sid': sess['id']}).json['state']['flow_report']
    assert after == tree_report
    c.post('/api/module-f/edit/worst', json={'sid': sess['id'], 'k': 1})
    c.post('/api/module-f/design/build', json={'sid': sess['id'], 'k': 1})
    assert sess['job']['result']['ok'], sess['job']['result']


def test_invalid_mode_and_stale_report_fail_without_source_mutation(flow_client):
    c, sess = flow_client
    install(c.application)
    switch(c, sess, 'grid')
    previous = sess['edit']
    bad = c.post('/api/module-f/edit/network-mode', json={'sid': sess['id'], 'network_mode': 'unknown'})
    assert bad.status_code == 400
    assert sess['edit'] is previous and previous.board.network_mode == 'grid'
    c.post('/api/module-f/edit/flow', json={'sid': sess['id']})
    sess['edit'].board.pts[1] = (1100, 0)
    assert c.get('/api/module-f/edit/flow/report', query_string={'sid': sess['id']}).status_code == 400


def test_legacy_calculation_exports_are_guarded_including_inactive_plan(flow_client, tmp_path):
    c, sess = flow_client
    api_convert.register(c.application, UPLOAD_DIR=tmp_path)
    api_merge.register(c.application, UPLOAD_DIR=tmp_path)
    switch(c, sess, 'loop')
    preview = c.get('/api/module-f/design/preview', query_string={'sid':sess['id']})
    assert preview.status_code == 200 and preview.json['view'] is None
    assert '루프·그리드' in preview.json['message']
    c.post('/api/module-f/design/build',json={'sid':sess['id']})
    assert not sess['job']['result']['ok'] and '기본 관경' in sess['job']['result']['error']
    for path in ('design/emit', 'convert/run', 'merge/build', 'merge/emit'):
        result = c.post('/api/module-f/' + path, json={'sid': sess['id']})
        assert result.status_code == 409, (path, result.json)
    for path in ('download', 'merge/download'):
        result = c.get('/api/module-f/' + path, query_string={'sid': sess['id']})
        assert result.status_code == 409, (path, result.json)
    from services.cad_import.convert.engine import ensure_planar
    with pytest.raises(ValueError, match='루프·그리드'):
        restrict_to_worst(sess['edit'].convert_payload(), sess['edit'].board, {'heads': [0]})
    with pytest.raises(ValueError, match='루프·그리드'):
        ensure_planar(sess['edit'].convert_payload())
    sess['slots'] = {'plan': {'edit': sess.pop('edit')}}
    for path in ('merge/build', 'merge/emit'):
        assert c.post('/api/module-f/' + path, json={'sid': sess['id']}).status_code == 409


def test_save_load_explicit_mode_and_old_drawing_defaults(flow_client, tmp_path):
    _, sess = flow_client
    b = sess['edit'].board
    restored = copy.deepcopy(b)
    b.network_mode = 'grid'
    write_edits(b, str(tmp_path))
    assert load_edits(restored, str(tmp_path))
    assert restored.network_mode == 'grid'
    # Old version-1 records have no mode; exercising loader, not migrations.
    import json
    path = tmp_path / 'legacy'
    path.mkdir()
    filename = write_edits(b, str(path))
    from pathlib import Path
    raw = json.loads(Path(filename).read_text(encoding='utf-8'))
    raw.pop('network_mode')
    Path(filename).write_text(json.dumps(raw), encoding='utf-8')
    assert load_edits(restored, str(path)) and restored.network_mode == 'tree'


@pytest.mark.parametrize('head_tee', [False, True])
def test_browser_selects_preserved_network_without_fake_loads(flow_client, tmp_path, head_tee):
    from test_module_f_viewport_browser import install_page, pw_api, ROOT
    c, sess = flow_client
    if head_tee:
        from test_module_f_head_junctions import junction_board
        from services.cad_import.edit.session import EditSession
        b = junction_board()
        sess['edit'] = EditSession(b, key=b.key)
    source = (ROOT / 'static/module_f.js').read_text(encoding='utf-8')
    source = source.replace('  setStage("open");\n  loadSaved();',
        '  window.__networkTest={setEdit,renderEdit};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(viewport={'width':1440, 'height':1000})
        errors = install_page(page, source)
        def route(req):
            url = urlparse(req.request.url)
            r = c.open(url.path + ('?' + url.query if url.query else ''), method=req.request.method,
                       data=req.request.post_data, content_type='application/json')
            if r.status_code == 404:
                req.fulfill(json={'ok':True, 'rows':[], 'items':[], 'fields':{}})
            else:
                req.fulfill(status=r.status_code, body=r.data, content_type=r.content_type)
        page.route('http://module-f.test/api/module-f/**', route)
        state = c.get('/api/module-f/edit/state', query_string={'sid':sess['id']}).json['state']
        page.evaluate('''({sid,state})=>{
          Object.assign(__mf,{sid,slot:'plan',method:'manual'});
          __networkTest.setEdit(state);__viewTest.setStage('edit');__networkTest.renderEdit();
        }''', {'sid':sess['id'], 'state':state})
        page.locator('#mf-advanced-edit > summary').click()
        page.select_option('#ed-network-mode', 'grid')
        page.wait_for_function("__mf.edit.network_mode==='grid'")
        page.get_by_text('루프·그리드 · 검토 기준',exact=True).click()
        assert page.locator('#ed-network-note').is_visible()
        assert page.locator('#ed-worst').inner_text() == '작동 헤드 후보 선정'
        assert '그리드 · 양단 연결 보존' in page.locator('#ed-info').text_content()
        if head_tee:
            assert page.evaluate('__mf.edit.cycle_recovery.split_edges') == 1
            assert '이음 복원 1곳' in page.locator('#ed-flow-summary').inner_text()
        page.click('#ed-flow')
        page.wait_for_function('!!__mf.edit.flow_report')
        assert '연결 보존' in page.locator('#ed-flow-summary').inner_text()
        assert 'null개' not in page.locator('#ed-info').text_content()
        assert '미산정' in page.locator('#ed-info').text_content()
        if head_tee:
            assert page.evaluate('__mf.edit.flow_report.heads') == 1
        # Real route selection -> real state -> actual rendering, all in fixtures.
        chosen = c.post('/api/module-f/edit/worst', json={'sid':sess['id'], 'k':1}).json
        assert chosen['ok'], chosen
        page.evaluate('''s=>{__networkTest.setEdit(s);__networkTest.renderEdit();__viewTest.draw();}''', chosen['state'])
        assert '부하 미산정' in page.locator('#ed-info').text_content()
        page.screenshot(path=str(tmp_path / 'preserved-network.png'))
        assert not errors, errors
        browser.close()


@pytest.mark.parametrize('key', ['B1F 현장조사 소화설비 평면도', '1. 입력도면 대명동 단위세대 평면도'])
def test_read_only_saved_network(key):
    """Opt-in local evidence; never rebuild a missing cache or write edits."""
    import os
    if os.environ.get('MODULE_F_REAL_NETWORK_CHECK') != '1':
        pytest.skip('Local saved drawing check is opt-in')
    from services.cad_import.edit.io import _disp_cache_load, _board_from_data
    data = _disp_cache_load(key)
    if data is None:
        pytest.skip('No current cache; no user cache rebuild in this check')
    board = _board_from_data(key, data)
    load_edits(board)
    if len(board.sources) != 1:
        pytest.skip('Saved drawing needs an explicit source')
    before = copy.deepcopy((board.pts, board.edges, getattr(board, 'edge_len_mm', {})))
    board.network_mode = 'grid'
    network = network_for_board(board)
    assert network.reference.edges <= network.edges
    assert network.edges <= set(network.reference.lengths_mm)
    chosen = sorted(set(network.reference.representatives.values()))[:30]
    out = network.extraction(board.pts, selected_heads=chosen)
    assert sum(h['selected'] for h in out['nozzles']) == len(chosen)
    assert all(p['bore_m'] is None for p in out['pipes'])
    assert (board.pts, board.edges, getattr(board, 'edge_len_mm', {})) == before
    print('READ ONLY NETWORK', key, network.report())
