"""Provenance cards and live edits through isolated real browser/server routes."""
import os
from pathlib import Path
from urllib.parse import urlparse

import pytest
from flask import Flask

from test_module_f_viewport_browser import install_page, settle, pw_api
from test_module_f_network_editor import design
from routes.module_f import api_design, api_merge, api_network_edit, jobs, network_edit as ne


@pytest.fixture()
def bore_ui(tmp_path, monkeypatch):
    monkeypatch.setattr(ne, 'HISTORY_DIR', tmp_path / 'history')
    monkeypatch.syspath_prepend(str(Path(ne.__file__).resolve().parents[2] / 'core'))
    sess = jobs._new_session(key='browser-bore')
    sess['design'] = design()
    a, b = sess['design']['tables'].pipes
    a.update(dia_src='nfpc_fallback', bore_provenance=dict(version=1, source='nfpc_fallback',
        auto_mm=25, rule_mm=25, head_count=1, reason='no_match'))
    b.update(dia_src='text', bore_provenance=dict(version=1, source='text',
        auto_mm=25, text_mm=25, text_xy_mm=[93000, 37000], distance_mm=154.09))
    app = Flask(__name__); app.testing = True
    api_design.register(app, UPLOAD_DIR=tmp_path)
    api_merge.register(app, UPLOAD_DIR=tmp_path)
    api_network_edit.register(app)
    client = app.test_client()
    source = (Path(ne.__file__).resolve().parents[2] / 'static/module_f.js').read_text(encoding='utf-8')
    source = source.replace('  setStage("open");\n  loadSaved();',
        '  window.__boreUi={select:insSelect,preview:designPreview,merged:loadMergeView}; window.__styleTest={drawMerged};\n  setStage("open");\n  loadSaved();')
    try:
        with pw_api.sync_playwright() as pw:
            browser = pw.chromium.launch()
            page = browser.new_page(viewport={'width':1500, 'height':1000})
            page.set_default_timeout(7000)
            errors = install_page(page, source)
            def route(req):
                url = urlparse(req.request.url)
                response = client.open(url.path + ('?' + url.query if url.query else ''),
                    method=req.request.method, data=req.request.post_data, content_type='application/json')
                if response.status_code == 404:
                    req.fulfill(json={'ok':True, 'items':[], 'rows':[], 'fields':{}})
                else:
                    req.fulfill(status=response.status_code, body=response.data, content_type=response.content_type)
            page.route('http://module-f.test/api/module-f/**', route)
            page.evaluate('''async sid=>{
                Object.assign(__mf,{sid,slot:'plan',method:'manual'});
                __viewTest.setStage('design');
                document.querySelector('#dg-iso').checked=true;
                await __boreUi.preview(); __boreUi.select('pipe','P2');
            }''', sess['id'])
            page.wait_for_function("document.querySelector('#ne-material').options.length>0")
            settle(page)
            yield page, sess
            assert not errors, errors
            browser.close()
    finally:
        jobs._SESSIONS.pop(sess['id'], None)


def test_source_legend_card_and_live_edit_undo_redo(bore_ui):
    page, sess = bore_ui
    assert page.locator('#bore-color').is_checked()
    legend = page.locator('#bore-legend').inner_text()
    assert '도면 표기 참조 1' in legend and '규약 보완 1' in legend
    card = page.locator('.bore-card')
    assert card.get_attribute('data-bore-category') == 'drawing'
    assert '최근접 문자 매칭' in card.inner_text()
    assert '93,000' in card.inner_text()
    before = page.evaluate('JSON.stringify(__mf.design.view.pipes)')
    page.uncheck('#bore-color'); settle(page)
    page.check('#bore-color'); settle(page)
    assert page.evaluate('JSON.stringify(__mf.design.view.pipes)') == before
    page.click('.bore-edit')
    page.wait_for_function("document.querySelector('#ne-op').value==='pipe'")
    page.select_option('#ne-dn','32')
    page.fill('#ne-note','원본 범위 표기 확인')
    page.click('#ne-apply')
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===32 && document.querySelector('#busy').classList.contains('hidden')")
    assert '사용자 수정' in card.inner_text() and '25 → 32' in card.inner_text()
    assert '원본 범위 표기 확인' in card.inner_text()
    assert card.get_attribute('data-bore-category') == 'drawing'
    assert sess['design']['tables'].pipes[1]['bore_provenance']['text_mm'] == 25
    page.click('#ne-undo')
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===25 && document.querySelector('#busy').classList.contains('hidden')")
    page.evaluate("__boreUi.select('pipe','P2')")
    assert '사용자 수정' not in card.inner_text()
    page.click('#ne-redo')
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===32 && document.querySelector('#busy').classList.contains('hidden')")
    page.evaluate("__boreUi.select('pipe','P2')")
    assert '사용자 수정' in card.inner_text()
    if os.environ.get('BORE_UI_SCREENSHOT'):
        page.screenshot(path=os.environ['BORE_UI_SCREENSHOT'], full_page=True)


def test_merged_source_and_quick_editor_preserve_drawing_evidence(bore_ui):
    page, sess = bore_ui
    from test_module_f_merge import _riser
    sess['slots']['system']['riser'] = _riser()
    sess['supply_mode'] = 'lsp_gravity'
    api_merge.rebuild_merged(sess)
    page.evaluate('''async()=>{
        __viewTest.setStage('merge'); document.querySelector('#mg-iso').checked=true;
        await __boreUi.merged(); __boreUi.select('mgpipe','P2');
    }''')
    card = page.locator('.bore-card')
    assert card.get_attribute('data-bore-category') == 'drawing'
    assert '근거 미확인' in page.locator('#bore-legend').inner_text()
    from test_module_f_design_style_browser import capture, strokes
    lines=[s for s in strokes(capture(page,'drawMerged')) if len(s['path'])==2
           and [p[0] for p in s['path']]==['moveTo','lineTo']]
    expected=page.evaluate('__mf.mergeView.pipes.map(p=>ModuleFBore.style(p).color)')
    assert sorted(s['color'] for s in lines)==sorted(expected)
    page.click('.bore-edit')
    page.wait_for_function("document.querySelector('#ne-op').value==='pipe'")
    page.select_option('#ne-dn','32')
    page.fill('#ne-note','통합에서 도면 확인')
    page.click('#ne-apply')
    page.wait_for_function("__mf.mergeView.pipes.find(p=>p.label==='P2').dia===32 && document.querySelector('#busy').classList.contains('hidden')")
    assert card.get_attribute('data-bore-category') == 'drawing'
    assert '통합에서 도면 확인' in card.inner_text()
    assert next(p for p in sess['merged']['combined'].pipes if p['label']=='P2')['bore_provenance']['text_mm'] == 25


def test_legacy_unknown_and_review_are_not_falsely_drawing_or_fallback(bore_ui):
    page, _ = bore_ui
    result = page.evaluate('''()=>{
        const unknown={dia:65};
        const review={dia:40,bore_info:{source:'nfpc_min',category:'review',text_mm:32,rule_mm:40,auto_mm:40}};
        return {unknown:ModuleFBore.card(unknown),review:ModuleFBore.card(review),
            styles:[ModuleFBore.style(unknown),ModuleFBore.style(review)]};
    }''')
    assert '근거 미확인' in result['unknown']
    assert '규약 보완' not in result['unknown']
    assert '상향되었습니다' in result['review']
    assert result['styles'][0]['color'] == '#8795a5'
    assert result['styles'][1]['color'] == '#e7a1ff'


def test_topology_evidence_card_survives_api_and_user_resolution(bore_ui):
    page,sess = bore_ui
    row = sess['design']['tables'].pipes[1]
    row['bore_provenance'].update(version=2,method='동일 관로 전파 · 추론',
        text_rotation_deg=0,full_head_count=13,block_export=True,candidate_mm=[25,65],
        review_reasons=['관경 경계를 확인하세요.'],
        excluded_nearest=dict(text_mm=25,reason='다른 관로의 표기로 분리'))
    page.evaluate("async()=>{await __boreUi.preview();__boreUi.select('pipe','P2');}")
    card=page.locator('.bore-card')
    assert card.get_attribute('data-bore-category')=='review'
    for text in ('동일 관로 전파','다른 관로의 표기로 분리','13개','25 / 65','확인 전 파일 산출 보류'):
        assert text in card.inner_text()
    page.click('.bore-edit')
    page.wait_for_function("document.querySelector('#ne-op').value==='pipe'")
    page.select_option('#ne-dn','65')
    page.fill('#ne-note','원문과 구간 확인')
    page.click('#ne-apply')
    page.wait_for_function("__mf.design.view.pipes.find(p=>p.label==='P2').dia===65 && document.querySelector('#busy').classList.contains('hidden')")
    assert not sess['design']['tables'].pipes[1]['bore_provenance']['block_export']
    assert '확인 전 파일 산출 보류' not in card.inner_text()
    assert '원문과 구간 확인' in card.inner_text()
