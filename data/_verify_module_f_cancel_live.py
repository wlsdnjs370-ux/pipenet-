"""Verify upload, selected-head confirmation and optional diagnosis on local 5051."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import shutil
import tempfile
import time
from datetime import datetime

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'cad_project_editor_g/docs/benchmarks'
SOURCE = ROOT / 'data/uploads/B1F 현장조사 소화설비 평면도_컨셉2.dxf'
NAME = f'B1F_컨셉2_중지검증_{datetime.now():%Y%m%d_%H%M%S}.dxf'
BASE = os.environ.get('MODULE_F_TEST_URL', 'http://127.0.0.1:5051')
report = {'file': NAME, 'bytes': SOURCE.stat().st_size, 'base': BASE}
errors = []
sessions = []

with tempfile.TemporaryDirectory(prefix='design_browser_', dir=OUT) as work, sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={'width': 1680, 'height': 1000})
    page.on('pageerror', lambda error: errors.append(str(error)))

    def received(response):
        if response.url.endswith('/api/module-f/slot/open') and response.ok:
            sessions.append(response.json()['sid'])

    page.on('response', received)
    page.goto(BASE + '/module-f', wait_until='networkidle')
    if '/login' in page.url:
        # Read the local login hint without disclosing it in logs or artifacts.
        page.fill('input[type=password]', page.locator('.pw-value').inner_text())
        page.click('button[type=submit]')
        page.wait_for_url('**/module-f', timeout=20000)
    upload = Path(work) / NAME
    shutil.copyfile(SOURCE, upload)
    page.set_input_files('#dxf', str(upload))
    start = time.perf_counter()
    page.click('#btn-open')
    page.wait_for_selector('#busy-cancel', state='visible')
    page.click('#busy-cancel')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=15000)
    report['compression_cancel_seconds'] = round(time.perf_counter() - start, 2)
    assert '중지' in page.inner_text('#status')
    print('Compression/upload cancelled', flush=True)
    start = time.perf_counter()
    page.click('#btn-open')
    print('Browser upload started: ' + NAME, flush=True)
    page.wait_for_function("document.querySelector('#start-note').textContent.includes('자동 인식 결과를 찍어 뒀습니다')", timeout=300000)
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=10000)
    report['upload_to_pick_seconds'] = round(time.perf_counter() - start, 2)
    sid = sessions[-1]
    report['sid'] = sid
    (OUT / 'cancel_live_session.json').write_text(json.dumps(report, ensure_ascii=False), encoding='utf-8')
    start = time.perf_counter()
    page.click('#pk-next')
    page.wait_for_function("document.querySelector('#status').textContent.includes(' · 노드 ')", timeout=300000)
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=10000)
    report['build_to_edit_seconds'] = round(time.perf_counter() - start, 2)

    def post(endpoint, **body):
        response = page.request.post(BASE + '/api/module-f/' + endpoint,
                                     data={'sid': sid, **body}, timeout=300000)
        result = response.json()
        assert response.ok and result.get('ok'), result
        return result

    def get(endpoint):
        response = page.request.get(BASE + '/api/module-f/' + endpoint + '?sid=' + sid, timeout=300000)
        assert response.ok, response.status
        return response.json()

    edit = get('edit/state')
    report['counts'] = edit['state']['counts']
    assert report['counts']['heads'] > 6000, 'Expected the large drawing fixture'
    post('edit/mode', mode='알람밸브위치')
    report['alarm'] = post('edit/click', x=662692, y=175774, max_d=3000)['report']
    report['selection'] = post('edit/worst', k=10, zones=[[655000,165000,675000,180000]])['summary']
    # Use the actual UI for confirmation and for the new diagnosis control.
    page.locator('#ed-k').fill('10')
    page.click('#ed-next')
    page.wait_for_selector('#dg-build', state='visible')
    with page.expect_response('**/api/module-f/design/build'):
        page.click('#dg-build')
    page.wait_for_timeout(400)
    page.screenshot(path=str(OUT / 'cancel_live_working.png'))
    stop_start = time.perf_counter()
    page.click('#busy-cancel')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=15000)
    report['build_cancel_seconds'] = round(time.perf_counter() - stop_start, 2)
    assert get('job')['state'] == 'cancelled', get('job')
    assert not get('design/preview').get('tables')
    assert get('edit/state')['state']['counts'] == report['counts']
    print('Design build cancelled; retrying', flush=True)
    start = time.perf_counter()
    page.click('#dg-build')
    page.wait_for_function("document.querySelector('#dg-diagnose-note').textContent.includes('선택한 헤드의 연결 검사는 완료')", timeout=120000)
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=10000)
    report['confirm_to_preview_seconds'] = round(time.perf_counter() - start, 2)
    report['confirm_job'] = get('job')
    report['summary'] = get('convert/result')['result']['summary']
    assert report['summary']['handoff']['in_table'] == 10
    assert report['summary']['diagnostics_state'] == 'pending'
    assert report['summary']['excluded_heads'] is None
    assert page.locator('#dg-mk-unatt').is_disabled()
    before = get('design/preview')['tables']
    before = {k: before[k] for k in ('nodes','pipes','nozzles','fittings','equipment')}
    print(json.dumps({k:report[k] for k in ('upload_to_pick_seconds','build_to_edit_seconds','confirm_to_preview_seconds','summary')},ensure_ascii=False),flush=True)
    page.locator('#dg-plan').uncheck()
    page.screenshot(path=str(OUT / 'cancel_live_pending.png'))

    with page.expect_response('**/api/module-f/design/diagnose'):
        page.click('#dg-diagnose')
    page.wait_for_timeout(700)
    stop_start = time.perf_counter()
    page.click('#busy-cancel')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=15000)
    report['diagnosis_cancel_seconds'] = round(time.perf_counter() - stop_start, 2)
    assert get('job')['state'] == 'cancelled', get('job')
    cancelled_preview = get('design/preview')
    assert {k:cancelled_preview['tables'][k] for k in before} == before
    assert cancelled_preview['diagnostics']['state'] == 'pending'
    page.screenshot(path=str(OUT / 'cancel_live_stopped.png'))
    print('Diagnosis cancelled; tables unchanged; retrying', flush=True)
    start = time.perf_counter()
    page.click('#dg-diagnose')
    page.wait_for_function("document.querySelector('#dg-diagnose-note').textContent.includes('전체 도면 진단 완료')", timeout=180000)
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=10000)
    report['diagnose_to_preview_seconds'] = round(time.perf_counter() - start, 2)
    report['diagnostic_job'] = get('job')
    after = get('design/preview')
    assert after['diagnostics']['state'] == 'done'
    assert not page.locator('#dg-mk-unatt').is_disabled()
    assert {k:after['tables'][k] for k in before} == before
    report['identical_hydraulic_rows'] = True
    report['fingerprint'] = hashlib.sha256(json.dumps(before,sort_keys=True,ensure_ascii=False).encode()).hexdigest()
    report['page_errors'] = errors
    assert not errors, errors
    page.screenshot(path=str(OUT / 'cancel_live_done.png'))
    (OUT / 'cancel_live.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({k:report[k] for k in ('compression_cancel_seconds','build_cancel_seconds','diagnosis_cancel_seconds','diagnose_to_preview_seconds','identical_hydraulic_rows','page_errors')},ensure_ascii=False),flush=True)
    browser.close()
