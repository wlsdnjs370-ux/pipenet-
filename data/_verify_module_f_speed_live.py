"""Exercise an isolated browser session against the updated local 5051 server."""
from __future__ import annotations

import json
from pathlib import Path
import time
import tempfile
import shutil
from datetime import datetime

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'cad_project_editor_g/docs/benchmarks'
SOURCE = ROOT / 'data/uploads/B1F 현장조사 소화설비 평면도_컨셉2.dxf'
NAME = f'B1F_컨셉2_속도검증_{datetime.now():%Y%m%d_%H%M%S}.dxf'
BASE = 'http://127.0.0.1:5051'
report = {'file': NAME, 'bytes': SOURCE.stat().st_size, 'base': BASE}
errors = []
sessions = []

with tempfile.TemporaryDirectory(prefix='browser_upload_', dir=OUT) as work, sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={'width': 1680, 'height': 1000})
    page.on('pageerror', lambda error: errors.append(str(error)))

    def received(response):
        if response.url.endswith('/api/module-f/slot/open') and response.ok:
            sessions.append(response.json()['sid'])

    page.on('response', received)
    page.goto(BASE + '/module-f', wait_until='networkidle')
    if '/login' in page.url:
        # The local-only hint supplies the configured password; never print it.
        password = page.locator('.pw-value').inner_text()
        page.fill('input[type=password]', password)
        page.click('button[type=submit]')
        page.wait_for_url('**/module-f', timeout=20000)
    assert '/module-f' in page.url
    upload = Path(work) / NAME
    shutil.copyfile(SOURCE, upload)
    page.set_input_files('#dxf', str(upload))
    start = time.perf_counter()
    page.click('#btn-open')
    print('Browser upload started', flush=True)
    page.wait_for_function("document.querySelector('#start-note').textContent.includes('자동 인식 결과를 찍어 뒀습니다')", timeout=300000)
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=10000)
    report['upload_to_pick_seconds'] = round(time.perf_counter() - start, 2)
    sid = sessions[-1]
    report['sid'] = sid
    adopted = page.request.get(BASE + f'/api/module-f/job?sid={sid}').json()
    report['adoption_log'] = adopted.get('lines')
    report['pick_status'] = page.inner_text('#status')
    print(json.dumps({k: report[k] for k in ('upload_to_pick_seconds', 'pick_status')}, ensure_ascii=False), flush=True)
    assert any('소요 ' in line and '[채택] 완료' in line for line in report['adoption_log']), 'Updated timing code was not loaded'
    start = time.perf_counter()
    page.click('#pk-next')
    page.wait_for_function("document.querySelector('#status').textContent.includes(' · 노드 ')", timeout=300000)
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')", timeout=10000)
    report['build_to_edit_seconds'] = round(time.perf_counter() - start, 2)
    report['edit_status'] = page.inner_text('#status')
    job = page.request.get(BASE + f'/api/module-f/job?sid={sid}').json()
    report['build_log'] = job.get('lines')
    edit = page.request.get(BASE + f'/api/module-f/edit/state?sid={sid}').json()
    report['counts'] = edit['state']['counts']
    assert report['counts']['heads'] == 6896, report['counts']
    report['page_errors'] = errors
    assert not errors, errors
    OUT.mkdir(parents=True, exist_ok=True)
    page.screenshot(path=str(OUT / 'upload_live.png'))
    (OUT / 'upload_live.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({k: report[k] for k in ('build_to_edit_seconds', 'edit_status', 'counts', 'page_errors')}, ensure_ascii=False), flush=True)
    browser.close()
