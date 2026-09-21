"""Exercise rectangle zoom and stage transitions on local 5051 with a test copy."""
from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import shutil
import tempfile

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'cad_project_editor_g/docs/benchmarks'
NAME = f'B1F_줌검증_{datetime.now():%Y%m%d_%H%M%S}.dxf'
BASE = 'http://127.0.0.1:5051'
report = {'base':BASE,'file':NAME}
errors = []

with tempfile.TemporaryDirectory(prefix='zoom_verify_',dir=OUT) as tmp, sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={'width':1680,'height':1000})
    page.on('pageerror',lambda error: errors.append(str(error)))
    page.goto(BASE+'/module-f')
    if '/login' in page.url:
        page.fill('input[type=password]',page.locator('.pw-value').inner_text())
        page.click('button[type=submit]')
        page.wait_for_url('**/module-f')
    assert page.locator('#open-zoom-window').count()==1
    target = Path(tmp)/NAME
    shutil.copyfile(ROOT/'data/uploads/B1F 현장조사 소화설비 평면도_컨셉2.dxf',target)
    page.set_input_files('#dxf',str(target))
    page.click('#btn-open')
    print('Upload: '+NAME,flush=True)
    page.wait_for_function("__mf.stage==='pick' && document.querySelector('#busy').classList.contains('hidden')",timeout=180000)
    page.locator('#steps div').filter(has_text='도면 열기').click()
    page.wait_for_function("__mf.stage==='open'")
    before = page.evaluate('({...__mf.view})')
    regions_before = page.evaluate('JSON.stringify(__mf.zones)')
    # The work area is around the known B1F alarm valve. Use a screen rectangle
    # around it, at least 30px each way even when the full sheet is very large.
    xy = page.evaluate('({x:__mf.toScreenX(662692),y:__mf.toScreenY(175774)})')
    box = page.locator('#cv').bounding_box()
    cx = max(40,min(box['width']-40,xy['x']))
    cy = max(40,min(box['height']-80,xy['y']))
    page.click('#open-zoom-window')
    page.mouse.move(box['x']+cx-30,box['y']+cy-30)
    page.mouse.down()
    page.mouse.move(box['x']+cx+30,box['y']+cy+30,steps=5)
    page.screenshot(path=str(OUT/'viewport_live_rectangle.png'))
    page.mouse.up()
    page.wait_for_function('s => __mf.view.scale>s',arg=before['scale'])
    chosen = page.evaluate('({...__mf.view})')
    report['camera'] = chosen
    assert page.evaluate('JSON.stringify(__mf.zones)')==regions_before
    page.locator('#steps div').filter(has_text='찍기').click()
    page.wait_for_function("__mf.stage==='pick'")
    assert page.evaluate('({...__mf.view})')==chosen
    page.click('#pk-next')
    page.wait_for_function("__mf.stage==='edit' && document.querySelector('#busy').classList.contains('hidden')",timeout=180000)
    assert page.evaluate('({...__mf.view})')==chosen
    report['counts'] = page.evaluate('__mf.edit.counts')
    report['sid'] = page.evaluate('__mf.sid')
    print('Pick -> edit camera retained',flush=True)
    page.click('#ed-next')
    page.wait_for_function("__mf.stage==='design' && !!__mf.design")
    page.wait_for_timeout(500)
    assert page.evaluate('({...__mf.view})')==chosen
    page.click('#dg-back')
    page.wait_for_function("__mf.stage==='edit'")
    assert page.evaluate('({...__mf.view})')==chosen
    page.locator('#steps div').filter(has_text='도면 열기').click()
    page.wait_for_function("__mf.stage==='open'")
    assert page.evaluate('({...__mf.view})')==chosen
    page.screenshot(path=str(OUT/'viewport_live_zoomed.png'))
    report['stage_roundtrip_camera_unchanged']=True
    report['page_errors']=errors
    assert not errors,errors
    (OUT/'viewport_live.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False),flush=True)
    browser.close()
