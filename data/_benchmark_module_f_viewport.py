"""Read an existing test session and compare old/new Canvas rendering locally."""
from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import statistics

from playwright.sync_api import sync_playwright

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / 'cad_project_editor_g/docs/benchmarks'
spec = importlib.util.spec_from_file_location('view_tests', ROOT / 'tests/test_module_f_viewport_browser.py')
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
sid = json.loads((OUT / 'cancel_live_session.json').read_text(encoding='utf-8'))['sid']

with sync_playwright() as pw:
    browser = pw.chromium.launch()
    page = browser.new_page(viewport={'width':1680,'height':1000})
    page.goto('http://127.0.0.1:5051/module-f')
    if '/login' in page.url:
        page.fill('input[type=password]', page.locator('.pw-value').inner_text())
        page.click('button[type=submit]')
        page.wait_for_url('**/module-f')
    payload = {}
    for key, endpoint in [('world','world'),('edit','edit/state')]:
        response = page.request.get(f'http://127.0.0.1:5051/api/module-f/{endpoint}?sid={sid}',timeout=120000)
        assert response.ok, response.status
        payload[key] = response.json()
    page.close()
    (OUT / 'viewport_fixture.json').write_text(json.dumps(payload,ensure_ascii=False),encoding='utf-8')
    report = {'drawing':payload['world']['key'],'counts':payload['world']['world']['counts']}
    for label, filename in [('before','data/_module_f_before_viewport.js'),('after','static/module_f.js')]:
        page = browser.new_page(viewport={'width':1680,'height':1000})
        errors = mod.install_page(page,(ROOT / filename).read_text(encoding='utf-8'))
        page.evaluate("""p => {
            Object.assign(__mf,{sid:'bench',key:'B1F',world:p.world.world,pick:null,
              edit:p.edit.state,method:'manual',stage:'open'});
            document.querySelector('#btn-fit').click();
        }""",payload)
        mod.settle(page)
        stats = {}
        for zoom in [1,10]:
            page.click('#btn-fit')
            mod.settle(page)
            page.evaluate("""z => { const v=__mf.view, r=document.querySelector('#cv').getBoundingClientRect();
                const cx=z===10 ? 662692 : v.ox+r.width/v.scale/2;
                const cy=z===10 ? 175774 : v.oy+r.height/v.scale/2;
                v.scale*=z; v.ox=cx-r.width/v.scale/2; v.oy=cy-r.height/v.scale/2;
            }""",zoom)
            timings = page.evaluate("""() => {
                const samples=[], stage=document.querySelector('#stage');
                const original=stage.getBoundingClientRect.bind(stage); let layouts=0;
                stage.getBoundingClientRect=()=>{layouts++; return original()};
                const c=document.querySelector('#cv').getContext('2d');
                for(let i=0;i<14;i++) {
                    __mf.view.ox += 2/__mf.view.scale;
                    const t=performance.now(); __viewTest.paint(); c.getImageData(0,0,1,1);
                    if(i>=2) samples.push(performance.now()-t);
                }
                stage.getBoundingClientRect=original;
                return {samples, layouts};
            }""")
            stats[str(zoom)] = {'median_ms':round(statistics.median(timings['samples']),2),
                               'layout_reads':timings['layouts']}
            if label=='after' and zoom==10:
                page.screenshot(path=str(OUT / 'viewport_zoomed.png'))
        assert not errors, errors
        report[label] = stats
        page.close()
    browser.close()
    for zoom in ['1','10']:
        report['after'][zoom]['reduction_pct'] = round(100*(1-report['after'][zoom]['median_ms']/report['before'][zoom]['median_ms']),1)
    (OUT / 'viewport_performance.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))
