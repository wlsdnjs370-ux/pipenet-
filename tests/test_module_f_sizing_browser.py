"""Use the real UI with isolated HTTP and the new solver, not the live session."""
from urllib.parse import urlparse

from test_module_f_viewport_browser import ROOT, install_page, pw_api
from test_module_f_sizing import client


def test_sizing_form_result_and_edit_invalidation(client,tmp_path):
    c,s=client
    with pw_api.sync_playwright() as pw:
        browser=pw.chromium.launch()
        page=browser.new_page(viewport={'width':1450,'height':1000})
        source=(ROOT/'static/module_f.js').read_text(encoding='utf-8')
        source=source.replace('  setStage("open");\n  loadSaved();',
            '  window.__sizingTest={refreshSizing};\n  setStage("open");\n  loadSaved();')
        errors=install_page(page,source)
        def route(req):
            url=urlparse(req.request.url)
            # The normal UI job watcher uses polling in this isolated test.
            if url.path.endswith('/job'):
                req.fulfill(json={'ok':True,'state':'done','lines':[],'elapsed':0})
                return
            if url.path.endswith('/job/stream'):
                req.fulfill(status=404,body='');return
            response=c.open(url.path+('?' + url.query if url.query else ''),
                method=req.request.method,data=req.request.post_data,content_type='application/json')
            if response.status_code==404:
                req.fulfill(json={'ok':True,'items':[],'rows':[],'fields':{}})
            else:
                req.fulfill(status=response.status_code,body=response.data,content_type=response.content_type)
        page.route('http://module-f.test/api/module-f/**',route)
        page.evaluate('''sid=>{__mf.sid=sid;__viewTest.setStage('merge');document.querySelector('#mg-sizing').open=true;}''',s['id'])
        page.locator('#mf-advanced-merge > summary').click()
        page.click('#sz-load')
        page.wait_for_function("document.querySelector('#sz-basis').textContent.includes('배관 5개')")
        assert '배관 5개' in page.locator('#sz-basis').inner_text()
        page.check('#sz-confirm')
        page.click('#sz-run')
        # Watcher may be waiting on its fallback interval; request the real
        # result as well, using the same refresh function used after a job.
        page.wait_for_function("document.querySelector('#sz-result').textContent.includes('조건 충족')",timeout=20000)
        assert s['sizing_report']['feasible']
        assert page.locator('#sz-emit').is_enabled()
        assert '원본' in page.locator('#sz-result').inner_text()
        page.locator('#sz-pressure').fill('7')
        assert page.locator('#sz-emit').is_disabled()
        assert '다시 계산' in page.locator('#sz-result').inner_text()
        tbl = s['merged']['combined']
        tbl.meta = [('부속 판정 불가','1')]
        tbl.unresolved = {'kind_items':[dict(pipe_label='1',node_label='a',n=1)]}
        page.click('#sz-load')
        page.wait_for_function("document.querySelector('#sz-basis').textContent.includes('부속 미확정 1건')")
        page.check('#sz-confirm')
        page.click('#sz-run')
        page.wait_for_function("document.querySelector('#sz-result').textContent.includes('역산을 시작하지 못했습니다')")
        assert '노드 a / 배관 1' in page.locator('#sz-result').inner_text()
        assert page.locator('#sz-emit').is_disabled()
        page.screenshot(path=str(tmp_path/'hydraulic_sizing_ui.png'),full_page=True)
        assert not errors,errors
        browser.close()
