"""Same-floor choice is sent by the real UI and reaches the extractor."""
from pathlib import Path
from urllib.parse import urlparse

from flask import Flask
from test_module_f_viewport_browser import install_page, pw_api
from routes.module_f import api_sub, jobs


def test_same_floor_selection_extracts_without_inventing_height(tmp_path):
    app=Flask(__name__); app.testing=True
    api_sub.register(app)
    client=app.test_client()
    sess=jobs._new_session(key='same-level-browser')
    sess.update(active='system',entities=[dict(t='L',l='PIPE',p=[0,0,3000,0]),
                                         dict(t='L',l='PIPE',p=[3000,0,3000,4000])],
                sub_layers=['PIPE'],sub_layers_auto=False)
    source=(Path(__file__).resolve().parents[1]/'static/module_f.js').read_text(encoding='utf-8')
    source=source.replace('  setStage("open");\n  loadSaved();',
        '  window.__systemUi={panel:renderSubPanel,load:loadSub};\n  setStage("open");\n  loadSaved();')
    try:
        with pw_api.sync_playwright() as pw:
            browser=pw.chromium.launch()
            page=browser.new_page(viewport={'width':1400,'height':1000})
            errors=install_page(page,source)
            def route(req):
                url=urlparse(req.request.url)
                response=client.open(url.path+('?' + url.query if url.query else ''),
                    method=req.request.method,data=req.request.post_data,content_type='application/json')
                if response.status_code==404:
                    req.fulfill(json={'ok':True,'items':[],'rows':[],'fields':{}})
                else:
                    req.fulfill(status=response.status_code,body=response.data,content_type=response.content_type)
            page.route('http://module-f.test/api/module-f/**',route)
            page.evaluate('''sid=>{
              Object.assign(__mf,{sid,slot:'system'});
              __mf.sub.picks=[[0,0],[3000,4000]];
              __viewTest.setStage('sub');__systemUi.panel();
            }''',sess['id'])
            assert page.locator('#sub-elevation-row').is_visible()
            page.select_option('#sub-elevation-mode','same_level')
            page.click('#sub-extract')
            page.wait_for_function("document.querySelector('#sub-summary').textContent.includes('7 m')")
            assert sess['riser']['elevation_mode']=='same_level'
            assert all(n['elevation']==0 for n in sess['riser']['nodes'])
            page.evaluate('__systemUi.load()')
            assert page.locator('#sub-elevation-mode').input_value()=='same_level'
            page.evaluate("__mf.slot='machineroom';__systemUi.panel()")
            assert page.locator('#sub-elevation-row').is_hidden()
            assert not errors,errors
            browser.close()
    finally:
        jobs._SESSIONS.pop(sess['id'],None)
