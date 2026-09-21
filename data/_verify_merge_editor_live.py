import hashlib,json,re,urllib.request
from pathlib import Path
from playwright.sync_api import sync_playwright
root=Path.cwd();base='http://192.168.0.105:5051'
html=urllib.request.urlopen(base+'/login',timeout=15).read().decode('utf-8')
hint=re.search(r'<span class="pw-value">([^<]+)<',html)
errors=[]
with sync_playwright() as pw:
    browser=pw.chromium.launch()
    page=browser.new_page(viewport={'width':1500,'height':1000})
    page.on('pageerror',lambda e:errors.append(str(e)))
    page.goto(base+'/module-f')
    if page.query_selector('input[type=password]'):
        assert hint,'Local login hint unavailable'
        page.fill('input[type=password]',hint.group(1))
        page.click('button[type=submit]');page.wait_for_load_state('load')
    page.wait_for_function('!!window.moduleFNetworkEditor')
    files={}
    for name in ('module_f.js','module_f_editor.js','module_f.css'):
        response=page.request.get(base+'/static/'+name)
        files[name]=hashlib.sha256(response.body()).hexdigest()==hashlib.sha256((root/'static'/name).read_bytes()).hexdigest()
    response=page.request.get(base+'/api/module-f/network-editor?sid=deployment-check')
    assert response.status==410,response.status
    ids=page.evaluate("['ne-focus','ne-panel','ne-preview','ne-apply','ne-undo','ne-redo','ne-pick-end','ne-pick-split','ne-gizmo','ne-arrows','ne-context','ne-close','ne-coordinates','mg-under-plan','mg-under-system','mg-under-machineroom','mg-under-note'].every(id=>!!document.getElementById(id))")
    assert page.locator('.ne-arrow').count()==6
    assert '20260916-merge-editor' in page.content()
    assert all(files.values()) and ids and not errors,(files,ids,errors)
    result={'ok':True,'url':page.url,'assets_match':files,'editor_controls':ids,'route_status':response.status,'page_errors':errors}
    (root/'data/direct_editor_validation/merge_local_server.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))
    browser.close()

