"""Real browser + real Flask editor endpoints; no live user session touched."""
import json
import os
from pathlib import Path
from urllib.parse import urlparse

import pytest
from flask import Flask

from test_module_f_viewport_browser import install_page, settle, pw_api
from test_module_f_network_editor import design
from routes.module_f import api_design, api_merge, api_network_edit, jobs, network_edit as ne


@pytest.fixture()
def editor_ui(tmp_path,monkeypatch):
    monkeypatch.setattr(ne,'HISTORY_DIR',tmp_path/'history')
    sess=jobs._new_session(key='browser-editor')
    sess['design']=design()
    app=Flask(__name__);app.testing=True
    api_design.register(app,UPLOAD_DIR=tmp_path)
    api_merge.register(app,UPLOAD_DIR=tmp_path)
    api_network_edit.register(app)
    client=app.test_client()
    source=(Path(__file__).resolve().parents[1]/'static/module_f.js').read_text(encoding='utf-8')
    source=source.replace('  setStage("open");\n  loadSaved();',
        '  window.__editUi = {select:insSelect, preview:designPreview, merged:loadMergeView, overlap:underlayOverlap};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser=pw.chromium.launch()
        page=browser.new_page(viewport={'width':1500,'height':1100})
        page.set_default_timeout(7000)
        errors=install_page(page,source)
        def route(req):
            url=urlparse(req.request.url)
            body=req.request.post_data
            response=client.open(url.path+('?' + url.query if url.query else ''),
                method=req.request.method,data=body,content_type='application/json')
            if response.status_code==404:
                req.fulfill(json={'ok':True,'items':[],'rows':[],'fields':{}})
            else:
                req.fulfill(status=response.status_code,body=response.data,
                            content_type=response.content_type)
        page.route('http://module-f.test/api/module-f/**',route)
        page.evaluate('''async sid => {
          Object.assign(__mf,{sid,slot:'plan',method:'manual'});
          __viewTest.setStage('design'); await __editUi.preview();
          __editUi.select('node','2');
        }''',sess['id'])
        page.wait_for_function("document.querySelector('#ne-material').options.length > 0")
        page.click('#ne-focus')
        page.evaluate("moduleFNetworkEditor.openAction('extend')")
        yield page,sess
        assert not errors,errors
        browser.close()
    jobs._SESSIONS.pop(sess['id'],None)


def apply(page):
    page.click('#ne-preview')
    page.locator('#ne-preview-canvas').wait_for(state='visible')
    assert page.locator('#ne-preview-canvas').is_visible()
    if os.environ.get('NETWORK_EDITOR_SCREENSHOT'):
        page.screenshot(path=os.environ['NETWORK_EDITOR_SCREENSHOT'],full_page=True)
    page.click('#ne-apply')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')")
    settle(page)


def test_draw_continuously_undo_redo_and_keep_zoom(editor_ui):
    page,sess=editor_ui
    camera=page.evaluate('({...__mf.view})')
    page.select_option('#ne-axis','+Y')
    page.fill('#ne-length','2')
    apply(page)
    assert len(sess['design']['tables'].pipes)==3
    assert page.evaluate('__mf.design.sel.label')=='4'
    # New endpoints outside the viewport pan into view; the zoom never resets.
    assert page.evaluate('__mf.view.scale')==camera['scale']
    page.evaluate("moduleFNetworkEditor.openAction('extend')")
    page.select_option('#ne-axis','+Z')
    apply(page)
    assert len(sess['design']['tables'].pipes)==4
    page.locator('#ne-history').focus()
    page.keyboard.press('Control+z')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')")
    assert len(sess['design']['tables'].pipes)==3
    page.keyboard.press('Control+Shift+z')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')")
    assert len(sess['design']['tables'].pipes)==4


def test_pipe_material_library_and_fitting_ui(editor_ui):
    page,sess=editor_ui
    page.evaluate("__editUi.select('pipe','P2')")
    page.wait_for_function("document.querySelector('#ne-op').value==='resize'")
    page.evaluate("moduleFNetworkEditor.openAction('pipe')")
    page.select_option('#ne-material','CPVC2')
    page.select_option('#ne-dn','25')
    assert '28.02' in page.locator('#ne-inner').inner_text()
    apply(page)
    assert sess['design']['tables'].pipes[1]['inner_mm']==28.02
    page.evaluate("moduleFNetworkEditor.openAction('fitting')")
    page.select_option('#ne-fitting','ELBOW_45')
    page.fill('#ne-count','2')
    assert '0.6m' in page.locator('#ne-loss').inner_text()
    apply(page)
    assert sess['design']['tables'].equipment[-1]['eq_len']==pytest.approx(.6)


def test_preview_error_keeps_original_network(editor_ui):
    page,sess=editor_ui
    page.select_option('#ne-axis','+X')
    page.fill('#ne-length','1')
    page.click('#ne-preview')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')")
    assert page.locator('#ne-apply').is_disabled()
    assert page.locator('#ne-message').inner_text()
    assert len(sess['design']['tables'].pipes)==2


def test_edit_entry_switches_from_plan_to_calculation_view(editor_ui):
    page,sess=editor_ui
    page.evaluate("document.querySelector('#dg-plan').checked=true")
    page.click('#ne-focus')
    assert not page.locator('#dg-plan').is_checked()
    assert page.locator('#ne-toolbar').is_visible()


def canvas_point(page,kind,label):
    return page.evaluate('''({kind,label})=>{
      const v=__mf.stage==='merge'?__mf.mergeView:__mf.design.view;
      const at=Object.fromEntries(v.nodes.map(n=>[n.label,n]));
      const p=v.pipes.find(p=>String(p.label)===label);
      const n=kind==='node'?at[label]:{x:(at[p.a].x+at[p.b].x)/2,y:(at[p.a].y+at[p.b].y)/2};
      const c=document.querySelector('#cv'),r=c.getBoundingClientRect(),cam=__mf.view;
      return {x:r.x+(n.x-cam.ox)*cam.scale,y:r.y+r.height-(n.y-cam.oy)*cam.scale};
    }''',dict(kind=kind,label=label))


def click_element(page,kind,label,button='left'):
    p=canvas_point(page,kind,label)
    page.mouse.click(p['x'],p['y'],button=button)
    settle(page)


def test_node_arrows_pipe_selection_right_click_and_direct_confirm(editor_ui):
    page,sess=editor_ui
    page.click('#ne-close')
    page.evaluate("document.querySelector('#dg-iso').checked=true; __editUi.preview()")
    click_element(page,'node','2')
    assert page.locator('#ne-gizmo').is_visible()
    assert page.locator('.ne-arrow:visible').count()==6
    assert page.locator('#ne-panel').is_hidden()
    page.click('.ne-arrow[data-axis="-Y"]')
    page.fill('#ne-length','3')
    if os.environ.get('NETWORK_EDITOR_DIRECT_SCREENSHOT'):
        page.screenshot(path=os.environ['NETWORK_EDITOR_DIRECT_SCREENSHOT'],full_page=True)
    # Enter alone validates and commits. No card selection or mandatory two-click flow.
    page.keyboard.press('Enter')
    page.wait_for_function("__mf.design.sel?.label==='4' && document.querySelector('#busy').classList.contains('hidden')")
    assert sess['design']['got']['kfp']['pipe_data']['EP1']['length_m']==3
    assert page.locator('#ne-panel').is_hidden()
    assert page.locator('#ne-gizmo').is_visible()
    page.click('#btn-fit');settle(page)
    click_element(page,'pipe','EP1')
    assert page.locator('#ne-gizmo').is_hidden(), page.evaluate('''p=>({p,selection:__mf.design.sel,element:document.elementFromPoint(p.x,p.y)?.outerHTML.slice(0,300)})''',canvas_point(page,'pipe','EP1'))
    click_element(page,'pipe','EP1',button='right')
    page.get_by_role('menuitem',name='배관길이 변경',exact=True).click()
    page.fill('#ne-length','4')
    page.click('#ne-apply')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')")
    assert sess['design']['got']['kfp']['pipe_data']['EP1']['length_m']==4


def test_node_context_copy_paste_and_disabled_items(editor_ui):
    page,sess=editor_ui
    page.click('#ne-close')
    page.click('#btn-fit');settle(page)
    click_element(page,'node','3',button='right')
    assert page.locator('#ne-context').is_visible(),page.evaluate('''p=>({p,sel:__mf.design.sel,element:document.elementFromPoint(p.x,p.y)?.outerHTML.slice(0,300),menu:document.querySelector('#ne-context').innerHTML})''',canvas_point(page,'node','3'))
    assert page.get_by_role('menuitem',name='노드 삭제 (배관 합치기)',exact=True).is_disabled()
    page.get_by_role('menuitem',name='복사',exact=True).click()
    click_element(page,'node','2',button='right')
    assert page.get_by_role('menuitem',name='자르기',exact=True).is_disabled()
    page.get_by_role('menuitem',name='붙여넣기',exact=True).click()
    page.select_option('#ne-axis','-Y');page.fill('#ne-length','3')
    page.click('#ne-apply')
    page.wait_for_function("__mf.design.sel?.label==='4' && document.querySelector('#busy').classList.contains('hidden')")
    assert len(sess['design']['tables'].nozzles)==2
    click_element(page,'node','4',button='right')
    page.get_by_role('menuitem',name='노드속성 변경 ›',exact=True).click()
    page.get_by_role('menuitem',name='일반 노드로 변경',exact=True).click()
    page.click('#ne-apply')
    page.wait_for_function("document.querySelector('#busy').classList.contains('hidden')")
    assert len(sess['design']['tables'].nozzles)==1


def test_gizmo_tracks_zoom_and_45_degree_option_is_display_only(editor_ui):
    page,sess=editor_ui
    page.click('#ne-close')
    page.evaluate("document.querySelector('#dg-iso').checked=true; __editUi.preview()")
    page.click('#btn-fit');settle(page)
    click_element(page,'node','2',button='right')
    before=json.dumps(sess['design']['tables'].as_dict(),sort_keys=True)
    page.get_by_role('menuitemcheckbox').click()
    assert page.get_by_role('menuitemcheckbox').get_attribute('aria-checked')=='true'
    assert json.dumps(sess['design']['tables'].as_dict(),sort_keys=True)==before
    page.keyboard.press('Escape')
    page.evaluate('__mf.view.ox+=350/__mf.view.scale;__mf.view.scale*=1.02;__viewTest.draw()')
    settle(page)
    expected=canvas_point(page,'node','2')
    actual=page.locator('#ne-gizmo').bounding_box()
    assert actual['x']==pytest.approx(expected['x'],abs=1)
    assert actual['y']==pytest.approx(expected['y'],abs=1)


def test_merged_canvas_node_arrows_edit_original_plan(editor_ui):
    page,sess=editor_ui
    from test_module_f_merge import _riser
    from routes.module_f.api_merge import rebuild_merged
    sess['slots']['system']['riser']=_riser();sess['supply_mode']='lsp_gravity'
    rebuild_merged(sess)
    page.evaluate("async()=>{__viewTest.setStage('merge'); document.querySelector('#mg-iso').checked=true; await __editUi.merged();}")
    page.click('#btn-fit');settle(page)
    click_element(page,'node','11')
    page.locator('.ne-arrow[data-axis="+Y"]').click()
    page.fill('#ne-length','1')
    page.click('#ne-apply')
    page.wait_for_function("__mf.mergeSel?.label==='13' && document.querySelector('#busy').classList.contains('hidden')")
    assert len(sess['design']['tables'].pipes)==3
    assert any(p['label']=='EP1' for p in sess['merged']['combined'].pipes)
    assert page.locator('#ne-gizmo').is_visible()


def test_merge_property_card_and_three_independent_underlays(editor_ui):
    page,sess=editor_ui
    from test_module_f_merge_underlays import drawings
    drawings(sess)
    page.evaluate("async()=>{__viewTest.setStage('merge');document.querySelector('#mg-iso').checked=true;await __editUi.merged();}")
    page.click('#btn-fit');settle(page)
    click_element(page,'node','11')
    card=page.locator('#dg-ins')
    assert card.is_visible()
    card_box=card.bounding_box();stage=page.locator('#stage').bounding_box()
    assert card_box['x']+card_box['width']==pytest.approx(stage['x']+stage['width']-12,abs=2)
    assert card_box['y']+card_box['height']==pytest.approx(stage['y']+stage['height']-35,abs=2)
    assert '통합 절점' in card.inner_text()
    page.evaluate("__editUi.select('mgpipe','P2')")
    assert '통합 배관' in card.inner_text() and 'KSD 3507' in card.inner_text()
    assert page.locator('#ne-gizmo').is_hidden()
    before=page.locator('#cv').screenshot()
    calls=[]
    page.on('request',lambda r:calls.append(r.url) if 'underlays=1' in r.url else None)
    page.check('#mg-under')
    page.wait_for_function("document.querySelector('#mg-under-note').textContent.includes('펌프')")
    settle(page)
    assert page.locator('#cv').screenshot()!=before
    for kind in ['plan','system','machineroom']:
        assert page.locator(f'#mg-under-{kind}').is_enabled()
        page.uncheck(f'#mg-under-{kind}')
    settle(page)
    assert page.locator('#cv').screenshot()==before
    page.check('#mg-under-plan');settle(page)
    assert page.locator('#cv').screenshot()!=before
    page.uncheck('#mg-under');settle(page)
    assert page.locator('#cv').screenshot()==before
    page.check('#mg-under');settle(page)
    assert page.locator('#cv').screenshot()!=before
    assert len(calls)==1
    # Changing projection retains both the selection and drawing visibility.
    page.uncheck('#mg-iso');settle(page)
    assert card.is_visible() and 'P2' in page.locator('#dg-ins-title').inner_text()
    assert page.locator('#mg-under-plan').is_checked() and len(calls)==1
    if os.environ.get('NETWORK_EDITOR_MERGE_SCREENSHOT'):
        page.screenshot(path=os.environ['NETWORK_EDITOR_MERGE_SCREENSHOT'],full_page=True)


def test_merge_underlay_frames_follow_visibility_and_keep_screen_dash_size(editor_ui):
    page,sess=editor_ui
    from test_module_f_merge_underlays import drawings
    drawings(sess)
    page.evaluate("async()=>{__viewTest.setStage('merge');await __editUi.merged();}")
    # Observe actual canvas strokes without replacing the rendering operation.
    page.evaluate('''()=>{
      const ctx=document.querySelector('#cv').getContext('2d'),stroke=ctx.stroke;
      window.__frames=[];
      ctx.stroke=function(...args){
        if (this.strokeStyle==='#ffffff' && this.globalAlpha===.75 && this.getLineDash().length){
          const t=this.getTransform(),scale=Math.hypot(t.a,t.b)/devicePixelRatio;
          __frames.push({width:this.lineWidth*scale,dash:this.getLineDash().map(x=>x*scale)});
        }
        return stroke.apply(this,args);
      };
    }''')
    page.check('#mg-under');settle(page)
    def frames():
        return page.evaluate('''()=>{__frames=[];__viewTest.paint();return __frames;}''')
    page.wait_for_function('window.__frames.length>=3')
    for iso in (False,True):
        page.locator('#mg-iso').set_checked(iso);settle(page)
        for zoom in (1,2):
            page.evaluate('zoom=>{__mf.view.scale*=zoom;}',zoom)
            rows=frames()
            assert len(rows)==3
            for row in rows:
                assert row['width']==pytest.approx(1)
                assert row['dash']==pytest.approx([6,4])
        # Every source independently owns its frame; the master hides all.
        for count,kind in zip((2,1,0),('plan','system','machineroom')):
            page.uncheck('#mg-under-'+kind)
            assert len(frames())==count
        for kind in ('plan','system','machineroom'):
            page.check('#mg-under-'+kind)
        page.uncheck('#mg-under')
        assert frames()==[]
        page.check('#mg-under')
        assert len(frames())==3
    # Dashed framing must not leak into subsequent network strokes.
    assert page.evaluate("document.querySelector('#cv').getContext('2d').getLineDash()") == []
    if os.environ.get('NETWORK_EDITOR_FRAME_SCREENSHOT'):
        page.click('#btn-fit');settle(page)
        page.screenshot(path=os.environ['NETWORK_EDITOR_FRAME_SCREENSHOT'],full_page=True)


def test_underlay_overlap_geometry_and_visible_gray_boundaries(editor_ui):
    page,sess=editor_ui
    square=[[0,0],[10,0],[10,10],[0,10]]
    cases=[
        ([[5,5],[15,5],[15,15],[5,15]],25),
        ([[2,2],[8,2],[8,8],[2,8]],36),
        ([[20,0],[30,0],[30,10],[20,10]],0),
        ([[10,0],[20,0],[20,10],[10,10]],0),
        ([[5,-5],[15,5],[5,15],[-5,5]],100),
        ([],0),
    ]
    for clip,area in cases:
        for polygon in (clip,list(reversed(clip))):
            points=page.evaluate('([a,b])=>__editUi.overlap(a,b)',[square,polygon])
            actual=abs(sum(p[0]*points[(i+1)%len(points)][1]-points[(i+1)%len(points)][0]*p[1]
                           for i,p in enumerate(points)))/2 if points else 0
            assert actual==pytest.approx(area)
    from test_module_f_merge_underlays import drawings
    drawings(sess)
    # Broad rectangular drawing extents, unlike the deliberately tiny default
    # fixture. Production transforms still come from the actual Flask route.
    for slot in [sess,sess['slots']['system'],sess['slots']['machineroom']]:
        slot['world']={'bounds':dict(minx=0,miny=0,maxx=100,maxy=100),
                       'bundles':[dict(segs=[0,0,100,0,100,0,100,100],circles=[],arcs=[])]}
    page.evaluate("async()=>{__viewTest.setStage('merge');await __editUi.merged();}")
    page.check('#mg-under');settle(page)
    page.wait_for_function("document.querySelector('#mg-under-note').textContent.includes('펌프')")
    page.evaluate('''()=>{
      const ctx=document.querySelector('#cv').getContext('2d'),stroke=ctx.stroke;
      window.__overlaps=[];
      ctx.stroke=function(...args){
        if (this.strokeStyle==='#9ca3af') {
          const t=this.getTransform(),scale=Math.hypot(t.a,t.b)/devicePixelRatio;
          __overlaps.push({width:this.lineWidth*scale,dash:this.getLineDash().map(x=>x*scale)});
        }
        return stroke.apply(this,args);
      };
      // Controlled displayed positions: plan/system overlap, machine is apart.
      __mf.mergeView.underlays.forEach((r,i)=>{r.matrix=[1,0,0,1,i===2?300:50*i,0];});
    }''')
    def outlines():
        return page.evaluate('()=>{__overlaps=[];__viewTest.paint();return __overlaps;}')
    for zoom in (1,2,.5):
        page.evaluate('zoom=>{__mf.view.scale*=zoom;}',zoom)
        rows=outlines()
        assert len(rows)==1
        assert rows[0]['width']==pytest.approx(1.2)
        assert rows[0]['dash']==pytest.approx([5,4])
    page.uncheck('#mg-under-system')
    assert outlines()==[]
    page.check('#mg-under-system')
    assert len(outlines())==1
    # Reprojection invalidates cached corners as well as the drawing path.
    page.evaluate("__mf.mergeView.underlays[1].matrix[4]=500")
    assert outlines()==[]
    page.evaluate("__mf.mergeView.underlays[1].matrix[4]=50")
    assert len(outlines())==1
    page.uncheck('#mg-under')
    assert outlines()==[]
    assert page.evaluate("document.querySelector('#cv').getContext('2d').getLineDash()") == []


def test_merge_ui_system_then_plan_edit_no_longer_rejected(editor_ui):
    page,sess=editor_ui
    from test_module_f_merge_editor_followup import combined
    combined(sess)
    page.evaluate("async()=>{__viewTest.setStage('merge');await __editUi.merged();}")
    for label,dn in [('r1','50'),('P2','25')]:
        page.evaluate("label=>__editUi.select('mgpipe',label)",label)
        page.wait_for_function("document.querySelector('#ne-op').value==='resize'")
        page.evaluate("moduleFNetworkEditor.openAction('pipe')")
        page.select_option('#ne-material','CPVC2');page.select_option('#ne-dn',dn)
        apply(page)
    rows={r['label']:r for r in sess['merged']['combined'].pipes}
    assert rows['r1']['type']==rows['P2']['type']=='CPVC2'
    assert sess['design']['tables'].pipes[1]['type']=='CPVC2'
    assert page.locator('#dg-ins').is_visible()
