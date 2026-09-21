"""Real UI + isolated Flask: always-on fittings, hit testing, cards and gizmos."""
import json
import os
from urllib.parse import urlparse

import pytest
from flask import Flask

from test_module_f_fitting_inspection import fixture
from test_module_f_viewport_browser import ROOT, install_page, settle, pw_api
from test_module_f_network_editor_browser import click_element
from routes.module_f import api_design, api_merge, api_network_edit, jobs, network_edit as ne


@pytest.fixture()
def ui(tmp_path,monkeypatch):
    monkeypatch.setattr(ne,'HISTORY_DIR',tmp_path/'history')
    tbl,got,keys,board=fixture()
    tbl.nozzles=[dict(label='H1',**{'in':'4','out':'@/4'},flow_lmin=80,
                     flow_m3s=80/60000,lib='SP-HEAD',status='1')]
    got['kfp']['nodes_meta_runtime']['N4']['type_id']='head'
    sess=jobs._new_session(key='fitting-browser-fixture')
    from types import SimpleNamespace
    sess.update(design=dict(tables=tbl,got=got,keys=keys),edit=SimpleNamespace(board=board))
    from routes.module_f.merge import merge_network
    from test_module_f_merge_view import _riser
    sess.update(supply_mode='lsp_gravity',merged=merge_network(tbl,riser=_riser(),mode='lsp_gravity'))
    app=Flask(__name__);app.testing=True
    api_design.register(app,UPLOAD_DIR=tmp_path);api_network_edit.register(app)
    api_merge.register(app,UPLOAD_DIR=tmp_path)
    client=app.test_client()
    source=(ROOT/'static/module_f.js').read_text(encoding='utf-8')
    source=source.replace('  setStage("open");\n  loadSaved();',
        '  window.__fitTest={select:insSelect,preview:designPreview,inspect:designInspect,merge:loadMergeView,mergeInspect};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser=pw.chromium.launch()
        page=browser.new_page(viewport={'width':1500,'height':1100})
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
        page.evaluate('''async sid=>{
          Object.assign(__mf,{sid,slot:'plan',method:'manual'});
          document.querySelector('#dg-plan').checked=false;
          __viewTest.setStage('design');await __fitTest.preview();
          __viewTest.fitDesignView();__viewTest.draw();
        }''',sess['id'])
        settle(page)
        yield page,sess
        assert not errors,errors
        browser.close()
    jobs._SESSIONS.pop(sess['id'],None)


def strokes(page):
    return page.evaluate('''async ()=>{
      const ctx=document.querySelector('#cv').getContext('2d'),old=ctx.stroke,result=[];
      ctx.stroke=function(){result.push({color:this.strokeStyle,alpha:this.globalAlpha,dash:this.getLineDash()});return old.call(this);};
      try{__viewTest.draw();await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));}finally{ctx.stroke=old;}
      return result.filter(r=>r.color==='#ff616c');
    }''')


def test_all_fittings_visible_without_click_and_remain_visible_when_selected(ui):
    page,sess=ui
    before=json.dumps(sess['design']['tables'].as_dict(),sort_keys=True)
    assert page.evaluate('__mf.design.sel') is None
    assert page.locator('#dg-ins').is_hidden()
    red=strokes(page)
    assert len(red)==2 and all(r['dash']==[4,3] and r['alpha']==1 for r in red)
    click_element(page,'node','2')
    assert page.locator('#dg-ins').is_visible() and page.locator('#ne-gizmo').is_visible()
    assert page.locator('.ne-arrow:visible').count()==6
    text=page.locator('#dg-ins-body').inner_text()
    assert '분류티' in text and '3 → 2 포트' in text and 'TEE_BRANCH' in text
    assert '90° 엘보' not in text  # A downstream node's fitting must not leak here.
    assert len(strokes(page))==2
    page.click('#dg-ins-close')
    assert page.locator('#dg-ins').is_hidden() and len(strokes(page))==2
    assert json.dumps(sess['design']['tables'].as_dict(),sort_keys=True)==before


def test_omitted_tee_arm_is_clickable_and_pipe_head_cards_show_values(ui):
    page,_=ui
    # Hit the end of the omitted original arm, outside the node's pick radius.
    page.evaluate('''()=>{
      const f=__mf.design.view.inspection.fittings[0],arm=f.arms.find(a=>a[0]>0.8);
      __fitTest.inspect(f.x+arm[0]*17/__mf.view.scale,f.y+arm[1]*17/__mf.view.scale,8/__mf.view.scale);
    }''')
    assert page.evaluate('__mf.design.sel')=={'kind':'node','label':'2'}
    assert page.locator('#dg-ins').is_visible()
    page.evaluate("__fitTest.select('pipe','P2')")
    text=page.locator('#dg-ins-body').inner_text()
    assert '실제 길이' in text and 'KSD 3507' in text and '분류티' in text
    page.evaluate("__fitTest.select('node','4')")
    assert page.locator('#dg-ins-kind').inner_text()=='헤드'
    assert '노즐' in page.locator('#dg-ins-body').inner_text()


def test_card_does_not_cover_selected_xyz_handles_and_screenshot(ui):
    page,_=ui
    page.evaluate("__fitTest.select('node','2')")
    settle(page)
    # A fitting overlay must not erase the actual pipe network underneath it.
    assert page.evaluate('''()=>{
      const v=__mf.design.view,p=v.pipes[0],a=v.nodes.find(n=>n.label===p.a),b=v.nodes.find(n=>n.label===p.b);
      const cv=document.querySelector('#cv'),cam=__mf.view;
      const x=((a.x+b.x)/2-cam.ox)*cam.scale,y=cv.clientHeight-((a.y+b.y)/2-cam.oy)*cam.scale;
      const pixels=cv.getContext('2d').getImageData(Math.round(x)-2,Math.round(y)-2,5,5).data;
      return [...pixels].some((v,i)=>i%4!==3 && v>20);
    }''')
    card=page.locator('#dg-ins').bounding_box()
    for arrow in page.locator('.ne-arrow:visible').all():
        b=arrow.bounding_box()
        assert not (card['x'] < b['x']+b['width'] and card['x']+card['width'] > b['x']
                    and card['y'] < b['y']+b['height'] and card['y']+card['height'] > b['y'])
    if os.environ.get('MODULE_F_INSPECTION_SCREENSHOT'):
        page.screenshot(path=os.environ['MODULE_F_INSPECTION_SCREENSHOT'],full_page=True)


def test_fitting_geometry_follows_network_directions_and_camera_zoom(ui):
    page,_=ui
    result=page.evaluate('''()=>{
      const ctx=document.querySelector('#cv').getContext('2d');
      const move=ctx.moveTo,line=ctx.lineTo;let last,lines=[];
      ctx.moveTo=function(x,y){last=[x,y];return move.call(this,x,y);};
      ctx.lineTo=function(x,y){if(last)lines.push([x-last[0],y-last[1]]);return line.call(this,x,y);};
      try {
        window.moduleFInspection.draw(); const before=lines; lines=[];
        __mf.view.scale*=2;window.moduleFInspection.draw();__mf.view.scale/=2;
        return {before,after:lines,arms:__mf.design.view.inspection.fittings[0].arms};
      } finally {ctx.moveTo=move;ctx.lineTo=line;}
    }''')
    assert result['before']
    for a,b in zip(result['before'],result['after']):
        assert b == pytest.approx([2*a[0],2*a[1]])
    for delta,arm in zip(result['before'],result['arms']):
        # Screen Y is flipped. The marker follows the projected network arm.
        assert delta[0]*(-arm[1])-delta[1]*arm[0] == pytest.approx(0,abs=1e-7)


def test_incomplete_edited_tee_draws_known_arms_not_fixed_t_icon(ui):
    page,_=ui
    count=page.evaluate('''()=>{
      const f=__mf.design.view.inspection.fittings[0];
      f.symbolic=true;f.arms=[[.6,.8],[-.8,.6]];
      __mf.design.view.inspection.fittings=[f];
      const ctx=document.querySelector('#cv').getContext('2d'),old=ctx.lineTo;
      let count=0;ctx.lineTo=function(...v){count++;return old.apply(this,v);};
      try{window.moduleFInspection.draw();return count;}finally{ctx.lineTo=old;}
    }''')
    assert count==4  # two actual arms, drawn once for casing and once in red


@pytest.mark.parametrize("iso",[False,True])
def test_merged_always_on_glyphs_cards_links_and_zoom(ui,iso):
    page,sess=ui
    from dataclasses import asdict
    before=asdict(sess['merged']['combined'])
    page.evaluate('''async iso=>{
      document.querySelector('#mg-iso').checked=iso;
      __viewTest.setStage('merge');await __fitTest.merge();__viewTest.draw();
    }''',iso)
    page.click('#btn-fit');settle(page)
    assert page.evaluate('__mf.mergeSel || null') is None
    red=strokes(page)
    assert len(red)==2 and all(r['dash']==[4,3] for r in red)
    # The riser spans many floors. Zoom into the plan's tee before picking it.
    page.evaluate('''()=>{
      const n=__mf.mergeView.nodes.find(n=>n.label==='11'),cv=document.querySelector('#cv');
      const scale=.18;Object.assign(__mf.view,{scale,ox:n.x-cv.clientWidth/2/scale,oy:n.y-cv.clientHeight/2/scale});
      __viewTest.draw();
    }''')
    settle(page)
    click_element(page,'node','11')
    text=page.locator('#dg-ins-body').inner_text()
    assert '분류티' in text and '3 → 2 포트' in text and 'TEE_BRANCH' in text
    assert '직류티' not in text
    assert page.locator('.ne-arrow:visible').count()==6
    page.locator('#dg-ins-body [data-ins-kind="mgpipe"][data-ins-label="P2"]').first.click()
    assert page.evaluate('__mf.mergeSel')==dict(kind='pipe',label='P2')
    assert page.evaluate('__mf.design.sel || null') is None
    assert '분류티' in page.locator('#dg-ins-body').inner_text()
    page.click('#dg-ins-close')
    assert len(strokes(page))==2
    page.evaluate('''()=>{
      const f=__mf.mergeView.inspection.fittings.find(f=>f.kind==='tee');
      const arm=f.arms.find(a=>a[0]>.8);
      __fitTest.mergeInspect(f.x+arm[0]*17/__mf.view.scale,f.y+arm[1]*17/__mf.view.scale,8/__mf.view.scale);
    }''')
    assert page.evaluate('__mf.mergeSel')==dict(kind='node',label='11')
    sizes=page.evaluate('''()=>{
      const ctx=document.querySelector('#cv').getContext('2d'),move=ctx.moveTo,line=ctx.lineTo;
      let last,lines=[];
      ctx.moveTo=function(x,y){last=[x,y];return move.call(this,x,y);};
      ctx.lineTo=function(x,y){if(last)lines.push([x-last[0],y-last[1]]);return line.call(this,x,y);};
      try {
        moduleFMergeInspection.draw();const before=lines;lines=[];
        __mf.view.scale*=2;moduleFMergeInspection.draw();__mf.view.scale/=2;
        return {before,after:lines};
      } finally {ctx.moveTo=move;ctx.lineTo=line;}
    }''')
    for a,b in zip(sizes['before'],sizes['after']):
        assert b==pytest.approx([2*a[0],2*a[1]])
    assert asdict(sess['merged']['combined'])==before
    if os.environ.get('MODULE_F_MERGED_INSPECTION_SCREENSHOT'):
        page.screenshot(path=os.environ['MODULE_F_MERGED_INSPECTION_SCREENSHOT'],full_page=True)


def test_old_straight_tee_payload_never_draws_or_enters_cards(ui):
    page,_=ui
    page.evaluate('''()=>{
      const old={...__mf.design.view.inspection.fittings[0],kind:'tee-run',name:'직류티'};
      __mf.design.view.inspection.fittings=[old];__viewTest.draw();
      __fitTest.select('node','2');
    }''')
    assert not strokes(page)
    assert '직류티' not in page.locator('#dg-ins-body').inner_text()
