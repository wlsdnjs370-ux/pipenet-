"""Display budgets must not erase late DXF layers or affect extraction data."""
from copy import deepcopy
from types import SimpleNamespace

import pytest

from routes.module_f.common import _boot
from routes.module_f import world as display
from src.pipenet_converter.render.display_budget import bundle_quotas, display_index

_boot()


def drawing():
    return SimpleNamespace(
        segs=[('background',7,(i,0),(i,100)) for i in range(40)] +
             [('late-loop',5,(100,100),(200,100)),('late-loop',5,(200,100),(200,200)),
              ('late-loop',5,(200,200),(100,200)),('late-loop',5,(100,200),(100,100))],
        circles=[('first',2,0,0,10)]*10+[('late-head',3,150,150,10)],
        arcs=[('first',2,0,0,10)]*10+[('late-arc',1,200,200,10)],
        arc_ang=[(0,90)]*10+[(45,180)])


def test_late_small_bundles_preserved_and_source_unchanged(monkeypatch):
    monkeypatch.setattr(display,'MAX_SEGS',12)
    for name in ('MAX_CIRCLES','MAX_ARCS'):
        monkeypatch.setattr(display,name,4)
    raw=drawing();before=deepcopy(raw.__dict__)
    got=display._world_payload(raw)
    bundles={b['layer']:b for b in got['bundles']}
    assert got['shown']['segs']==12
    assert got['shown']['circles']==got['shown']['arcs']==4
    assert bundles['late-loop']['n_seg']==bundles['late-loop']['n_all']==4
    assert not bundles['late-loop']['display_partial']
    assert bundles['background']['display_partial']
    assert bundles['late-head']['n_circle']==1
    assert bundles['late-arc']['arcs']==[200,200,10,45,180]
    assert raw.__dict__==before
    # Every layer is listed even if an extremely small display budget cannot draw it.
    monkeypatch.setattr(display,'MAX_SEGS',0)
    assert {b['layer'] for b in display._world_payload(raw)['bundles']}==set(bundles)


def test_quotas_independent_of_layer_insertion_order():
    counts={'huge':300000,'later':4,'heads':500,'empty':0}
    assert bundle_quotas(counts,1000)==bundle_quotas(dict(reversed(list(counts.items()))),1000)
    assert bundle_quotas(counts,1000)['later']==4
    for total in range(1,45):
        for quota in range(total+1):
            indices=[i for i in range(total) if display_index(i,total,quota)]
            assert len(indices)==quota
            if quota>1: assert indices[0]==0 and indices[-1]==total-1


@pytest.fixture
def display_client(monkeypatch):
    from flask import Flask
    from routes.module_f import api_open,jobs
    monkeypatch.setattr(display,'MAX_SEGS',12)
    raw=drawing()
    sess=jobs._new_session(key='display-regression')
    sess.update(pick=SimpleNamespace(world=raw),world=display._world_payload(raw))
    app=Flask(__name__);app.testing=True
    api_open.register(app,_save_upload=lambda *a,**k:None)
    yield app.test_client(),sess
    jobs._SESSIONS.pop(sess['id'],None)


def test_full_bundle_route_uses_current_cropped_world_only(display_client):
    c,sess=display_client
    query={'sid':sess['id'],'bundle_id':'background\x1f7'}
    got=c.get('/api/module-f/world/bundle',query_string=query)
    assert got.status_code==200,got.json
    assert got.json['bundle']['n_seg']==40 and not got.json['bundle']['display_partial']
    assert sess['world']['shown']['segs']==12  # Read-only; does not change another reader's display.
    # Replacing the working world by a cropped one must not reload excluded DXF geometry.
    sess['pick'].world.segs=sess['pick'].world.segs[:2]
    got=c.get('/api/module-f/world/bundle',query_string=query)
    assert got.json['bundle']['n_seg']==2
    query['bundle_id']='not-in-drawing'
    assert c.get('/api/module-f/world/bundle',query_string=query).status_code==404


def test_browser_full_bundle_button(display_client,tmp_path):
    from urllib.parse import urlparse
    from test_module_f_viewport_browser import install_page,pw_api,ROOT
    c,sess=display_client
    source=(ROOT/'static/module_f.js').read_text(encoding='utf-8').replace(
        '  setStage("open");\n  loadSaved();',
        '  window.__displayTest={buildLayers};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser=pw.chromium.launch();page=browser.new_page(viewport={'width':1440,'height':1000})
        errors=install_page(page,source)
        def route(req):
            url=urlparse(req.request.url)
            res=c.get(url.path+'?'+url.query)
            req.fulfill(status=res.status_code,body=res.data,content_type=res.content_type)
        page.route('http://module-f.test/api/module-f/world/bundle?**',route)
        page.evaluate('''({sid,world})=>{Object.assign(__mf,{sid,world,key:'display-regression',slot:'plan'});
            __viewTest.setStage('pick');__displayTest.buildLayers();}''',{'sid':sess['id'],'world':sess['world']})
        assert page.locator('.layer-complete').count()==1
        # Open only the layers panel; do not change the selected pipe bundles.
        page.locator('#layers').evaluate("el=>{el.classList.remove('hidden');let p=el.parentElement;while(p){p.classList.remove('hidden');p=p.parentElement;}}")
        page.locator('.layer-complete').click()
        page.wait_for_function("__mf.world.bundles.find(b=>b.layer==='background').n_seg===40")
        assert page.locator('.layer-complete').count()==0
        assert page.evaluate('__mf.world.dropped.segs')==0
        page.screenshot(path=str(tmp_path/'complete-layer.png'))
        assert not errors,errors
        browser.close()
