"""Working crop is geometric, early, persisted, and never an original-file edit."""
from copy import deepcopy
from types import SimpleNamespace

import pytest

from src.pipenet_converter.dxf.work_region import WorkRegion, crop_world


def world():
    return SimpleNamespace(segs=[('P', 1, (-5, 5), (15, 5)), ('P', 1, (30, 5), (40, 5))],
        raw_segs=[('P', 1, (-5, 5), (15, 5))], circles=[('H', 2, 5, 5, 1), ('H', 2, 0, 5, 1)],
        arcs=[('A', 3, 5, 5, 1), ('A', 3, 15, 5, 1)], arc_ang=[(0, 90), (90, 90)],
        texts=[('T', 4, 5, 5, 1, '65'), ('T', 4, 20, 5, 1, '150')], _source_path='original.dxf')


def test_clip_counts_and_original_unchanged():
    w = world(); before = deepcopy(w.__dict__)
    got, report = crop_world(w, [0, 0, 10, 10])
    assert got.segs == [('P', 1, (0, 5), (10, 5))]
    assert len(got.circles) == len(got.arcs) == len(got.texts) == 1
    assert got.arc_ang == [(0, 90)] and got._source_path == 'original.dxf'
    assert report['boundary_symbols_excluded'] == 1
    assert w.__dict__ == before


def test_concave_region_does_not_fill_missing_middle():
    region = WorkRegion({'type':'polygon','points':[[0,0],[10,0],[10,10],[7,10],[7,3],[3,3],[3,10],[0,10]]})
    assert region.clip((-5,5), (15,5)) == [((0,5),(3,5)), ((7,5),(10,5))]
    assert not region.contains((5,5))
    assert region.clip((0,0), (10,0)) == [((0,0),(10,0))]


@pytest.mark.parametrize('zone', [None, [0,0,0,5], {'type':'polygon','points':[[0,0],[10,10],[0,8],[10,0]]}])
def test_invalid_region_rejected(zone):
    with pytest.raises((TypeError, ValueError)):
        WorkRegion(zone)


def test_nested_crop_replay_is_identical():
    first, _ = crop_world(world(), [0,0,10,10])
    second, _ = crop_world(first, [2,0,8,10])
    replay = world()
    for region in second._work_regions:
        replay, _ = crop_world(replay, region)
    assert replay.__dict__ == second.__dict__


@pytest.fixture
def crop_client(tmp_path,monkeypatch):
    import ezdxf
    from flask import Flask,jsonify
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pick.session import PickSession
    from routes.module_f import api_pick,api_convert,jobs
    from routes.module_f.world import _world_payload
    from routes.module_f.views import _pick_state
    doc=ezdxf.new();m=doc.modelspace()
    doc.layers.new('P',dxfattribs={'color':3})
    m.add_line((0,2000),(10000,2000),dxfattribs={'layer':'P'})
    m.add_line((15000,2000),(25000,2000),dxfattribs={'layer':'P'})
    m.add_circle((4000,2000),100,dxfattribs={'layer':'P'})
    path=tmp_path/'crop_fixture.dxf';doc.saveas(path)
    ps=PickSession.open(str(path));ps.select_pipe();ps.click(2000,2000);ps.complete_pipe()
    sess=jobs._new_session(key=ps.key)
    sess.update(pick=ps,dxf=str(path),world=_world_payload(ps.world),design={'stale':True},
                network_editor={'stale':True},design_review_path='old.json')
    app=Flask(__name__);app.testing=True
    api_pick.register(app);api_convert.register(app,UPLOAD_DIR=tmp_path)
    @app.get('/api/module-f/world')
    def world_route():
        return jsonify(ok=True,key=sess['key'],world=sess['world'],state=_pick_state(sess))
    def run(session,phase,fn):
        session['job']={'state':'done','result':fn(),'started':0,'ended':1,'phase':phase,'error':None}
    monkeypatch.setattr(api_pick,'_run_job',run)
    yield app.test_client(),sess,path
    jobs._SESSIONS.pop(sess['id'],None)


def test_crop_route_reset_and_persistent_early_pipeline(crop_client,tmp_path,monkeypatch):
    import hashlib
    from services.cad_import.pipeline import expand
    c,sess,path=crop_client
    original=hashlib.sha256(path.read_bytes()).hexdigest()
    old=sess['pick'].world
    response=c.post('/api/module-f/pick/crop',json={'sid':sess['id'],'zone':[1000,500,8000,3500]})
    assert response.status_code==200,response.json
    assert sess['job']['result']['ok'],sess['job']
    assert sess['pick'].world.segs==[('P',3,(1000.,2000.),(8000.,2000.))]
    assert len(old.segs)==2 and not sess.get('design') and not sess.get('network_editor')
    assert sess['pick'].mat_done and sess['pick'].board.mat
    spec_path=sess['pick'].commit(str(tmp_path/'spec'))
    monkeypatch.setattr(expand,'_spec_path',lambda key:spec_path)
    st=expand.stage1_body(sess['key'])
    assert st['w'].segs==sess['pick'].world.segs
    bad=c.post('/api/module-f/pick/crop',json={'sid':sess['id'],'zone':[30000,0,31000,1000]})
    assert not sess['job']['result']['ok']
    assert len(sess['pick'].world.segs)==1
    assert c.post('/api/module-f/pick/crop',json={'sid':sess['id'],'reset':True}).status_code==200
    assert sess['job']['result']['ok'] and len(sess['pick'].world.segs)==2
    assert not getattr(sess['pick'].world,'_work_regions',None)
    assert hashlib.sha256(path.read_bytes()).hexdigest()==original


def test_postprocess_cannot_extend_outside_crop(monkeypatch):
    from routes.module_f import api_pick
    sess={}
    ps=SimpleNamespace(world=SimpleNamespace(_work_regions=[[0,0,10,10]]))
    es=SimpleNamespace(board=SimpleNamespace(pts=[(0,5),(10,5)],edges={(0,1)}))
    def bad_fold(sess,ps,es):
        es.board.pts[1]=(20,5)
        sess['fold']={'folded':1}
    monkeypatch.setattr(api_pick,'_apply_fold',bad_fold)
    monkeypatch.setattr(api_pick,'_apply_chamfer',lambda *args:None)
    api_pick._postprocess_with_crop_guard(sess,ps,es)
    assert es.board.pts==[(0,5),(10,5)] and 'fold' not in sess


@pytest.mark.parametrize('stroke', ['simple', 'crossed_closure', 'double_outline'])
def test_browser_pen_crop_and_reset(crop_client,tmp_path,stroke):
    from urllib.parse import urlparse
    from test_module_f_viewport_browser import install_page,pw_api,ROOT
    c,sess,_=crop_client
    source=(ROOT/'static/module_f.js').read_text(encoding='utf-8')
    source=source.replace('  setStage("open");\n  loadSaved();',
        '  window.__cropTest={loadWorld};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser=pw.chromium.launch();page=browser.new_page(viewport={'width':1440,'height':1000})
        errors=install_page(page,source)
        def route(req):
            url=urlparse(req.request.url)
            if url.path.endswith('/job') or url.path.endswith('/progress'):
                req.fulfill(json={'ok':True,'job':sess.get('job',{}),'state':'done','phase':'완료'});return
            res=c.open(url.path+('?' + url.query if url.query else ''),method=req.request.method,
                       data=req.request.post_data,content_type='application/json')
            if res.status_code==404: req.fulfill(json={'ok':True,'items':[],'rows':[],'fields':{}})
            else: req.fulfill(status=res.status_code,body=res.data,content_type=res.content_type)
        page.route('http://module-f.test/api/module-f/**',route)
        page.evaluate('async sid=>{Object.assign(__mf,{sid,slot:"plan",method:"manual"});await __cropTest.loadWorld(false);}',sess['id'])
        page.click('#pk-crop-pen')
        points=[[1000,500],[8000,500],[9000,2000],[8000,3500],[1000,3500]]
        if stroke=='crossed_closure':
            points += [[900,400],[1400,700]]
        elif stroke=='double_outline':
            points += points
        coords=page.evaluate('ps=>ps.map(([x,y])=>[__mf.toScreenX(x),__mf.toScreenY(y)])',points)
        box=page.locator('#cv').bounding_box()
        page.mouse.move(box['x']+coords[0][0],box['y']+coords[0][1]);page.mouse.down()
        for x,y in coords[1:]+coords[:1]: page.mouse.move(box['x']+x,box['y']+y,steps=6)
        page.mouse.up()
        assert page.locator('#pk-crop-apply').is_enabled()
        page.click('#pk-crop-apply')
        page.wait_for_function('__mf.pick?.work_regions?.length===1')
        assert len(sess['pick'].world.segs)==1
        assert sess['pick'].world._work_regions[0]['fill_rule']=='enclosed'
        assert len(sess['pick'].world.circles)==1
        assert not page.locator('#status').inner_text().startswith('영역 선이')
        assert page.locator('#pk-crop-reset').is_enabled()
        page.screenshot(path=str(tmp_path/'pen-crop.png'))
        page.click('#pk-crop-reset')
        page.wait_for_function('!__mf.pick?.work_regions?.length')
        assert len(sess['pick'].world.segs)==2
        assert not errors,errors
        browser.close()


def test_crossed_pen_crop_persists_all_internal_entities_for_pipeline_replay(crop_client,tmp_path,monkeypatch):
    import hashlib
    from services.cad_import.pipeline import expand
    c,sess,path=crop_client
    original=hashlib.sha256(path.read_bytes()).hexdigest()
    world_before=deepcopy(sess['pick'].world.__dict__)
    points=[[1000,500],[8000,500],[9000,2000],[8000,3500],[1000,3500],[900,400],[1400,700]]
    # Older clients need not know the new policy; the crop API upgrades it.
    response=c.post('/api/module-f/pick/crop',json={'sid':sess['id'],'zone':{'type':'polygon','points':points}})
    assert response.status_code==200,response.json
    assert sess['job']['result']['ok'],sess['job']
    cropped=sess['pick'].world
    assert len(cropped.segs)==1 and len(cropped.circles)==1
    assert len(world_before['segs'])==2
    assert cropped._work_regions[0]['fill_rule']=='enclosed'
    spec_path=sess['pick'].commit(str(tmp_path/'spec'))
    monkeypatch.setattr(expand,'_spec_path',lambda key:spec_path)
    replay=expand.stage1_body(sess['key'])['w']
    assert replay.segs==cropped.segs and replay.circles==cropped.circles
    assert replay.texts==cropped.texts and replay._work_regions==cropped._work_regions
    assert hashlib.sha256(path.read_bytes()).hexdigest()==original
