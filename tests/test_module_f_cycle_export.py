"""Loop/grid export must preserve cycles without any tree-schedule sizing."""
from copy import deepcopy

import networkx as nx
import pytest

from test_module_f_flow_integration import flow_client
from routes.module_f import api_convert, api_design


def prepare(c, sess, mode='loop', default=65):
    sid = sess['id']
    c.post('/api/module-f/edit/network-mode',json={'sid':sid,'network_mode':mode})
    c.post('/api/module-f/edit/flow',json={'sid':sid})
    c.post('/api/module-f/edit/worst',json={'sid':sid,'k':1})
    result = c.post('/api/module-f/design/build',json={'sid':sid,'review_default_mm':default})
    assert result.status_code==200, result.json
    return sess['job']['result']


@pytest.mark.parametrize('mode',['loop','grid'])
def test_loop_tables_and_sdf_roundtrip(flow_client,tmp_path,mode):
    c,sess=flow_client
    original=deepcopy((sess['edit'].board.pts,sess['edit'].board.edges))
    result=prepare(c,sess,mode)
    assert result['ok'],result
    tbl=sess['design']['tables']
    g=nx.Graph((p['in'],p['out']) for p in tbl.pipes)
    assert nx.is_connected(g) and g.number_of_edges()-g.number_of_nodes()+1==1
    assert all(p['dia']==65 and p['dia_src']=='review_default' for p in tbl.pipes)
    assert all(p['bore_provenance']['head_count'] is None for p in tbl.pipes)
    assert (sess['edit'].board.pts,sess['edit'].board.edges)==original
    view=c.get('/api/module-f/design/preview',query_string={'sid':sess['id']})
    assert view.status_code==200 and view.json['view'],view.json
    assert not view.json['stale'],view.json['stale']
    emitted=c.post('/api/module-f/design/emit',json={'sid':sess['id']})
    assert emitted.status_code==200,emitted.json
    from services.cad_import.design.emit import _pc_models
    _pc_models()
    from pipenet_converter.sdf_parser import parse_sdf
    parsed=parse_sdf(sess['design_sdf_path'])
    assert len(parsed.pipes)==len(tbl.pipes)
    assert len(parsed.nozzles)==len(tbl.nozzles)
    restored=nx.Graph((p.from_node,p.to_node) for p in parsed.pipes.values())
    assert restored.number_of_edges()-restored.number_of_nodes()+1==1
    assert sum(p.length_m for p in parsed.pipes.values())==pytest.approx(sum(p['length'] for p in tbl.pipes))


def test_merged_review_keeps_cycles_and_unconfirmed_mark(flow_client,tmp_path):
    import json
    import zipfile
    from pathlib import Path
    from routes.module_f.merge import merge_network, bake_combined_iso
    from routes.module_f.emit import emit_merged
    from test_module_f_system_layout import system
    c,sess=flow_client
    assert prepare(c,sess)['ok']
    got=merge_network(sess['design']['tables'],riser=system(),mode='hsp_pump')
    files=emit_merged(got['combined'],tmp_path,iso_nodes=bake_combined_iso(got)[0])
    from pipenet_converter.sdf_parser import parse_sdf
    for key in ('sdf','sdf_iso'):
        path=Path(files[key])
        assert '검토용_미확정' in path.name
        assert 'NOT VALIDATED' in path.read_text(encoding='utf-8')
        net=parse_sdf(path)
        g=nx.MultiGraph((p.from_node,p.to_node) for p in net.pipes.values())
        assert nx.is_connected(g) and g.number_of_edges()-g.number_of_nodes()+1==1
    assert json.loads(Path(files['review']).read_text(encoding='utf-8'))['hydraulics_solved'] is False
    with zipfile.ZipFile(files['zip']) as zf:
        assert Path(files['review']).name in zf.namelist()


def test_zero_head_stem_rejected(flow_client):
    from services.cad_import.design.preserved import expand_preserved
    c,sess=flow_client
    assert prepare(c,sess)['ok']
    with pytest.raises(ValueError,match='헤드 접속관'):
        expand_preserved(sess['edit'].board,sess['worst']['heads'],
                         dto={'upright_1_m':0,'pendant_2_m':0,'combo_3_m':0,'combo_up_m':0})


def test_review_diameter_requires_explicit_selection(flow_client):
    c,sess=flow_client
    result=prepare(c,sess,default=None)
    assert not result['ok'] and '기본 관경' in result['error']
    assert not sess.get('design')


def test_all_conversion_outputs_and_review_record(flow_client,tmp_path,monkeypatch):
    import json
    from pathlib import Path
    c,sess=flow_client
    api_convert.register(c.application,UPLOAD_DIR=tmp_path)
    monkeypatch.setattr(api_convert,'_run_job',api_design._run_job)
    from services.cad_import.design import bore
    monkeypatch.setattr(bore,'decide_bores',lambda *a,**k:pytest.fail('Tree sizing called'))
    assert prepare(c,sess)['ok']
    response=c.post('/api/module-f/convert/run',json={'sid':sess['id'],
        'outputs':{'full_kfp':True,'worst_kfp':True,'worst_sdf':True}})
    assert response.status_code==200,response.json
    result=sess['job']['result']; assert result['ok'],result
    audit=json.loads(Path(sess['design_review_path']).read_text(encoding='utf-8'))
    assert not audit['hydraulics_solved'] and audit['cycle_rank']==1
    assert all(p['bore_provenance']['review_only'] for p in audit['tables']['pipes'])
    for key in ('kfp_path','worst_kfp_path'):
        kfp=json.loads(Path(sess[key]).read_text(encoding='utf-8'))
        graph=nx.Graph((p['start'],p['end']) for p in kfp['pipe_data'].values())
        assert graph.number_of_edges()-graph.number_of_nodes()+1==1
        assert kfp['review_only'] and not kfp['hydraulics_solved']
        assert all(p['nominal_mm']==65 and p['diameter']>0 for p in kfp['pipe_data'].values())
    assert '검토용' in Path(sess['design_sdf_path']).name
    assert c.get('/api/module-f/download',query_string={'sid':sess['id'],'what':'design-review'}).status_code==200
    sess['edit'].board.pts[1]=(1234,0)
    assert c.get('/api/module-f/download',query_string={'sid':sess['id'],'what':'design'}).status_code==409


def test_grid_with_multiple_cycles_and_combo_lengths(flow_client):
    from services.cad_import.design.preserved import expand_preserved
    from services.cad_import.design.flow import network_for_board
    from services.cad_import.edit.board import EditBoard
    points=[(x*1000,y*1000) for y in range(3) for x in range(3)]
    edges={(i,i+1) for i in range(9) if i%3<2}|{(i,i+3) for i in range(6)}
    board=EditBoard('grid',points,edges,[(2000,2000,30)])
    board.sources=[0];board.network_mode='grid';board.set_head_kind(board.disks[0],'상하향식')
    got=expand_preserved(board,[0],datum_m=4,dto={'combo_1_m':0.3,'combo_2_m':0.7,'combo_up_m':0.2,'combo_3_m':0.5})
    assert got['cycle_rank']==4
    pipes=got['kfp']['pipe_data']
    added=[p['length_m'] for pid,p in pipes.items() if pid not in got['edge_ref']]
    assert sorted(added)==pytest.approx([.2,.3,.5,.7])
    assert len(got['node_head_kinds'])==2
    assert len(network_for_board(board).edges)==12


def test_browser_review_controls_and_unconfirmed_colours(flow_client,tmp_path,monkeypatch):
    from urllib.parse import urlparse
    from test_module_f_viewport_browser import install_page,pw_api,ROOT
    c,sess=flow_client
    api_convert.register(c.application,UPLOAD_DIR=tmp_path)
    assert prepare(c,sess)['ok']
    source=(ROOT/'static/module_f.js').read_text(encoding='utf-8')
    source=source.replace('  setStage("open");\n  loadSaved();',
        '  window.__reviewTest={setEdit,designPreview,syncDesignForMethod};\n  setStage("open");\n  loadSaved();')
    with pw_api.sync_playwright() as pw:
        browser=pw.chromium.launch();page=browser.new_page(viewport={'width':1440,'height':1000})
        errors=install_page(page,source)
        def route(req):
            url=urlparse(req.request.url)
            res=c.open(url.path+('?' + url.query if url.query else ''),method=req.request.method,
                       data=req.request.post_data,content_type='application/json')
            if res.status_code==404:req.fulfill(json={'ok':True,'items':[],'rows':[],'fields':{}})
            else:req.fulfill(status=res.status_code,body=res.data,content_type=res.content_type)
        page.route('http://module-f.test/api/module-f/**',route)
        state=c.get('/api/module-f/edit/state',query_string={'sid':sess['id']}).json['state']
        page.evaluate('''async ({sid,state})=>{Object.assign(__mf,{sid,slot:'plan',method:'manual'});
            __reviewTest.setEdit(state);__viewTest.setStage('design');__reviewTest.syncDesignForMethod();
            await __reviewTest.designPreview();}''',{'sid':sess['id'],'state':state})
        assert page.locator('#dg-review-inputs').is_visible()
        assert page.locator('#dg-review-dia').input_value()=='65'
        page.get_by_text('검토용 출력 · 미확정 유지',exact=True).click()
        assert '규약 관경은 적용하지 않습니다' in page.locator('#dg-review-notice').inner_text()
        assert page.locator('#dg-diagnose').is_disabled()
        colours=page.evaluate('()=>__mf.design.view.pipes.map(p=>ModuleFBore.info(p).category)')
        assert colours and set(colours)=={'review'}
        assert '규약 미적용' in page.locator('#bore-legend').inner_text()
        page.screenshot(path=str(tmp_path/'cycle-review.png'))
        assert not errors,errors
        browser.close()


def test_physical_tee_not_changed_to_elbow_by_scenario_pruning():
    from services.cad_import.design.preserved import review_fittings
    net={'nodes_meta_runtime':{'a':{'coords':[-1,0,0]},'j':{'coords':[0,0,0]},'b':{'coords':[0,1,0]}},
         'pipe_data':{'A':{'start':'a','end':'j'},'B':{'start':'j','end':'b'}}}
    fits=review_fittings(net,{'A':(65,'review_default'),'B':(65,'review_default')},
                        {'j':[[-1,0,0],[1,0,0],[0,1,0]]})
    assert fits['counts']=={'tee':1}
    assert fits['per_pipe']['B']['fittings']==['tee']


def test_real_b1f_review_sdf_preserves_loop_grid(flow_client,tmp_path,monkeypatch):
    """Opt-in fresh DXF extraction; saved edits/cache/DXF remain byte-identical."""
    import os,json,hashlib,shutil
    from pathlib import Path
    if os.environ.get('MODULE_F_REAL_REVIEW_CHECK')!='1': pytest.skip('Actual DXF opt-in')
    from services.cad_import.pipeline.expand import stage1_body
    from services.cad_import.pipeline.flow import pipeline,ho_from_spots
    from services.cad_import.edit.io import _board_from_data,load_edits
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.flow import network_for_board
    from routes.module_f.preserved_design import build_review_design
    from routes.module_f.api_design import _settings,emit_design_files
    from services.cad_import.design.emit import _pc_models
    key='B1F 현장조사 소화설비 평면도_컨셉2-수정본 (1)'
    folder=Path('cad_project_editor_g/docs/import')
    saved=folder/'DWG'/(key+'_유저손질.json')
    fingerprint=hashlib.sha256(saved.read_bytes()).hexdigest()
    result=pipeline(stage1_body(key)); result['ho']=ho_from_spots(result['spots'])
    board=_board_from_data(key,result); assert load_edits(board,str(folder/'DWG'))
    c,sess=flow_client
    sess['edit']=EditSession(board,key=key,out_dir=str(tmp_path))
    sess['key']='b1f_review'
    destination=Path(os.environ.get('MODULE_F_REVIEW_OUTPUT',str(tmp_path)))
    destination.mkdir(parents=True,exist_ok=True)
    reports=[]
    for mode in ('loop','grid'):
        board.network_mode=mode
        network=network_for_board(board)
        # Use all attached heads as a compatibility scenario, not a certified worst area.
        heads=sorted(set(network.reference.representatives.values()))
        sess['worst']={'heads':heads,'edges':network.selected_edges(heads),
            'loads':{},'physical_loads':{},'source_tag':'Z1','flow_revision':network.revision,'network_mode':mode}
        cfg=_settings(sess,{'review_default_mm':65,'review_datum_m':0,'iso':True})
        built=build_review_design(sess,sess['edit'],cfg,'Z1')
        assert built['ok'],built
        path,error=emit_design_files(sess,tmp_path)
        assert not error,error
        _pc_models()
        from pipenet_converter.sdf_parser import parse_sdf
        parsed=parse_sdf(path)
        graph=nx.Graph((p.from_node,p.to_node) for p in parsed.pipes.values())
        rank=graph.number_of_edges()-graph.number_of_nodes()+1
        assert nx.is_connected(graph) and rank==network.report()['cycle_rank']
        tbl=sess['design']['tables']
        assert len(parsed.pipes)==len(tbl.pipes)
        assert len(parsed.nozzles)==len(tbl.nozzles)
        assert all(p.diameter_m>0 and p.length_m>0 and p.length_m>=abs(p.rise_m)-1e-6 for p in parsed.pipes.values())
        length=sum(p.length_m for p in parsed.pipes.values())
        assert length==pytest.approx(sum(p['length'] for p in tbl.pipes))
        for suffix in ('.sdf','.slf','.review.json'):
            target=destination/mode/path.with_suffix(suffix).name
            target.parent.mkdir(exist_ok=True)
            shutil.copy2(path.with_suffix(suffix),target)
        reports.append({'mode':mode,'cycle_rank':rank,'pipes':len(parsed.pipes),'nozzles':len(parsed.nozzles),
                        'length_m':length,'default_nominal_mm':65,'source_heads':len(heads),
                        'hydraulics_solved':False,'fittings_unresolved':tbl.unresolved})
    (destination/'verification.json').write_text(json.dumps(reports,ensure_ascii=False,indent=2),encoding='utf-8')
    assert hashlib.sha256(saved.read_bytes()).hexdigest()==fingerprint
    print('REAL_REVIEW_OUTPUT',destination)
