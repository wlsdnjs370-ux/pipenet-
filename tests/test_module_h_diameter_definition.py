"""H drawing priority and bounded structural correspondence, independent of Flask."""
import pytest
from src.pipenet_converter.graph.diameter_definition import (
    define_diameters, resolve_pipe_definitions, pattern_preview, DefinitionConfig,
)
from src.pipenet_converter.graph.diameter_inference import DiameterAnnotation as Text
from src.pipenet_converter.graph.flow import build_flow_tree
from src.pipenet_converter.validate.diameter_evidence import require_resolved_diameters
from test_module_f_flow_integration import flow_client


@pytest.fixture(autouse=True)
def isolated_edit_histories(tmp_path,monkeypatch):
    from routes.module_f import network_edit
    monkeypatch.setattr(network_edit,'HISTORY_DIR',tmp_path/'histories')


def context(points, edges, texts=(), heads=None, barriers=()):
    flow = build_flow_tree(points, edges, [[len(points)-1]], [0])
    return define_diameters(points, edges, flow, texts, heads=heads, barriers=barriers)


def test_100_continues_through_main_tee_and_bend_not_side_branch():
    pts=[(0,0),(5000,0),(10000,0),(10000,5000),(5000,5000)]
    ctx=context(pts,[(0,1),(1,2),(2,3),(1,4)],[Text('100',2000,100,100,0)])
    assert ctx.decisions[2,3]['text_mm']==100
    assert ctx.decisions[2,3]['definition_source']=='drawing_run'
    assert not ctx.decisions[1,4].get('text_mm')


def test_valve_is_a_propagation_barrier_and_conflicts_are_not_overwritten():
    pts=[(0,0),(5000,0),(10000,0),(15000,0)]
    links=[(0,1),(1,2),(2,3)]
    ctx=context(pts,links,[Text('a',1000,100,65,0)],barriers=[1])
    assert not ctx.decisions[1,2].get('text_mm')
    ctx=context(pts,links,[Text('a',1000,100,65,0),Text('b',14000,100,40,0)])
    assert ctx.decisions[1,2]['block_export']


def test_repeat_uses_head_landmarks_not_cad_fragment_count():
    pts=[(0,0),(5000,0),(10000,0),(15000,0),(5000,4000),(5000,8000),
         (10000,2000),(10000,4000),(10000,8000)]
    links=[(0,1),(1,2),(2,3),(1,4),(4,5),(2,6),(6,7),(7,8)]
    ctx=context(pts,links,[Text('40',5100,2000,40,90),Text('25',5100,6000,25,90)],
                heads={4:'up',5:'up',7:'up',8:'up'})
    assert ctx.decisions[2,6]['text_mm']==40
    assert ctx.decisions[6,7]['text_mm']==40
    assert ctx.decisions[7,8]['text_mm']==25
    assert ctx.decisions[7,8]['definition_source']=='drawing_repeat'
    assert ctx.decisions[7,8]['representative_nodes']==[1,4,5]


def test_tree_drawing_below_rule_is_kept_loop_unknown_not_guessed_manual_last():
    pts=[(0,0),(5000,0)]
    ctx=context(pts,[(0,1)],[Text('25',1000,100,25,0)])
    net={'pipe_data':{'P1':{}}}
    vals,ev,_=resolve_pipe_definitions(net,{'P1':(0,1)},ctx,{'P1':15},tree=True,rule=lambda n:65)
    assert vals['P1']==(25,'text') and ev['P1']['rule_mm']==65
    ctx=context(pts,[(0,1)])
    vals,ev,_=resolve_pipe_definitions(net,{'P1':(0,1)},ctx,{},tree=False,rule=None)
    assert vals['P1']==(0,'unresolved')
    with pytest.raises(ValueError,match='미지정'):
        require_resolved_diameters([{'label':'P1','dia':0,'bore_provenance':ev['P1']}])
    vals,ev,_=resolve_pipe_definitions(net,{'P1':(0,1)},ctx,{},tree=False,rule=None,overrides={(0,1):(40,'도면 확인')})
    assert vals['P1']==(40,'user') and ev['P1']['manual']['previous_mm'] is None
    require_resolved_diameters([{'label':'P1','dia':40,'bore_provenance':ev['P1']}])


def test_copy_different_shape_requires_choice_at_boundary_and_keeps_geometry():
    points=[(0,0),(0,4),(3,4)]
    source=[{'length_mm':1,'nominal_mm':65},{'length_mm':1,'nominal_mm':25}]
    result=pattern_preview(points,source,[(0,1),(1,2)])
    assert result[0]['needs_choice'] and result[0]['candidates']==[25,65]
    assert result[1]['nominal_mm']==25
    assert pattern_preview(points,source,[(0,1),(1,2)],reverse=True)[1]['nominal_mm']==65
    assert points==[(0,0),(0,4),(3,4)]
    with pytest.raises(ValueError):
        pattern_preview(points,source,[(0,1),(1,2),(0,2)])


@pytest.mark.parametrize('kw',[{'repeat_length_tolerance':float('nan')},{'max_branch_edges':0}])
def test_invalid_config_fails(kw):
    with pytest.raises(ValueError): DefinitionConfig(**kw)


@pytest.mark.parametrize('mode',['tree','loop','grid'])
def test_h_build_without_default_and_keeps_all_heads(flow_client,mode):
    c,sess=flow_client
    sess['edit'].board.network_mode=mode
    body={'sid':sess['id'],'selection_mode':'area_all','zones':[[1900,1900,3100,2100]]}
    assert c.post('/api/module-f/edit/worst',json=body).status_code==200
    assert c.post('/api/module-f/design/build',json={**body,'diameter_policy':'drawing_first_v1'}).status_code==200
    assert sess['job']['result']['ok'],sess['job']['result']
    assert len(sess['design']['tables'].nozzles)==2
    assert 'review_default_mm' not in sess['design_settings']
    if mode!='tree':
        assert any(p['dia']==0 for p in sess['design']['tables'].pipes)


@pytest.mark.parametrize('mode',['tree','loop'])
def test_h_batch_saved_stable_keys_rebuild_and_undo(flow_client,monkeypatch,tmp_path,mode):
    from routes.module_h_attributes import register
    from routes.module_f import overrides
    c,sess=flow_client
    register(c.application)
    monkeypatch.setattr(overrides,'path_for',lambda key:str(tmp_path/'overrides.json'))
    sess['edit'].board.network_mode=mode
    body={'sid':sess['id'],'selection_mode':'area_all','zones':[[1900,1900,3100,2100]],'diameter_policy':'drawing_first_v1'}
    c.post('/api/module-f/edit/worst',json=body)
    c.post('/api/module-f/design/build',json=body)
    assert sess['job']['result']['ok'],sess['job']['result']
    status=c.get('/api/module-f/h-attributes/state',query_string={'sid':sess['id']})
    assert status.status_code==200,status.json
    d=status.json
    pipes=[r for r in d['records'] if r['kind']=='pipe']
    assert pipes
    result=c.post('/api/module-f/h-attributes/apply',json={'sid':sess['id'],'revision':d['revision'],
        'ids':[r['id'] for r in pipes],'nominal_mm':65})
    assert result.status_code==200,result.json
    assert (tmp_path/'overrides.json').is_file()
    c.post('/api/module-f/design/build',json=body)
    assert sess['job']['result']['ok'],sess['job']['result']
    assert all(p['dia']==65 for p in sess['design']['tables'].pipes)
    status=c.get('/api/module-f/h-attributes/state',query_string={'sid':sess['id']}).json
    undo=c.post('/api/module-f/h-attributes/undo',json={'sid':sess['id'],'revision':status['revision']})
    assert undo.status_code==200,undo.json
    stale=c.post('/api/module-f/h-attributes/apply',json={'sid':sess['id'],'revision':d['revision'],'ids':[pipes[0]['id']],'nominal_mm':40})
    assert stale.status_code==409


def test_h_repeat_conflicting_donors_do_not_choose_nearest():
    points=[(i*5000,0) for i in range(5)]+[(5000,4000),(10000,4000),(15000,4000)]
    edges=[(0,1),(1,2),(2,3),(3,4),(1,5),(2,6),(3,7)]
    ctx=context(points,edges,[Text('40',5100,2000,40,90),Text('65',15100,2000,65,90)],heads={5:'up',6:'up',7:'up'})
    assert not ctx.decisions[2,6].get('text_mm')
    assert ctx.decisions[2,6]['block_export']


@pytest.mark.parametrize('mode',['tree','loop'])
def test_bulk_head_properties_keep_geometry_and_reach_nozzle_library(flow_client,monkeypatch,tmp_path,mode):
    from copy import deepcopy
    from routes.module_h_attributes import register
    from routes.module_f import overrides
    c,sess=flow_client;register(c.application)
    monkeypatch.setattr(overrides,'path_for',lambda key:str(tmp_path/'head-overrides.json'))
    sess['edit'].board.network_mode=mode
    body={'sid':sess['id'],'selection_mode':'area_all','zones':[[1900,1900,3100,2100]],'diameter_policy':'drawing_first_v1'}
    c.post('/api/module-f/edit/worst',json=body);c.post('/api/module-f/design/build',json=body)
    d=c.get('/api/module-f/h-attributes/state',query_string={'sid':sess['id']}).json
    before=deepcopy((sess['edit'].board.pts,sess['edit'].board.edges))
    result=c.post('/api/module-f/h-attributes/apply',json={'sid':sess['id'],'revision':d['revision'],
        'ids':[r['id'] for r in d['records'] if r['kind']=='head'],
        'head_properties':{'k_factor_si':90,'required_pressure_bar':1.5}})
    assert result.status_code==200,result.json
    c.post('/api/module-f/design/build',json=body)
    assert sess['job']['result']['ok'],sess['job']['result']
    metadata=sess['design']['got']['kfp']['nodes_meta_runtime'].values()
    assert sum(r.get('k_factor_si')==90 for r in metadata)==2
    assert all(r.get('k_factor_si')==90 and r.get('required_pressure_bar')==1.5
               for r in sess['design']['tables'].nozzles)
    assert before==(sess['edit'].board.pts,sess['edit'].board.edges)
    current=c.get('/api/module-f/h-attributes/state',query_string={'sid':sess['id']}).json
    undo=c.post('/api/module-f/h-attributes/undo',json={'sid':sess['id'],'revision':current['revision']})
    assert undo.status_code==200,undo.json


def test_fresh_state_does_not_allow_overwriting_another_windows_file(flow_client,monkeypatch,tmp_path):
    import json
    from routes.module_h_attributes import register
    from routes.module_f import overrides
    c,sess=flow_client;register(c.application)
    path=tmp_path/'shared.json';monkeypatch.setattr(overrides,'path_for',lambda key:str(path))
    body={'sid':sess['id'],'selection_mode':'area_all','zones':[[1900,1900,3100,2100]],'diameter_policy':'drawing_first_v1'}
    c.post('/api/module-f/edit/worst',json=body);c.post('/api/module-f/design/build',json=body)
    d=c.get('/api/module-f/h-attributes/state',query_string={'sid':sess['id']}).json
    pipe=next(r for r in d['records'] if r['kind']=='pipe')
    path.write_text(json.dumps({'items':[{'key':pipe['key'],'field':'dia','new':100}]}))
    saved=path.read_bytes()
    fresh=c.get('/api/module-f/h-attributes/state',query_string={'sid':sess['id']}).json
    result=c.post('/api/module-f/h-attributes/apply',json={'sid':sess['id'],'revision':fresh['revision'],'ids':[pipe['id']],'nominal_mm':80})
    assert result.status_code==409 and '다른 창' in result.json['message']
    assert path.read_bytes()==saved


def test_compacted_pipe_bulk_edit_covers_all_original_stable_keys(flow_client,monkeypatch,tmp_path):
    from test_module_f_export_compaction import chain
    from src.pipenet_converter.graph.continuous_runs import compact_continuous
    from routes.module_h_attributes import register
    from routes.module_f import overrides
    c,sess=flow_client;register(c.application)
    monkeypatch.setattr(overrides,'path_for',lambda key:str(tmp_path/'compact.json'))
    tbl=chain()
    for i,p in enumerate(tbl.pipes):
        p['bore_provenance']={'policy':'drawing_first_v1','definition_source':'drawing_run','text_mm':65,'stable_key':['pipe',i,i+1]}
    tbl,_=compact_continuous(tbl);tbl.pipe_labels={}
    board=sess['edit'].board
    board.pts=[(i*1000,0) for i in range(6)];board.edges={(i,i+1) for i in range(5)}
    sess['design_settings']={'diameter_policy':'drawing_first_v1'}
    sess['_h_diameter_context']=context(board.pts,board.edges,[Text('a',100,100,65,0)])
    sess['design']={'tables':tbl,'got':{}}
    monkeypatch.setattr(overrides,'label_keys',lambda *a:{'pipe':{},'node':{}})
    d=c.get('/api/module-f/h-attributes/state',query_string={'sid':sess['id']}).json
    run=next(r for r in d['records'] if r['key'][0]=='run')
    assert len(run['keys'])==4 and len(run['edges'])==4
    result=c.post('/api/module-f/h-attributes/apply',json={'sid':sess['id'],'revision':d['revision'],'ids':[run['id']],'nominal_mm':80})
    assert result.status_code==200,result.json
    assert {tuple(r['key']) for r in overrides.load(sess)}=={('pipe',i,i+1) for i in range(4)}


def test_h_missing_text_cache_rereads_source_but_f_is_unchanged(tmp_path,monkeypatch):
    import json
    from src.pipenet_converter.dxf import diameter_annotations as native
    from routes.module_f.api_design import _dia_texts
    from services.cad_import.pipeline import handoff,stage1
    source=tmp_path/'source.dxf';source.write_text('fixture')
    (tmp_path/'drawing_찍은스펙.json').write_text(json.dumps({'source_dxf':str(source)}),encoding='utf-8')
    monkeypatch.setattr(handoff,'pick_out_dir',lambda:str(tmp_path))
    monkeypatch.setattr(handoff,'load_world',lambda *a:None)
    reads=[]
    def read(path, *, extended=False):
        assert extended
        reads.append(path)
        return [native.NativeText(12,34,2,'100A',90,'mapped')]
    monkeypatch.setattr(native,'read_native_texts',read)
    sess={'key':'drawing','design_settings':{'diameter_policy':'drawing_first_v1'}}
    assert _dia_texts(sess)==[(12.,34.,100)]
    assert _dia_texts(sess)==[(12.,34.,100)] and len(reads)==1
    assert _dia_texts({'key':'drawing'})==[] and len(reads)==1
    sess.pop('_h_diameter_text_points')
    def fail(*args): raise OSError('read failed')
    monkeypatch.setattr(native,'read_native_texts',fail)
    with pytest.raises(ValueError,match='원본 DXF'):
        _dia_texts(sess)
