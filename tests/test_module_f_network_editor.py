"""Graph edits, authoritative libraries, replay, cross-view and export tests."""
from copy import deepcopy
import json

import pytest
from flask import Flask

from core.network_editor import EditError, Network, apply_edit
from routes.module_f import network_edit as ne, api_network_edit, jobs


def table():
    ne._boot()
    from services.cad_import.design.tables import PipeTablesG
    # Source -> junction -> head; off-grid XY must not be rounded on extension.
    return PipeTablesG(nodes=[dict(label=str(i+1),x=x*1000,y=37.1540934468307*1000,
        elevation=0,io_node='Input' if i==0 else 'No') for i,x in enumerate((90.1,91.16033980168356,93.1))],
        pipes=[dict(label='P1',**{'in':'1','out':'2'},length=1.06033980168356,dia=25,type='KSD 3507',c=120,eq_len=0),
               dict(label='P2',**{'in':'2','out':'3'},length=1.93966019831644,dia=25,type='KSD 3507',c=120,eq_len=0)],
        nozzles=[dict(label='H1',**{'in':'3','out':'@/3'},flow_lmin=80,flow_m3s=80/60000,lib='SP-HEAD',status='1')],
        equipment=[dict(label='AV',pipe='P1',**{'in':'1','out':'2'},rel_pos=.8,eq_len=1.5,desc='A/V')],
        pipe_labels={'KP1':'P1','KP2':'P2'})


def design():
    t=table()
    nodes={f'N{n["label"]}':dict(coords=[n['x']/1000,n['y']/1000,0],elevation_m=0,
             type_id='head' if n['label']=='3' else 'base') for n in t.nodes}
    return dict(tables=t,method='manual',got={'kfp':{'nodes_meta_runtime':nodes,
        'pipe_data':{f'KP{i+1}':dict(start='N'+p['in'],end='N'+p['out'],length_m=p['length'],
             nominal_mm=25,diameter=27.5,C=120,equivalent_length=0) for i,p in enumerate(t.pipes)}}},
        keys={'node':{},'pipe':{},'nid':{str(i):f'N{i}' for i in (1,2,3)}})


def cmd(op='extend',target='2',**kw):
    return dict(op=op,target=target,axis='+Y',length=2,schedule='KSD 3507',dn=25,**kw)


@pytest.fixture()
def session(tmp_path,monkeypatch):
    monkeypatch.setattr(ne,'HISTORY_DIR',tmp_path/'history')
    sess=jobs._new_session(key='editor-fixture')
    sess['design']=design()
    yield sess
    jobs._SESSIONS.pop(sess['id'],None)


@pytest.fixture()
def client(session):
    app=Flask(__name__)
    app.testing=True
    api_network_edit.register(app)
    return app.test_client()


def state(client,sess,scope='design'):
    r=client.get('/api/module-f/network-editor',query_string=dict(sid=sess['id'],scope=scope))
    assert r.status_code==200,r.json
    return r.json


def post(client,sess,command=None,action='apply',scope='design',**kw):
    return client.post('/api/module-f/network-editor',json=dict(sid=sess['id'],scope=scope,
          revision=state(client,sess,scope)['revision'],action=action,command=command or {},**kw))


@pytest.mark.parametrize('direction',['+X','-X','+Y','-Y','+Z','-Z'])
def test_extension_preserves_perpendicular_coordinates(direction):
    t=table(); t.pipes=[];t.nozzles=[];t.nodes=t.nodes[1:2]
    n=Network.from_tables(t)
    c=cmd();c['axis']=direction
    changed,r=apply_edit(n,c,ne.catalog())
    orig=n.nodes['2'].xyz;added=changed.nodes[r['label']].xyz
    i='XYZ'.index(direction[1])
    assert all(added[j]==orig[j] for j in range(3) if j!=i)
    assert abs(added[i]-orig[i])==pytest.approx(2)
    assert len(n.nodes)==1  # preview is pure


def test_split_preserves_total_length_and_moves_equipment_by_position():
    n=Network.from_tables(table())
    out,result=apply_edit(n,cmd('split','P1',distance=.5),ne.catalog())
    assert sum(p.row['length'] for p in n.pipes.values())==pytest.approx(sum(p.row['length'] for p in out.pipes.values()))
    av=out.tables.equipment[0]
    assert av['pipe']=='EP1' and 0<av['rel_pos']<1
    assert av['eq_len']==1.5
    assert len(out.incident(result['label']))==2


def test_split_then_branch_and_add_head():
    n,r=apply_edit(Network.from_tables(table()),cmd('split','P2',distance=.6),ne.catalog())
    n,r=apply_edit(n,cmd(target=r['label']),ne.catalog())
    n,_=apply_edit(n,cmd('head',r['label'],nozzle='SP-HEAD',flow=120),ne.catalog())
    assert len(n.tables.nozzles)==2


def test_resize_moves_downstream_only_and_preserves_other_lengths():
    n=Network.from_tables(table());c=cmd('resize','P1');c['length']=2
    out,_=apply_edit(n,c,ne.catalog())
    assert out.nodes['1'].xyz==n.nodes['1'].xyz
    assert out.nodes['3'].xyz[0]-n.nodes['3'].xyz[0]==pytest.approx(2-n.pipes['P1'].row['length'])
    assert out.pipes['P2'].row['length']==n.pipes['P2'].row['length']


@pytest.mark.parametrize('edit',[
    dict(op='extend',target='3',axis='+Z',length=1,schedule='KSD 3507',dn=25),
    dict(op='extend',target='2',axis='+X',length=1,schedule='KSD 3507',dn=25),
    dict(op='connect',target='1',end='3',schedule='KSD 3507',dn=25),
    dict(op='pipe',target='P1',schedule='UNKNOWN',dn=25),
    dict(op='pipe',target='P1',schedule='CPVC2',dn=200),
    dict(op='delete',target='P1'),
    dict(op='split',target='P1',distance=0),
    dict(op='resize',target='P1',length='NaN'),
])
def test_invalid_edits_rejected_without_changes(edit):
    n=Network.from_tables(table());before=ne.fingerprint(n)
    with pytest.raises(EditError):apply_edit(n,edit,ne.catalog())
    assert ne.fingerprint(n)==before


def test_same_height_crossing_rejected_but_different_elevation_allowed():
    t=table()
    t.nodes += [dict(label='4',x=90100,y=38154.0934468307,elevation=0),
                dict(label='5',x=92100,y=38154.0934468307,elevation=0)]
    t.pipes += [dict(label='P3',**{'in':'4','out':'5'},length=2,dia=25)]
    with pytest.raises(EditError,match='교차'):
        apply_edit(Network.from_tables(t),cmd(),ne.catalog())
    t.nodes[-1]['elevation']=t.nodes[-2]['elevation']=2
    out,_=apply_edit(Network.from_tables(t),cmd(),ne.catalog())
    assert len(out.pipes)==4


def test_library_inner_and_missing_fitting_are_not_guessed():
    n=Network.from_tables(table());c=cmd('pipe','P2');c['schedule']='CPVC2'
    out,_=apply_edit(n,c,ne.catalog())
    assert out.pipes['P2'].row['inner_mm']==28.02
    assert out.pipes['P2'].row['c']==150
    out.pipes['P2'].row['dia']=15
    with pytest.raises(EditError,match='등가길이'):
        apply_edit(out,cmd('fitting','P2',fitting='ELBOW_45',count=1),ne.catalog())


def test_preview_commit_undo_redo_and_reopen(client,session):
    before=deepcopy(session['design']['tables'].as_dict())
    r=post(client,session,cmd(),action='preview');assert r.status_code==200,r.json
    assert session['design']['tables'].as_dict()==before
    r=post(client,session,cmd());assert r.status_code==200,r.json
    assert len(session['design']['tables'].pipes)==3
    assert len(session['design']['got']['kfp']['pipe_data'])==3
    assert ne._path(session,'design').is_file()
    assert post(client,session,action='undo').status_code==200
    assert len(session['design']['tables'].pipes)==2
    assert post(client,session,action='redo').status_code==200
    session.pop('network_editor');session['design']=design()
    assert state(client,session)['undo']==1
    assert len(session['design']['got']['kfp']['pipe_data'])==3


def test_no_stale_label_replay_and_reset(client,session):
    assert post(client,session,cmd()).status_code==200
    session['design']=design();session['design']['tables'].pipes[0]['length']=333
    d=state(client,session)
    assert d['conflict']
    assert post(client,session,cmd()).status_code==409
    assert len(session['design']['tables'].pipes)==2
    assert post(client,session,action='reset').status_code==200
    assert not state(client,session)['conflict']
    assert session['design']['tables'].pipes[0]['length']==333


def test_stale_revision_rejected(client,session):
    rev=state(client,session)['revision']
    assert post(client,session,cmd()).status_code==200
    r=client.post('/api/module-f/network-editor',json=dict(sid=session['id'],revision=rev,action='apply',command=cmd()))
    assert r.status_code==409
    assert len(session['design']['tables'].pipes)==3


def test_head_change_undo_restores_kfp_metadata(client,session):
    assert post(client,session,cmd('node','3')).status_code==200
    assert session['design']['got']['kfp']['nodes_meta_runtime']['N3']['type_id']=='base'
    assert post(client,session,action='undo').status_code==200
    assert session['design']['got']['kfp']['nodes_meta_runtime']['N3']['type_id']=='head'


def test_delete_undo_restores_original_pipe_identity_and_nozzle(client,session):
    assert post(client,session,cmd('delete','P2')).status_code==200
    assert post(client,session,action='undo').status_code==200
    assert 'KP2' in session['design']['got']['kfp']['pipe_data']
    assert session['design']['tables'].nozzles[0]['label']=='H1'


def test_split_kfp_has_unique_node_identity_and_two_nonzero_pipes(client,session):
    r=post(client,session,cmd('split','P2',distance=.6))
    assert r.status_code==200,r.json
    d=session['design'];kf=d['got']['kfp']
    assert len(kf['nodes_meta_runtime'])==len(d['tables'].nodes)==4
    assert len(set(d['keys']['nid'].values()))==4
    for p in kf['pipe_data'].values():
        assert p['start']!=p['end']
        assert kf['nodes_meta_runtime'][p['start']]['coords']!=kf['nodes_meta_runtime'][p['end']]['coords']
    from core.kfp_editability import prepare_kfp
    assert not prepare_kfp(kf)[1].errors


def test_fitting_and_bore_export_match_tables(client,session,tmp_path):
    c=cmd('pipe','P2');c['schedule']='CPVC2'
    assert post(client,session,c).status_code==200
    assert post(client,session,cmd('fitting','P2',fitting='ELBOW_45',count=2)).status_code==200
    from services.cad_import.design.emit import emit_design_kfp,tables_to_network,emit_design_sdf
    d=session['design'];t=d['tables']
    kfp=emit_design_kfp(t,d['got'],None)['kfp']
    p=kfp['pipe_data']['KP2']
    assert p['diameter']==28.02 and p['C']==150 and p['equivalent_length']==pytest.approx(.6)
    sdf_net=tables_to_network(t,project_title='editor')
    assert sdf_net.pipes['P2'].equipment[0].equivalent_length_m==pytest.approx(.6)
    emit_design_sdf(t,tmp_path/'edited.sdf')
    import xml.etree.ElementTree as ET
    root=ET.parse(tmp_path/'edited.sdf').getroot()
    assert 'CPVC2' in (tmp_path/'edited.sdf').read_text(encoding='utf-8')
    assert root is not None


def test_plan_edit_in_merge_updates_both_views(client,session):
    from test_module_f_merge import _riser
    from routes.module_f.api_merge import rebuild_merged
    session['slots']['system']['riser']=_riser()
    session['supply_mode']='lsp_gravity'
    rebuild_merged(session)
    r=post(client,session,cmd(target='11'),scope='merge')
    assert r.status_code==200,r.json
    assert len(session['design']['tables'].pipes)==3
    assert any(p['label']=='EP1' for p in session['merged']['combined'].pipes)
    assert r.json['selection']['label']=='13'
    assert post(client,session,action='undo',scope='merge').status_code==200
    assert len(session['design']['tables'].pipes)==2


def test_merged_system_split_keeps_part_and_undo(client,session):
    from test_module_f_merge import _riser
    from routes.module_f.api_merge import rebuild_merged
    session['slots']['system']['riser']=_riser()
    session['supply_mode']='lsp_gravity';rebuild_merged(session)
    p=next(p for p in session['merged']['combined'].pipes if str(p['label']).startswith('r'))
    r=post(client,session,cmd('split',p['label'],distance=float(p['length'])/2),scope='merge')
    assert r.status_code==200,r.json
    added=r.json['selection']['label']
    assert added in session['merged']['parts']['system']
    assert post(client,session,action='undo',scope='merge').status_code==200
    assert added not in {str(n['label']) for n in session['merged']['combined'].nodes}
    assert post(client,session,action='redo',scope='merge').status_code==200
    assert added in {str(n['label']) for n in session['merged']['combined'].nodes}


def test_late_cancel_restores_both_history_file_and_memory(session):
    from routes.module_f.cancellation import Operation,OperationCancelled
    editor=ne.ensure(session,'design')
    candidate,_=apply_edit(editor['current'],cmd(),ne.catalog())
    operation=Operation('network-editor-late-cancel')
    operation.snapshot(session)
    # Function name protects the final commit, just like the real writer boundary.
    def commit_and_cancel():
        api_network_edit.commit_edit(session,'design',editor,candidate,[cmd()],1)
        operation.cancel()
    with pytest.raises(OperationCancelled):
        with operation.scope():
            commit_and_cancel()
    assert len(session['design']['tables'].pipes)==2
    assert not ne._path(session,'design').exists()


def test_another_window_cannot_overwrite_saved_history(client,session):
    other=jobs._new_session(key=session['key']);other['design']=design()
    ne.ensure(other,'design')  # Both windows read the same initial version.
    assert post(client,session,cmd()).status_code==200
    r=post(client,other,cmd())
    assert r.status_code==409
    assert len(other['design']['tables'].pipes)==2
    assert json.loads(ne._path(session,'design').read_text(encoding='utf-8'))['cursor']==1
    jobs._SESSIONS.pop(other['id'],None)


def test_native_fitting_loss_updates_for_changed_dn():
    t=table();t.fittings=[dict(pipe='P2',type='elbow',count=2)]
    c=cmd('pipe','P2');c['dn']=50
    n,_=apply_edit(Network.from_tables(t),c,ne.catalog())
    expected=ne.catalog()['fittings'][next(i for i,r in enumerate(ne.catalog()['fittings']) if r['id']=='ELBOW_90_STD')]['lengths']['50']*2
    assert n.pipes['P2'].row['eq_len']==expected


def test_merge_preserves_explicit_bore_and_inner_pair(client,session):
    from test_module_f_merge import _riser
    from routes.module_f.api_merge import rebuild_merged
    session['design']['tables'].pipes[1]['dia']=65
    c=cmd('pipe','P1');c['schedule']='CPVC2'
    assert post(client,session,c).status_code==200
    session['slots']['system']['riser']=_riser();session['supply_mode']='lsp_gravity'
    rebuild_merged(session)
    p=next(p for p in session['merged']['combined'].pipes if p['label']=='P1')
    assert p['dia']==25 and p['inner_mm']==28.02 and p['type']=='CPVC2'
