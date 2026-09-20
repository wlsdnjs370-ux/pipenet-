"""Mixed source/merge edits must preserve both graphs and their actual history."""
from copy import deepcopy

import pytest

from test_module_f_network_editor import client, session, post, state, cmd, design
from test_module_f_merge import _riser
from routes.module_f import api_merge, network_edit as ne


def combined(sess):
    sess['slots'] = {'system': {'riser': _riser()}}
    sess['supply_mode'] = 'lsp_gravity'
    api_merge.rebuild_merged(sess, persist_overrides=False)
    ne.ensure(sess,'merge')


def ok(response):
    assert response.status_code == 200, response.json
    return response.json


def test_system_edit_then_plan_edit_preserves_both_and_undo_order(client, session):
    combined(session)
    initial_r1=deepcopy(ne.ensure(session,'merge')['current'].pipes['r1'].row)
    ok(post(client,session,dict(op='pipe',target='r1',schedule='CPVC2',dn=50),scope='merge'))
    ok(post(client,session,dict(op='pipe',target='P2',schedule='CPVC2',dn=25),scope='merge'))
    def pipes(): return ne.ensure(session,'merge')['current'].pipes
    assert pipes()['r1'].row['type']=='CPVC2'
    assert pipes()['P2'].row['type']=='CPVC2'
    assert session['design']['tables'].pipes[1]['type']=='CPVC2'
    assert state(client,session,'merge')['undo']==2
    ok(post(client,session,action='undo',scope='merge'))
    assert pipes()['r1'].row['type']=='CPVC2'
    assert pipes()['P2'].row['type']=='KSD 3507'
    ok(post(client,session,action='undo',scope='merge'))
    assert all(pipes()['r1'].row.get(k)==v for k,v in initial_r1.items())
    assert pipes()['r1'].row.get('type')==initial_r1.get('type')
    ok(post(client,session,action='redo',scope='merge'))
    assert pipes()['r1'].row['type']=='CPVC2'
    ok(post(client,session,action='redo',scope='merge'))
    assert pipes()['P2'].row['type']=='CPVC2'
    session.pop('network_editor'); session.pop('merge_editor'); session['design']=design()
    ne.ensure(session,'design')
    api_merge.rebuild_merged(session,persist_overrides=False)
    assert not state(client,session,'merge')['conflict']
    assert pipes()['r1'].row['type']=='CPVC2' and pipes()['P2'].row['type']=='CPVC2'


def test_generated_node_and_pipe_collisions_rebase_and_preview_is_pure(client, session):
    combined(session)
    r=ok(post(client,session,dict(op='split',target='r1',distance=.5),scope='merge'))
    node=r['selection']['label']
    ok(post(client,session,dict(op='extend',target=node,axis='+Z',length=1,schedule='KSD 3507',dn=25),scope='merge'))
    before=ne.fingerprint(ne.ensure(session,'merge')['current'])
    disk=ne._path(session,'merge').read_bytes()
    c=cmd(target='11')
    ok(post(client,session,c,action='preview',scope='merge'))
    assert ne.fingerprint(ne.ensure(session,'merge')['current'])==before
    assert ne._path(session,'merge').read_bytes()==disk
    ok(post(client,session,c,scope='merge'))
    net=ne.ensure(session,'merge')['current']
    assert len(net.nodes)==9 and len(net.pipes)==8
    assert net.nodes['13'].row['editor_parent']=='11'  # new plan source owns 13
    assert net.nodes['15'].row['editor_parent']=='14'  # merge branch remapped
    assert net.nodes['15'].xyz[2]-net.nodes['14'].xyz[2]==pytest.approx(1)
    for _ in range(3): ok(post(client,session,action='undo',scope='merge'))
    for _ in range(3): ok(post(client,session,action='redo',scope='merge'))
    assert ne.fingerprint(ne.ensure(session,'merge')['current'])==ne.fingerprint(net)


def test_new_edit_discards_redo_in_both_scopes(client, session):
    combined(session)
    ok(post(client,session,dict(op='pipe',target='r1',schedule='CPVC2',dn=50),scope='merge'))
    ok(post(client,session,action='undo',scope='merge'))
    ok(post(client,session,cmd(target='11'),scope='merge'))
    assert state(client,session,'merge')['redo']==0


def test_second_file_failure_rolls_back_first_file_and_graph(client, session, monkeypatch):
    combined(session)
    ok(post(client,session,dict(op='pipe',target='r1',schedule='CPVC2',dn=50),scope='merge'))
    before={s:ne._path(session,s).read_bytes() for s in ('design','merge')}
    graph=ne.fingerprint(ne.ensure(session,'merge')['current'])
    plan=deepcopy(session['design']['tables'].as_dict())
    save=ne.save_history
    def fail(sess,scope,editor):
        if scope=='merge': raise OSError('save-failure')
        save(sess,scope,editor)
    monkeypatch.setattr(ne,'save_history',fail)
    assert post(client,session,cmd(target='11'),scope='merge').status_code==409
    assert session['design']['tables'].as_dict()==plan
    assert ne.fingerprint(ne.ensure(session,'merge')['current'])==graph
    assert {s:ne._path(session,s).read_bytes() for s in before}==before
    monkeypatch.setattr(ne,'save_history',save)
    ok(post(client,session,cmd(target='11'),scope='merge'))


def test_mixed_edits_from_inactive_plan_slot(client,session):
    from routes.module_f.slots import _slot_switch
    combined(session)
    _slot_switch(session,'system')
    ok(post(client,session,dict(op='pipe',target='r1',schedule='CPVC2',dn=50),scope='merge'))
    ok(post(client,session,cmd(target='11'),action='preview',scope='merge'))
    ok(post(client,session,cmd(target='11'),scope='merge'))
    assert len(ne.plan_state(session)['design']['tables'].pipes)==3
    assert ne.ensure(session,'merge')['current'].pipes['r1'].row['type']=='CPVC2'


def test_auto_zero_offset_preview_still_contains_the_combined_network(client,session):
    d=session['design'];t=d['tables'];d['method']='auto'
    for row in t.nodes: row['label']=str(int(row['label'])+9)
    for rows in (t.pipes,t.nozzles,t.equipment):
        for row in rows:
            for key in ('in','out'):
                if str(row.get(key,'')).isdigit(): row[key]=str(int(row[key])+9)
    d['keys']['nid']={str(int(k)+9):v for k,v in d['keys']['nid'].items()}
    combined(session)
    result=ok(post(client,session,cmd(target='11'),action='preview',scope='merge'))
    assert result['effective_scope']=='design'
    assert 'r1' in {p['label'] for p in result['preview']['pipes']}
    assert result['selection']['label']=='13'


def test_plan_only_merge_does_not_prevent_design_edits(client,session):
    session['merged']={'combined':None}
    ok(post(client,session,cmd()))
    assert len(session['design']['tables'].pipes)==3


def test_replay_missing_dependency_is_explicit_and_does_not_change_network():
    from core.network_editor import EditError,Network,apply_edit
    from core.network_edit_replay import rebase_commands
    base=Network.from_tables(design()['tables'])
    old,_=apply_edit(base,cmd(),ne.catalog())
    before=ne.fingerprint(base)
    with pytest.raises(EditError,match='원본 요소'):
        rebase_commands(old,base,[dict(op='node',target='4')],1,ne.catalog())
    assert ne.fingerprint(base)==before


def test_undo_selection_uses_combined_label_and_keeps_system_edit(client,session):
    combined(session)
    ok(post(client,session,cmd(target='11'),scope='merge'))
    ok(post(client,session,dict(op='pipe',target='r1',schedule='CPVC2',dn=50),scope='merge'))
    ok(post(client,session,dict(op='node',target='12'),scope='merge'))
    result=ok(post(client,session,action='undo',scope='merge'))
    assert result['selection']['label']=='13'
    assert result['selection']['label'] in ne.ensure(session,'merge')['current'].nodes
    assert ne.ensure(session,'merge')['current'].pipes['r1'].row['type']=='CPVC2'


def test_cancel_after_mixed_commit_restores_both_files(client,session):
    from core.network_editor import apply_edit
    from routes.module_f.api_network_edit import commit_edit
    from routes.module_f.cancellation import Operation,OperationCancelled
    combined(session)
    ok(post(client,session,dict(op='pipe',target='r1',schedule='CPVC2',dn=50),scope='merge'))
    files={s:ne._path(session,s).read_bytes() for s in ('design','merge')}
    before=ne.fingerprint(ne.ensure(session,'merge')['current'])
    editor=ne.ensure(session,'design')
    candidate,_=apply_edit(editor['current'],cmd(),ne.catalog())
    operation=Operation('mixed-editor-late-cancel');operation.snapshot(session)
    def commit_and_cancel():
        commit_edit(session,'design',editor,candidate,[cmd()],1)
        operation.cancel()
    with pytest.raises(OperationCancelled):
        with operation.scope(): commit_and_cancel()
    assert {s:ne._path(session,s).read_bytes() for s in files}==files
    assert ne.fingerprint(ne.ensure(session,'merge')['current'])==before
