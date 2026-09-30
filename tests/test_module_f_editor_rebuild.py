"""Explicit table confirmation must preserve old edits without locking new bases."""
from copy import deepcopy
import json

import pytest

from test_module_f_network_editor import session,client,cmd,post,state,design
from test_module_f_design_speed import design_client,confirm
from routes.module_f import network_edit as ne,api_network_edit


def test_new_basis_archives_exact_bytes_and_old_basis_restores(client,session):
    assert post(client,session,cmd()).status_code==200
    old=ne._path(session,'design').read_bytes()
    session['design']=design();session['design']['tables'].pipes[0]['length']=333
    assert state(client,session)['conflict']  # An unsolicited graph change is still guarded.
    editor=ne.accept_rebuilt(session)
    assert not editor['conflict'] and editor['cursor']==0
    assert session['design']['tables'].pipes[0]['length']==333
    assert len(session['design']['tables'].pipes)==2
    archives=ne._history_versions(session,'design')
    assert any(p.read_bytes()==old for p,_ in archives)
    assert json.loads(ne._path(session,'design').read_bytes())['cursor']==0
    # Returning to exactly the original basis replays its own branch.
    session['design']=design()
    editor=ne.accept_rebuilt(session)
    assert editor['cursor']==1 and not editor['conflict']
    assert len(session['design']['tables'].pipes)==3
    assert len(session['design']['got']['kfp']['pipe_data'])==3


def test_identical_table_reconfirmation_preserves_edits_and_undo(client,session):
    assert post(client,session,cmd()).status_code==200
    session['design']=design()
    editor=ne.accept_rebuilt(session)
    assert editor['cursor']==1 and not ne._history_versions(session,'design')
    assert post(client,session,action='undo').status_code==200
    assert len(session['design']['tables'].pipes)==2


def test_h_auto_attributes_do_not_silently_archive_manual_edits(client,session):
    from routes.module_f.selection import _design_stale
    assert post(client,session,cmd()).status_code==200
    path=ne._path(session,'design');old=path.read_bytes()
    session['design_settings']={'diameter_policy':'drawing_first_v1'}
    session['design']=design();session['design']['tables'].pipes[0]['dia']=80
    editor=ne.accept_rebuilt(session)
    assert editor['conflict'] and editor['notice']
    assert path.read_bytes()==old
    assert _design_stale(session)['why']==[editor['conflict']]


def test_reopen_conflicting_v1_history_then_confirm_without_reset(client,session):
    assert post(client,session,cmd()).status_code==200
    session.pop('network_editor');session['design']=design()
    session['design']['tables'].pipes[0]['length']=444
    assert state(client,session)['conflict']
    ne.accept_rebuilt(session)
    assert not state(client,session)['conflict']
    assert post(client,session,cmd()).status_code==200
    info=client.get('/api/module-f/network-editor',query_string=dict(sid=session['id'],history='1')).json
    assert info['archived'][0]['commands'][0]['op']=='extend'


def test_redo_history_is_not_discarded_when_changing_basis(client,session):
    assert post(client,session,cmd()).status_code==200
    assert post(client,session,action='undo').status_code==200
    session['design']=design();session['design']['tables'].pipes[0]['length']=333
    ne.accept_rebuilt(session)
    session['design']=design();ne.accept_rebuilt(session)
    assert state(client,session)['redo']==1
    assert post(client,session,action='redo').status_code==200


def test_failed_backup_cannot_overwrite_previous_file(client,session,monkeypatch):
    assert post(client,session,cmd()).status_code==200
    path=ne._path(session,'design');old=path.read_bytes()
    session['design']=design();session['design']['tables'].pipes[0]['length']=333
    def fail(*args): raise OSError('backup failed')
    monkeypatch.setattr(ne,'backup_history',fail)
    with pytest.raises(OSError,match='backup failed'): ne.accept_rebuilt(session)
    assert path.read_bytes()==old


def test_real_table_route_and_export_after_old_history_conflict(design_client,tmp_path,monkeypatch):
    client,sess=design_client
    monkeypatch.setattr(ne,'HISTORY_DIR',tmp_path/'histories')
    # Install production guards, which used to reject before the build could run.
    api_network_edit.install_guards(client.application)
    confirm(client,sess)
    path=ne._path(sess,'design');path.parent.mkdir(parents=True,exist_ok=True)
    old=json.dumps(dict(version=1,base='old-selection',commands=[cmd(target='999')],
                        cursor=1,drawing=sess['key'],scope='design')).encode()
    path.write_bytes(old)
    sess.pop('network_editor');ne.ensure(sess,'design')
    assert sess['network_editor']['conflict']
    result=confirm(client,sess)
    assert result['ok'] and not sess['network_editor']['conflict']
    assert result['summary']['counts']['nozzles']==2
    assert any(p.read_bytes()==old for p,_ in ne._history_versions(sess,'design'))
    preview=client.get('/api/module-f/design/preview',query_string={'sid':sess['id']})
    assert preview.status_code==200,preview.json
    emit=client.post('/api/module-f/design/emit',json={'sid':sess['id']})
    assert emit.status_code==200,emit.json
    assert (tmp_path/'module_f'/f"{sess['id']}_design").is_dir()


def test_real_table_reconfirmation_keeps_live_pipe_properties(design_client,tmp_path,monkeypatch):
    client,sess=design_client
    monkeypatch.setattr(ne,'HISTORY_DIR',tmp_path/'histories')
    api_network_edit.register(client.application)
    api_network_edit.install_guards(client.application)
    confirm(client,sess)
    label=str(sess['design']['tables'].pipes[-1]['label'])
    r=post(client,sess,dict(op='pipe',target=label,schedule='CPVC2',dn=25))
    assert r.status_code==200,r.json
    base=ne.ensure(sess,'design')['base_hash']
    confirm(client,sess)
    editor=ne.ensure(sess,'design')
    assert editor['base_hash']==base
    assert editor['cursor']==1 and editor['current'].pipes[label].row['type']=='CPVC2'


def test_cancel_reconfirmation_restores_active_file_and_preserves_archive(client,session):
    from routes.module_f.cancellation import Operation,OperationCancelled
    assert post(client,session,cmd()).status_code==200
    old=ne._path(session,'design').read_bytes()
    operation=Operation('table-history-rebuild');operation.snapshot(session)
    def commit_and_cancel():
        session['design']=design();session['design']['tables'].pipes[0]['length']=333
        ne.accept_rebuilt(session)
        operation.cancel()
    with pytest.raises(OperationCancelled):
        with operation.scope(): commit_and_cancel()
    assert ne._path(session,'design').read_bytes()==old
    assert len(session['design']['tables'].pipes)==3
    assert any(p.read_bytes()==old for p,_ in ne._history_versions(session,'design'))
