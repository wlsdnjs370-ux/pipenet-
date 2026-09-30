"""Property batches are atomic, replayable, and do not change graph geometry."""
from copy import deepcopy

import pytest

from core.network_editor import Network, apply_edit, EditError
from core.network_edit_replay import rebase_commands
from core.network_property_fittings import fitting_rows
from test_module_f_network_editor import table, design, ne, session, client, post, state


def batch(*commands):
    return dict(op='batch_properties', commands=list(commands))


def test_batch_all_or_nothing_and_graph_preservation():
    original = Network.from_tables(table())
    before = ne.fingerprint(original)
    commands = batch(dict(op='pipe_c', target='P1', c=140), dict(op='pipe_c', target='P2', c=135))
    changed, _ = apply_edit(original, commands, ne.catalog())
    assert [p.row['c'] for p in changed.pipes.values()] == [140, 135]
    assert [(p.a,p.b,p.row['length']) for p in changed.pipes.values()] == [(p.a,p.b,p.row['length']) for p in original.pipes.values()]
    assert {k:n.xyz for k,n in original.nodes.items()} == {k:n.xyz for k,n in changed.nodes.items()}
    commands['commands'][1]['c'] = 'NaN'
    with pytest.raises(EditError, match='2번'):
        apply_edit(original, commands, ne.catalog())
    assert ne.fingerprint(original) == before


@pytest.mark.parametrize('commands', [[], [dict(op='delete',target='P2')], [dict(op='batch_properties',commands=[])]])
def test_batch_rejects_geometry_and_nested_commands(commands):
    with pytest.raises(EditError):
        apply_edit(Network.from_tables(table()), batch(*commands), ne.catalog())


def test_batch_one_undo_redo_and_reopen(client, session):
    command=batch(dict(op='pipe_c',target='P1',c=140),dict(op='pipe_c',target='P2',c=145))
    assert post(client,session,command,action='preview').status_code == 200
    assert [p['c'] for p in session['design']['tables'].pipes] == [120,120]
    assert post(client,session,command).status_code == 200
    assert state(client,session)['undo'] == 1
    assert post(client,session,action='undo').status_code == 200
    assert [p['c'] for p in session['design']['tables'].pipes] == [120,120]
    assert post(client,session,action='redo').status_code == 200
    session.pop('network_editor'); session['design']=design()
    assert state(client,session)['undo'] == 1
    assert [p['c'] for p in session['design']['tables'].pipes] == [140,145]


def test_fitting_batch_changes_existing_losses_not_topology():
    t=table();lib=ne.catalog()
    from core.network_editor import fitting_value
    t.fittings=[dict(pipe=p['label'],node=p['in'],type='elbow',count=1) for p in t.pipes]
    for p in t.pipes:p['eq_len']=fitting_value(lib,'ELBOW_90_STD',25)
    original=Network.from_tables(t)
    rows=fitting_rows(original,lib)
    assert len(rows)==2
    commands=[dict(op='fitting_properties',target=r['pipe'],collection=r['collection'],index=r['index'],expected_type=r['expected_type'],count=2) for r in rows]
    changed,_=apply_edit(original,batch(*commands),lib)
    assert [r['count'] for r in changed.tables.fittings]==[2,2]
    assert all(changed.pipes[k].row['eq_len']==pytest.approx(p.row['eq_len']*2) for k,p in original.pipes.items())
    assert len(changed.tables.equipment)==len(original.tables.equipment)
    assert ne.fingerprint(original)==ne.fingerprint(Network.from_tables(t))


def test_rebase_keeps_property_batch_as_one_unit():
    original=Network.from_tables(table());new=deepcopy(original)
    new.nodes['1'].xyz=(1,2,3)
    command=batch(dict(op='pipe_c',target='P1',c=130),dict(op='head',target='3',nozzle='SP-HEAD',flow=95))
    result,commands=rebase_commands(original,new,[command],1,ne.catalog())
    assert len(commands)==1 and result.pipes['P1'].row['c']==130
    assert result.tables.nozzles[0]['flow_lmin']==95
    assert result.nodes['1'].xyz==(1,2,3)


def test_merge_mixed_owner_batch_rejected_without_partial_write(client,session):
    from test_module_f_merge import _riser
    from routes.module_f.api_merge import rebuild_merged
    session['slots']['system']['riser']=_riser()
    session['supply_mode']='lsp_gravity';rebuild_merged(session)
    before=deepcopy(session['merged']['combined'].pipes)
    system=next(p['label'] for p in before if str(p['label']).startswith('r'))
    command=batch(dict(op='pipe_c',target='P1',c=135),dict(op='pipe_c',target=system,c=135))
    response=post(client,session,command,scope='merge')
    assert response.status_code==409
    assert '나누어' in response.json['message']
    assert session['merged']['combined'].pipes==before
