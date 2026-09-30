"""A changed basis requires confirmation and exact archival, never label replay."""
from copy import deepcopy

import pytest

from test_module_f_network_editor import client, session, post, state, design
from routes.module_f import network_edit as ne


def conflict(client, session):
    """Create the H case: a saved property edit followed by a different base."""
    session['design_settings'] = {'diameter_policy': 'drawing_first_v1'}
    command = dict(op='pipe', target='P1', schedule='KSD 3507', dn=80, note='H 하단 속성표')
    assert post(client, session, command).status_code == 200
    old = ne._path(session, 'design').read_bytes()
    session['design'] = design()
    session['design']['tables'].pipes[0]['length'] = 7
    editor = ne.accept_rebuilt(session)
    assert editor['conflict']
    return old


def test_accept_basis_archives_and_does_not_replay_old_pipe_label(client, session):
    old = conflict(client, session)
    info = state(client, session)
    assert info['can_accept_basis'] and info['sid'] == session['id']
    assert '이전 편집 보관 후 계속' in info['conflict']
    response = post(client, session, action='accept_basis')
    assert response.status_code == 200, response.json
    editor = ne.ensure(session, 'design')
    assert not editor['conflict'] and not editor['commands']
    assert editor['current'].pipes['P1'].row['dia'] == 25
    assert editor['current'].pipes['P1'].row['length'] == 7
    assert any(path.read_bytes() == old for path, _ in ne._history_versions(session, 'design'))
    assert not state(client, session)['can_accept_basis']
    assert post(client, session, dict(op='pipe_c', target='P1', c=135)).status_code == 200
    assert post(client, session, action='undo').status_code == 200
    assert ne.ensure(session, 'design')['current'].pipes['P1'].row['c'] == 120
    # The exact original base can still restore its archived branch, not guesses.
    session['design'] = design()
    restored = ne.accept_rebuilt(session)
    assert restored['current'].pipes['P1'].row['dia'] == 80
    assert not restored['conflict']


@pytest.mark.parametrize('failure', ['backup', 'save', 'stale', 'changed_disk'])
def test_recovery_failures_preserve_saved_history_and_graph(client, session, monkeypatch, failure):
    old = conflict(client, session)
    path = ne._path(session, 'design')
    before = deepcopy(session['design']['tables'].as_dict())
    def fail(*args, **kwargs):
        raise OSError('test archival failure')
    if failure == 'backup':
        monkeypatch.setattr(ne, 'backup_history', fail)
    elif failure == 'save':
        monkeypatch.setattr(ne, 'save_history', fail)
    elif failure == 'stale':
        session['design']['h_attributes_dirty'] = True
    elif failure == 'changed_disk':
        path.write_bytes(old + b'\n')
        old = path.read_bytes()
    response = post(client, session, action='accept_basis')
    assert response.status_code == 409
    assert path.read_bytes() == old
    assert session['design']['tables'].as_dict() == before
    assert ne.ensure(session, 'design')['conflict']


def test_accept_basis_rejects_nonconflict_and_stale_revision(client, session):
    assert post(client, session, action='accept_basis').status_code == 409
    old = conflict(client, session)
    response = client.post('/api/module-f/network-editor', json=dict(
        sid=session['id'], scope='design', action='accept_basis', revision='old-revision'))
    assert response.status_code == 409
    assert ne._path(session, 'design').read_bytes() == old


def test_state_explains_stale_selection_without_touching_saved_history(client, session):
    old=conflict(client, session)
    session['design']['h_attributes_dirty']=True
    info=state(client, session)
    assert info['conflict'] and info['basis_stale']
    assert not info['can_accept_basis']
    assert ne._path(session,'design').read_bytes()==old
