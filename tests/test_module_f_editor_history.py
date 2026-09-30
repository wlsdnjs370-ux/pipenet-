"""Checkpoint replay equals full replay and cannot leak across history bases."""
from copy import deepcopy
import json

from core.network_editor import Network, apply_edit
from routes.module_f import editor_history as history, network_edit as ne
from routes.module_f.performance import report
from test_module_f_network_editor import table, design, client, session, post, state


def command(index):
    return dict(op='pipe', target='P2', schedule='KSD 3507', dn=25 if index % 2 else 32)


def editor_with_history(count=45):
    base = Network.from_tables(table())
    editor = dict(base=base, base_hash=ne.fingerprint(base), library='test-library',
                  current=base, commands=[], cursor=0)
    for i in range(count):
        graph, _ = apply_edit(editor['current'], command(i), ne.catalog())
        editor = dict(editor, current=graph, commands=editor['commands']+[command(i)], cursor=i+1)
        history.remember(editor)
    return editor


def test_recent_undo_replays_only_suffix_with_identical_hydraulic_graph():
    editor = editor_with_history()
    original = ne.fingerprint(editor['current'])
    expected, expected_selection = history.replay(dict(editor, _replay_cache=None),
        editor['commands'], 44, ne.catalog(), sid='uncached-history')
    actual, selection = history.replay(editor, editor['commands'], 44, ne.catalog(), sid='cached-history')
    assert ne.fingerprint(actual) == ne.fingerprint(expected)
    assert selection == expected_selection
    assert report('uncached-history')[-1]['replayed'] == 44
    assert report('cached-history')[-1]['replayed'] == 4
    assert history.cache_for(editor).count == history.MAX_CHECKPOINTS
    assert ne.fingerprint(editor['current']) == original
    assert deepcopy(editor['_replay_cache']) is editor['_replay_cache']
    # Returned models are never the checkpoint's private mutable model.
    actual.pipes['P2'].row['length'] = -100
    again, _ = history.replay(editor, editor['commands'], 44, ne.catalog())
    assert ne.fingerprint(again) == ne.fingerprint(expected)


def test_branch_prefix_base_and_library_invalidate_checkpoints():
    editor = editor_with_history()
    commands = deepcopy(editor['commands'])
    commands[0]['dn'] = 40
    changed, _ = history.replay(editor, commands, 44, ne.catalog(), sid='branch-history')
    expected, _ = history.replay(dict(editor, _replay_cache=None), commands, 44, ne.catalog())
    assert ne.fingerprint(changed) == ne.fingerprint(expected)
    assert report('branch-history')[-1]['replayed'] == 44
    assert history.cache_for(dict(editor, base_hash='different-base')).count == 0
    assert history.cache_for(dict(editor, library='new-library')).count == 0
    assert history.cache_for(editor).pruned([]).count == 0


def test_real_editor_undo_redo_branch_reset_and_reopen(client, session):
    for i in range(25):
        r = post(client, session, command(i))
        assert r.status_code == 200, r.json
    editor = ne.ensure(session, 'design')
    assert editor['_replay_cache'].count == 2
    expected = ne.fingerprint(editor['current'])
    saved = json.loads(ne._path(session, 'design').read_text(encoding='utf-8'))
    assert saved['version'] == 1 and '_replay_cache' not in saved
    assert post(client, session, action='undo').status_code == 200
    assert [e for e in report(session['id']) if e['phase'] == 'history_replay'][-1]['replayed'] == 4
    assert post(client, session, action='redo').status_code == 200
    assert ne.fingerprint(ne.ensure(session, 'design')['current']) == expected
    session.pop('network_editor')
    session['design'] = design()
    assert state(client, session)['undo'] == 25
    assert ne.fingerprint(ne.ensure(session, 'design')['current']) == expected
    assert ne.ensure(session, 'design')['_replay_cache'].count == 2
    assert post(client, session, action='undo').status_code == 200
    assert post(client, session, dict(command(0), dn=40)).status_code == 200
    assert state(client, session)['redo'] == 0
    assert post(client, session, action='reset').status_code == 200
    assert ne.ensure(session, 'design')['_replay_cache'].count == 0


def test_failed_commit_does_not_change_checkpoint_or_saved_history(client, session, monkeypatch):
    for i in range(9):
        assert post(client, session, command(i)).status_code == 200
    before = ne.ensure(session, 'design')
    before_graph = ne.fingerprint(before['current'])
    path = ne._path(session, 'design')
    data = path.read_bytes()
    def fail(*args):
        raise OSError('fixture disk full')
    monkeypatch.setattr(ne, 'save_history', fail)
    result = post(client, session, command(9))
    assert result.status_code == 409
    restored = ne.ensure(session, 'design')
    assert restored['cursor'] == 9 and restored['_replay_cache'].count == 0
    assert ne.fingerprint(restored['current']) == before_graph
    assert path.read_bytes() == data
