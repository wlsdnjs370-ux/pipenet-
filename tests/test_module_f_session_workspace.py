"""Confirmed-state isolation, failures, cancellation and HTTP writer races."""
from __future__ import annotations

import threading
import uuid

import pytest
from flask import Flask, jsonify, request

from routes.module_f import cancellation as cancel, jobs
from routes.module_f.session_workspace import SessionConflict, SessionWorkspace
from routes.module_f.session_policy import SELECTION_FIELDS


@pytest.fixture
def session():
    sess = jobs._new_session(key='workspace-fixture')
    sess['edit'] = {'nodes': [1, 2]}
    sess['worst'] = {'old': True}
    yield sess
    jobs._SESSIONS.pop(sess['id'], None)
    assert id(sess) not in cancel._SESSION_OWNERS
    assert id(sess) not in cancel._SESSION_READERS


def application(handler):
    app = Flask(__name__)
    app.testing = True
    @app.post('/api/module-f/edit/flow')
    def flow():
        return handler(jobs._sess(request.json['sid']))
    @app.post('/api/module-f/legacy')
    def legacy():
        jobs._sess(request.json['sid'])['key'] = 'changed'
        return jsonify(ok=True)
    @app.get('/api/module-f/edit/state')
    def state():
        return jsonify(ok=True, edit=jobs._sess(request.args['sid'])['edit'])
    cancel.install(app)
    return app


def test_success_publishes_owned_values_only_and_never_copies_other_slots(session):
    class NoCopy:
        def __deepcopy__(self, memo):
            raise AssertionError('unrelated drawing copied')
    unrelated = NoCopy()
    session['slots']['system']['geometry'] = unrelated
    old = session['edit']
    def handler(draft):
        assert draft is not session and draft['edit'] is not old
        draft['edit']['nodes'].append(3)
        draft['worst'] = {'new': True}
        assert old == {'nodes': [1, 2]} and session['worst'] == {'old': True}
        return jsonify(ok=True)
    r = application(handler).test_client().post('/api/module-f/edit/flow', json={'sid': session['id']})
    assert r.status_code == 200, r.json
    assert session['edit']['nodes'] == [1, 2, 3]
    assert session['worst'] == {'new': True}
    assert session['slots']['system']['geometry'] is unrelated
    assert session['_state_revision'] == 1


@pytest.mark.parametrize('failure', ['http', 'ok_false', 'exception'])
def test_failure_discards_working_state(session, failure):
    old = session['edit']
    def handler(draft):
        draft['edit']['nodes'].append(3)
        draft.pop('worst')
        if failure == 'exception':
            raise RuntimeError('fixture')
        return jsonify(ok=False), (400 if failure == 'http' else 200)
    c = application(handler).test_client()
    if failure == 'exception':
        with pytest.raises(RuntimeError, match='fixture'):
            c.post('/api/module-f/edit/flow', json={'sid': session['id']})
    else:
        c.post('/api/module-f/edit/flow', json={'sid': session['id']})
    assert session['edit'] is old and old['nodes'] == [1, 2]
    assert session['worst'] == {'old': True}


def test_stale_or_undeclared_changes_reject_publication(session):
    workspace = SessionWorkspace(session, SELECTION_FIELDS)
    workspace.working['edit']['nodes'].append(3)
    session['_state_revision'] = 7
    with pytest.raises(SessionConflict):
        workspace.commit()
    assert session['edit']['nodes'] == [1, 2]
    workspace = SessionWorkspace(session, SELECTION_FIELDS)
    workspace.working['unexpected_output'] = True
    with pytest.raises(SessionConflict, match='unexpected_output'):
        workspace.commit()
    assert 'unexpected_output' not in session


def test_inflight_draft_blocks_competing_writes_and_lazy_reads_but_not_polling(session):
    entered, release = threading.Event(), threading.Event()
    responses = []
    def handler(draft):
        draft['edit']['nodes'].append(3)
        entered.set()
        assert release.wait(3)
        return jsonify(ok=True)
    app = application(handler)
    token = uuid.uuid4().hex
    def run():
        with app.test_client() as c:
            responses.append(c.post('/api/module-f/edit/flow', json={'sid': session['id']},
                                    headers={'X-Module-F-Operation': token}))
    worker = threading.Thread(target=run)
    worker.start()
    try:
        assert entered.wait(3)
        with app.test_client() as c:
            assert c.post('/api/module-f/legacy', json={'sid': session['id']}).status_code == 409
            assert c.get('/api/module-f/edit/state', query_string={'sid': session['id']}).status_code == 409
            assert c.get('/api/module-f/performance', query_string={'sid': session['id']}).status_code == 200
            assert session['edit']['nodes'] == [1, 2]
            assert c.post('/api/module-f/job/cancel', json={'operation': token}).status_code == 200
    finally:
        release.set()
        worker.join(3)
    assert not worker.is_alive()
    assert responses[0].status_code == 499
    assert session['edit']['nodes'] == [1, 2]
    assert session['key'] == 'workspace-fixture'
    assert app.test_client().post('/api/module-f/legacy', json={'sid': session['id']}).status_code == 200


def test_duplicate_same_operation_does_not_discard_first_success(session):
    entered, release = threading.Event(), threading.Event()
    responses = []
    def handler(draft):
        draft['edit']['nodes'].append(3)
        entered.set()
        assert release.wait(3)
        return jsonify(ok=True)
    app = application(handler)
    headers = {'X-Module-F-Operation': uuid.uuid4().hex}
    def run():
        with app.test_client() as c:
            responses.append(c.post('/api/module-f/edit/flow', json={'sid': session['id']}, headers=headers))
    worker = threading.Thread(target=run)
    worker.start()
    try:
        assert entered.wait(3)
        r = app.test_client().post('/api/module-f/edit/flow', json={'sid': session['id']}, headers=headers)
        assert r.status_code == 409
    finally:
        release.set()
        worker.join(3)
    assert responses[0].status_code == 200
    assert session['edit']['nodes'] == [1, 2, 3]


def test_revision_is_session_wide_and_slot_change_rejects_draft(session):
    from routes.module_f.slots import _slot_switch
    session['_state_revision'] = 7
    workspace = SessionWorkspace(session, SELECTION_FIELDS)
    workspace.working['edit']['nodes'].append(3)
    _slot_switch(session, 'system')
    assert session['_state_revision'] == 7
    assert '_state_revision' not in session['slots']['plan']
    with pytest.raises(SessionConflict):
        workspace.commit()
    _slot_switch(session, 'plan')
    assert session['edit']['nodes'] == [1, 2]


def test_merge_mode_invalid_height_does_not_partially_change_mode(session, tmp_path):
    from routes.module_f import api_merge
    from routes.module_f.merge import SUPPLY_MODES
    app = Flask(__name__)
    app.testing = True
    api_merge.register(app, UPLOAD_DIR=tmp_path)
    cancel.install(app)
    modes = list(SUPPLY_MODES)
    session['supply_mode'] = modes[0]
    r = app.test_client().post('/api/module-f/merge/mode',
        json={'sid': session['id'], 'mode': modes[-1], 'source_drop_m': 'invalid'})
    assert r.status_code == 400
    assert session['supply_mode'] == modes[0]
