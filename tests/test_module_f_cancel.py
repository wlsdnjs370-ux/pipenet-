"""Real worker/request cancellation, state rollback and persistence boundaries."""
from __future__ import annotations

import os
import sys
import threading
import time
import uuid
from pathlib import Path

import pytest
from flask import Flask, jsonify

from routes.module_f import cancellation as cancel
from routes.module_f import jobs


def wait_until(predicate, timeout=4):
    deadline = time.monotonic() + timeout
    while not predicate():
        if time.monotonic() > deadline:
            pytest.fail("Worker did not settle within the test timeout")
        time.sleep(.01)


@pytest.fixture
def session():
    sess = jobs._new_session()
    sess['confirmed'] = {'heads': [1, 2], 'length': 12.5}
    yield sess
    if jobs._job_running(sess):
        cancel.operation(sess['job']['operation']).cancel()
        wait_until(lambda: not jobs._job_running(sess))
    jobs._SESSIONS.pop(sess['id'], None)


def test_running_cpu_job_stops_and_restores_confirmed_state(session):
    entered = threading.Event()
    stream = sys.stdout
    def calculate():
        session['confirmed']['heads'].append(999)
        session['new_partial'] = True
        entered.set()
        n = 0
        while True:
            n += 1
    job = jobs._run_job(session, 'CPU', calculate)
    assert entered.wait(3)
    before = time.monotonic()
    cancel.operation(job['operation']).cancel()
    wait_until(lambda: job['state'] != 'run')
    assert time.monotonic() - before < 2
    assert job['state'] == 'cancelled'
    assert session['confirmed'] == {'heads': [1, 2], 'length': 12.5}
    assert 'new_partial' not in session
    assert sys.stdout is stream
    assert not jobs._HEAVY_LOCK.locked()
    next_job = jobs._run_job(session, 'next', lambda: {'ok': True})
    wait_until(lambda: next_job['state'] != 'run')
    assert next_job['state'] == 'done'
    assert next_job['result'] == {'ok': True}


def test_queued_job_stops_without_waiting_for_heavy_lock(session):
    ran = []
    with jobs._HEAVY_LOCK:
        job = jobs._run_job(session, 'queued', lambda: ran.append(True))
        assert jobs._job_view(session)['queued']
        cancel.operation(job['operation']).cancel()
        wait_until(lambda: job['state'] == 'cancelled')
    assert ran == []


def test_stop_waits_for_open_output_to_close(session, tmp_path):
    entered, finish = threading.Event(), threading.Event()
    path = tmp_path / 'result.txt'
    def output():
        with path.open('w') as file:
            file.write('complete ')
            entered.set()
            finish.wait(3)
            file.write('file')
        return {'ok': True}
    job = jobs._run_job(session, 'output', output)
    assert entered.wait(3)
    token = cancel.operation(job['operation'])
    token.cancel()
    assert not token.status()['stopped']
    finish.set()
    wait_until(lambda: job['state'] != 'run')
    assert path.read_text() == 'complete file'
    assert job['state'] == 'cancelled'


def test_sync_request_can_be_cancelled_and_late_request_is_rejected(session):
    app = Flask(__name__)
    entered = threading.Event()
    @app.post('/api/module-f/test-compute')
    def compute():
        session['confirmed']['length'] = -1
        entered.set()
        n = 0
        while True:
            n += 1
    cancel.install(app)
    token = uuid.uuid4().hex
    responses = []
    def request_work():
        with app.test_client() as client:
            responses.append(client.post('/api/module-f/test-compute',
                json={'sid': session['id']}, headers={'X-Module-F-Operation': token}))
    thread = threading.Thread(target=request_work)
    thread.start()
    assert entered.wait(3)
    with app.test_client() as client:
        response = client.post('/api/module-f/job/cancel', json={'operation': token})
        assert response.status_code == 200
        thread.join(3)
        assert not thread.is_alive()
        assert responses[0].status_code == 499
        assert session['confirmed']['length'] == 12.5
        state = client.get('/api/module-f/job/cancel-state', query_string={'operation': token}).get_json()
        assert state['stopped']
        late = client.post('/api/module-f/test-compute', json={'sid':session['id']},
                           headers={'X-Module-F-Operation':token})
        assert late.status_code == 499


def test_cancel_before_request_arrives_does_not_start_it():
    app = Flask(__name__)
    ran = []
    @app.post('/api/module-f/test-late')
    def late():
        ran.append(True)
        return jsonify(ok=True)
    cancel.install(app)
    with app.test_client() as client:
        token = uuid.uuid4().hex
        assert client.post('/api/module-f/job/cancel', json={'operation':token}).status_code == 200
        assert client.post('/api/module-f/test-late', headers={'X-Module-F-Operation':token}).status_code == 499
    assert ran == []


def test_stop_old_job_id_does_not_cancel_new_job(session):
    app = Flask(__name__)
    cancel.install(app)
    job = jobs._run_job(session, 'new', lambda: {'ok': True})
    wait_until(lambda: job['state'] != 'run')
    with app.test_client() as client:
        result = client.post('/api/module-f/job/cancel',
                             json={'sid':session['id'], 'job_id':'old'})
        assert result.status_code == 409
    assert not cancel.operation(job['operation']).cancelled


def test_success_and_engine_exit_release_worker(session):
    def fail():
        raise SystemExit('fixture error')
    job = jobs._run_job(session, 'failure', fail)
    wait_until(lambda: job['state'] != 'run')
    assert job['state'] == 'error'
    assert not jobs._HEAVY_LOCK.locked()
    assert cancel.operation(job['operation']).status()['stopped']


def test_monitor_is_disabled_after_cancel(session):
    entered = threading.Event()
    def compute():
        entered.set()
        while True:
            pass
    job = jobs._run_job(session, 'CPU', compute)
    assert entered.wait(3)
    cancel.operation(job['operation']).cancel()
    wait_until(lambda: job['state'] != 'run')
    if cancel._MONITOR_ID is not None:
        assert sys.monitoring.get_events(cancel._MONITOR_ID) == 0


def test_fallback_trace_stops_without_monitoring(session, monkeypatch):
    monkeypatch.setattr(cancel, '_MONITORING', None)
    monkeypatch.setattr(cancel, '_MONITOR_ID', None)
    entered = threading.Event()
    def calculate():
        entered.set()
        while True:
            pass
    job = jobs._run_job(session, 'fallback', calculate)
    assert entered.wait(3)
    cancel.operation(job['operation']).cancel()
    wait_until(lambda: job['state'] != 'run')
    assert job['state'] == 'cancelled'


def test_cancelling_queued_session_leaves_other_session_running(session):
    entered, release = threading.Event(), threading.Event()
    other = jobs._new_session()
    def keep_running():
        entered.set()
        release.wait(3)
        return {'ok': True}
    first = jobs._run_job(other, 'other', keep_running)
    try:
        assert entered.wait(3)
        queued = jobs._run_job(session, 'queued', lambda: {'ok': True})
        cancel.operation(queued['operation']).cancel()
        wait_until(lambda: queued['state'] == 'cancelled')
        assert first['state'] == 'run'
        assert not cancel.operation(first['operation']).cancelled
    finally:
        release.set()
        wait_until(lambda: first['state'] != 'run')
        jobs._SESSIONS.pop(other['id'], None)
    assert first['state'] == 'done'


@pytest.mark.parametrize('value', ['', '../bad', 123, ['bad']])
def test_cancel_rejects_invalid_operation_ids(value):
    app = Flask(__name__)
    cancel.install(app)
    with app.test_client() as client:
        response = client.post('/api/module-f/job/cancel', json={'operation':value})
        assert response.status_code == 400


# ── 2026-09-20 실측 회귀 — 「중지」가 서버 전체를 잠그던 자리 ─────────────────
#
# 손질 재구성은 edit/io.py(open_board) → stage1_body → pipeline 로 내려간다.
# 종전 규칙은 «보호 모듈이 조상에 하나라도 있으면 중지 불가» 였고, 그러면
# 중지 요청 뒤 매 줄마다 스택 전체를 훑는 검사가 돌아 0.9 s 작업이 수십 분이
# 됐다(측정: 2.7 s → 228 s). 그 스레드가 _HEAVY_LOCK 을 쥔 채 GIL 을 독점해
# 다른 모든 업로드가 «대기» 로 굳었다.

_PERSISTENCE_SRC = """
def open_board(fn):
    # 보호 모듈(edit/io.py)의 함수가 계산을 '부르기만' 하는 자리
    return fn()
"""


def _persistence_caller():
    """co_filename 이 edit/io.py 로 끝나는 함수 — 보호 모듈의 조상 역할."""
    root = str(Path(cancel.__file__).resolve().parents[2])
    fake = os.path.join(root, "services", "cad_import", "edit", "io.py")
    namespace = {}
    exec(compile(_PERSISTENCE_SRC, fake, "exec"), namespace)
    return namespace["open_board"]


def _spin(n):
    acc = 0
    for i in range(n):
        acc += i % 7
    return acc


def test_stop_lands_inside_computation_called_from_persistence(session):
    entered = threading.Event()
    open_board = _persistence_caller()
    def rebuild():
        entered.set()
        while True:
            pass
    job = jobs._run_job(session, "손질", lambda: open_board(rebuild))
    assert entered.wait(3)
    started = time.monotonic()
    cancel.operation(job["operation"]).cancel()
    wait_until(lambda: job["state"] != "run", timeout=4)
    assert job["state"] == "cancelled"
    assert time.monotonic() - started < 2


def save_result(n):
    """이름 규칙(save*)으로 정당하게 보호되는 구간 — 끝난 뒤에만 중지된다."""
    return _spin(n)


def test_stop_under_persistence_does_not_slow_the_job(session):
    n = 300_000
    t = time.perf_counter()
    _spin(n)
    baseline = time.perf_counter() - t
    entered = threading.Event()
    def protected_work():
        entered.set()
        return save_result(n * 3)
    job = jobs._run_job(session, "저장", protected_work)
    assert entered.wait(3)
    cancel.operation(job["operation"]).cancel()
    t = time.perf_counter()
    wait_until(lambda: job["state"] != "run", timeout=30)
    elapsed = time.perf_counter() - t
    assert job["state"] == "cancelled"          # 보호 구간이 끝난 뒤 중지된다
    # 종전: 13~85 배. 검사를 쉬는 동안(0.05 s) 은 원래 속도로 돈다.
    assert elapsed < baseline * 3 * 4 + 1.0, (elapsed, baseline)
    if cancel._MONITOR_ID is not None:
        wait_until(lambda: cancel._PAUSE_TIMER is None, timeout=2)
        assert sys.monitoring.get_events(cancel._MONITOR_ID) == 0
