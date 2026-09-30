"""Cancel Module F operations at Python checkpoints, retaining confirmed state.

CPython 3.12+ monitoring is enabled only while a stop request is pending.
Native calls finish before the next checkpoint; active file writes are allowed
to close. No thread is killed and no exception is injected with ctypes.
"""
from __future__ import annotations

import copy
import functools
import io
import re
import sys
import threading
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

from flask import jsonify, request
from routes.module_f.session_workspace import SessionConflict


class OperationCancelled(BaseException):
    """Control flow, deliberately not swallowed by engine Exception handlers."""


_LOCAL = threading.local()
_LOCK = threading.RLock()
_READ_CONDITION = threading.Condition(_LOCK)
_OPERATIONS: dict[str, "Operation"] = {}
_ROOT = str(Path(__file__).resolve().parents[2]).lower().replace("\\", "/") + "/"
_MONITOR_ID = None
_MONITORING = getattr(sys, "monitoring", None)
_PROTECTED_MODULES = ("/pick/io.py", "/edit/io.py", "/pipeline/handoff.py",
                      "/pipeline/disp_cache.py")
_INFRA_MODULES = ("/module_f/cancellation.py", "/module_f/jobs.py",
                  "/module_f/session_workspace.py", "/module_f/performance.py")
_CONTROL_PATHS = {"/api/module-f/job", "/api/module-f/job/stream",
                  "/api/module-f/job/cancel", "/api/module-f/job/cancel-state",
                  "/api/module-f/performance"}
_SESSION_OWNERS: dict[int, "Operation"] = {}
_SESSION_READERS: dict[int, int] = {}


def _interruptible(frame) -> bool:
    """Keep persistence and infrastructure out of computational checkpoints.

    A protected module (persistence) blocks the stop only while the innermost
    project code is its own — i.e. nothing but protected/infra/third-party
    frames sit between here and it. Computation that a protected function
    merely *calls* (open_board → stage1_body → pipeline) stays interruptible:
    the exception unwinds through the caller before it opens any file, and a
    file that is already open is caught by the writable-handle scan below.
    Measured 2026-09-20: treating every ancestor in edit/io.py as protected
    made 「중지」 during 손질 재구성 undeliverable — see `_line`.
    """
    path = frame.f_code.co_filename.lower().replace("\\", "/")
    if not (path.startswith(_ROOT) or "/ezdxf/" in path):
        return False
    if path.endswith(_INFRA_MODULES):
        return False
    computation_below = False  # a plain project frame between here and the ancestor
    while frame is not None:
        filename = frame.f_code.co_filename.lower().replace("\\", "/")
        name = frame.f_code.co_name.lower().lstrip("_")
        protected = filename.endswith(_PROTECTED_MODULES)
        if protected and not computation_below:
            return False
        if name.startswith(("write", "save", "dump", "emit", "commit", "backup")):
            return False
        # Also protect user/third-party writers with arbitrary function names.
        if filename.startswith(_ROOT) and not filename.endswith(_INFRA_MODULES):
            for value in tuple(frame.f_locals.values()):
                if isinstance(value, io.IOBase) and not value.closed:
                    try:
                        if value.writable():
                            return False
                    except (OSError, ValueError):
                        pass
            if not protected:
                computation_below = True
        frame = frame.f_back
    return True


# A line that proves non-interruptible must not be re-examined on the very
# next line. The stack walk (every ancestor's filename, name and f_locals)
# costs tens of µs, and even the bare per-line callback costs ~0.5 µs; doing
# either on *every* line of a job that cannot be stopped yet slowed the job
# 85× and starved every other request of the GIL (measured 2026-09-20:
# 2.7 s job → 228 s; an unrelated 0.06 ms request → 2 ms). So after a failed
# check the LINE events are switched off for `_PAUSE_SECONDS` and switched
# back on by a timer — a stop still lands within that pause once the
# protected region returns. The per-thread `skip` counter is the same idea
# for the settrace fallback, where events cannot be paused.
_PAUSE_SECONDS = 0.05
_BACKOFF_LINES = 4096
_PAUSE_TIMER = None


def _pause_monitoring() -> None:
    global _PAUSE_TIMER
    if _MONITORING is None or _MONITOR_ID is None:
        return
    with _LOCK:
        if _PAUSE_TIMER is not None:
            return
        _MONITORING.set_events(_MONITOR_ID, 0)
        _PAUSE_TIMER = threading.Timer(_PAUSE_SECONDS, _resume_monitoring)
        _PAUSE_TIMER.daemon = True
        _PAUSE_TIMER.start()


def _resume_monitoring() -> None:
    global _PAUSE_TIMER
    with _LOCK:
        _PAUSE_TIMER = None
        _monitor_refresh()


def _pending_stop(frame) -> bool:
    """True when this thread's operation is cancelled and `frame` may raise."""
    op = getattr(_LOCAL, "operation", None)
    if op is None or not op.cancelled or getattr(_LOCAL, "delivered", False):
        return False
    skip = getattr(_LOCAL, "skip", 0)
    if skip:
        _LOCAL.skip = skip - 1
        return False
    if _interruptible(frame):
        _LOCAL.delivered = True  # Do not interrupt exception/finally cleanup again.
        return True
    _LOCAL.skip = _BACKOFF_LINES
    _pause_monitoring()
    return False


def _line(_code, _line_number) -> None:
    if _pending_stop(sys._getframe(1)):
        raise OperationCancelled("작업을 중지했습니다.")


def _trace(frame, event, arg):
    # Python 3.11 fallback. Current local server uses 3.13 monitoring instead.
    if event == "line" and _pending_stop(frame):
        raise OperationCancelled("작업을 중지했습니다.")
    return _trace


def _monitor_refresh() -> None:
    global _MONITOR_ID
    if _MONITORING is None:
        return
    pending = any(op.cancelled and op.active for op in _OPERATIONS.values())
    if _MONITOR_ID is None:
        for index in (3, 4, 5, 2, 1, 0):
            if _MONITORING.get_tool(index) is None:
                _MONITORING.use_tool_id(index, "module-f-cancel")
                _MONITOR_ID = index
                _MONITORING.register_callback(index, _MONITORING.events.LINE, _line)
                break
    if _MONITOR_ID is not None:
        _MONITORING.set_events(_MONITOR_ID, _MONITORING.events.LINE if pending else 0)


def current_operation():
    """Return the operation belonging to this request/worker, if any."""
    return getattr(_LOCAL, "operation", None)


def checkpoint() -> None:
    """Explicit checkpoint before starting work and before accepting results."""
    op = current_operation()
    if op is not None and op.cancelled:
        _LOCAL.delivered = True
        raise OperationCancelled("작업을 중지했습니다.")


def _session_copy(sess: dict, fields: frozenset[str] | None = None) -> dict:
    """Copy editable state; retain immutable DXF geometry and spatial indexes."""
    memo = {}
    states = [sess]
    if fields is None or 'slots' in fields:
        states.extend((sess.get("slots") or {}).values())
    for state in states:
        for name in ("world", "entities", "recon"):
            value = state.get(name)
            if value is not None:
                memo[id(value)] = value
        pick = state.get("pick")
        if pick is not None:
            memo[id(pick.world)] = pick.world
            for name in ("w", "_sgrid", "by_bundle", "_csmall", "_cgrid", "fp_index"):
                value = getattr(pick.board, name, None)
                if value is not None:
                    memo[id(value)] = value
    from routes.module_f.performance import measure
    state = {k: v for k, v in sess.items() if k not in {"job", "log", "touched"}
             and (fields is None or k in fields)}
    with measure('session_snapshot', sid=sess.get('id',''), keys=len(state), shared=len(memo)):
        return copy.deepcopy(state, memo)


class Operation:
    """One UI Working operation, possibly spanning a request and a worker."""

    def __init__(self, identifier: str):
        self.id = identifier
        self.cancelled = False
        self.active = 0
        self.touched = time.monotonic()
        self.snapshots: dict[int, tuple[dict, dict, frozenset[str] | None]] = {}
        self.rollback_callbacks: list = []
        self.rollback_errors: list[str] = []
        self.workspaces: dict[int, object] = {}
        self.sessions: dict[int, dict] = {}
        self.failed = False
        self.draft_pending = False

    def reserve(self) -> None:
        with _LOCK:
            if self.cancelled:
                raise OperationCancelled("작업을 중지했습니다.")
            if self.active == 0:
                self.failed = False
            self.active += 1
            self.touched = time.monotonic()
            _monitor_refresh()

    def snapshot(self, sess: dict, fields: frozenset[str] | None = None) -> None:
        if id(sess) not in self.snapshots:
            self.snapshots[id(sess)] = (sess, _session_copy(sess, fields), fields)

    def claim(self, sess: dict, *, draft: bool = False, wait_readers: float = 0) -> None:
        """One writer per session, including requests without a UI token."""
        from routes.module_f.session_workspace import SessionConflict
        with _LOCK:
            deadline = time.monotonic() + wait_readers
            while _SESSION_READERS.get(id(sess)) and id(sess) not in _SESSION_OWNERS and time.monotonic() < deadline:
                if self.cancelled:
                    raise OperationCancelled('작업을 중지했습니다.')
                _READ_CONDITION.wait(min(.05, max(0, deadline-time.monotonic())))
            if id(sess) in _SESSION_OWNERS or _SESSION_READERS.get(id(sess)):
                raise SessionConflict('이 도면의 작업이 진행 중입니다. 완료 후 다시 실행하세요.')
            _SESSION_OWNERS[id(sess)] = self
            self.sessions[id(sess)] = sess
            self.draft_pending = draft

    def draft(self, sess: dict, fields: frozenset[str]) -> None:
        from routes.module_f.session_workspace import SessionWorkspace
        self.workspaces[id(sess)] = SessionWorkspace(sess, fields)

    def resolve(self, sess: dict) -> dict:
        """Only the owning request sees its draft; polling sees confirmed data."""
        workspace = self.workspaces.get(id(sess))
        return workspace.working if workspace is not None else sess

    def release(self) -> None:
        with _LOCK:
            self.active -= 1
            self.touched = time.monotonic()
            if self.active == 0:
                try:
                    self._finish()
                finally:
                    for identity in self.sessions:
                        _SESSION_OWNERS.pop(identity, None)
                    self.sessions.clear()
                    self.workspaces.clear()
                    self.draft_pending = False
                    self.snapshots.clear()
                    self.rollback_callbacks.clear()
                    _monitor_refresh()
            _monitor_refresh()

    def _finish(self) -> None:
        """Finalize with the lock held; draft failures never restore over live data."""
        if not self.cancelled and not self.failed:
            for workspace in self.workspaces.values():
                workspace.validate()
            for workspace in self.workspaces.values():
                workspace.commit()
        if not self.cancelled:
            for sess in self.sessions.values():
                sess['_state_revision'] = sess.get('_state_revision', 0) + 1
        if self.cancelled:
            for rollback in reversed(self.rollback_callbacks):
                try:
                    rollback()
                except OSError as exc:
                    self.rollback_errors.append(str(exc))
            for sess, before, fields in self.snapshots.values():
                runtime = {k: sess[k] for k in ("job", "log", "touched") if k in sess}
                sess.update(before)
                for key in (set(sess) if fields is None else fields) - set(before) - set(runtime):
                    sess.pop(key, None)
                sess.update(runtime)

    @contextmanager
    def scope(self, *, reserved: bool = False):
        """Bind cancellation to one thread; native work is never killed."""
        if not reserved:
            self.reserve()
        previous = current_operation()
        previous_delivered = getattr(_LOCAL, "delivered", False)
        previous_trace = sys.gettrace()
        previous_skip = getattr(_LOCAL, "skip", 0)
        _LOCAL.operation, _LOCAL.delivered, _LOCAL.skip = self, False, 0
        use_trace = _MONITOR_ID is None
        if use_trace:
            sys.settrace(_trace)
        try:
            checkpoint()
            yield self
            checkpoint()
        except BaseException:
            self.failed = True
            raise
        finally:
            _LOCAL.operation, _LOCAL.delivered = previous, previous_delivered
            _LOCAL.skip = previous_skip
            if use_trace:
                sys.settrace(previous_trace)
            self.release()

    def cancel(self) -> dict:
        with _LOCK:
            self.cancelled = True
            self.touched = time.monotonic()
            _monitor_refresh()
            return self.status()

    def status(self) -> dict:
        return {"ok": True, "operation": self.id, "cancelled": self.cancelled,
                "stopped": self.active == 0, "active": self.active,
                "rollback_errors": list(self.rollback_errors)}


def operation(identifier: str | None = None) -> Operation:
    """Find/create an unguessable token; retain cancelled tombstones for races."""
    identifier = uuid.uuid4().hex if identifier is None else identifier
    if not isinstance(identifier, str) or not re.fullmatch(r"[a-zA-Z0-9_-]{16,80}", identifier):
        raise ValueError("작업 번호가 올바르지 않습니다.")
    with _LOCK:
        now = time.monotonic()
        for key in [key for key, value in _OPERATIONS.items()
                    if not value.active and now - value.touched > 1800]:
            del _OPERATIONS[key]
        return _OPERATIONS.setdefault(identifier, Operation(identifier))


def install(app, *, register_controls: bool = True, endpoint_prefix: str | None = None) -> None:
    """Wrap every F request carrying the UI operation id, including uploads."""
    @(app.get("/api/module-f/performance") if register_controls else lambda fn: fn)
    def module_f_performance():
        from routes.module_f.jobs import _sess
        from routes.module_f.performance import report
        try:
            sess = _sess(request.args.get('sid'))
        except ValueError as exc:
            return jsonify(ok=False, message=str(exc)), 410
        return jsonify(ok=True, timings=report(sess['id']))

    @(app.post("/api/module-f/job/cancel") if register_controls else lambda fn: fn)
    def module_f_job_cancel():
        body = request.get_json(silent=True) or {}
        identifier = body.get("operation")
        if not identifier and body.get("sid"):
            from routes.module_f.jobs import _sess
            try:
                job = _sess(body["sid"]).get("job") or {}
                identifier = job.get("operation")
                if body.get("job_id") and body["job_id"] != job.get("id"):
                    return jsonify(ok=False, message="이미 다른 작업으로 바뀌었습니다."), 409
            except ValueError as exc:
                return jsonify(ok=False, message=str(exc)), 410
        if not identifier:
            return jsonify(ok=False, message="중지할 작업 번호가 없습니다."), 400
        try:
            return jsonify(operation(identifier).cancel())
        except ValueError as exc:
            return jsonify(ok=False, message=str(exc)), 400

    @(app.get("/api/module-f/job/cancel-state") if register_controls else lambda fn: fn)
    def module_f_cancel_state():
        identifier = request.args.get("operation", "")
        try:
            return jsonify(operation(identifier).status())
        except ValueError as exc:
            return jsonify(ok=False, message=str(exc)), 400

    @contextmanager
    def reading():
        """Lazy GET caches must not overlap a draft's capture/publication."""
        sess = None
        if request.method != 'POST' and request.args.get('sid'):
            from routes.module_f.jobs import _sess
            try:
                sess = _sess(request.args['sid'])
            except ValueError:
                pass
        if sess is not None:
            with _LOCK:
                owner = _SESSION_OWNERS.get(id(sess))
                if owner and (owner.draft_pending or owner.workspaces):
                    raise SessionConflict('계산 중입니다. 확정 결과가 나온 뒤 다시 조회하세요.')
                _SESSION_READERS[id(sess)] = _SESSION_READERS.get(id(sess), 0) + 1
        try:
            yield
        finally:
            if sess is not None:
                with _LOCK:
                    count = _SESSION_READERS[id(sess)] - 1
                    if count:
                        _SESSION_READERS[id(sess)] = count
                    else:
                        _SESSION_READERS.pop(id(sess), None)
                    _READ_CONDITION.notify_all()

    def wrap(fn):
        def run(*args, **kwargs):
            identifier = request.headers.get("X-Module-F-Operation")
            try:
                if request.method != 'POST':
                    if not identifier:
                        return fn(*args, **kwargs)
                token = operation(identifier)
                with _LOCK:
                    if request.method == 'POST' and token.active:
                        raise SessionConflict('같은 작업의 요청이 이미 진행 중입니다. 완료 후 다시 실행하세요.')
                    token.reserve()
                with token.scope(reserved=True):
                    # Bodies are read inside the scope so uploads can be stopped
                    # before parsing/starting a worker, including late arrivals.
                    if request.method == "POST":
                        body = (request.get_json(silent=True) or {}) if request.is_json else request.form
                        if isinstance(body, dict) or hasattr(body, "get"):
                            sid = body.get("sid")
                            if sid:
                                from routes.module_f.jobs import _sess
                                try:
                                    from routes.module_f.session_policy import snapshot_fields, background_snapshot_fields
                                    sess = _sess(sid)
                                    fields = snapshot_fields(request.path,sess)
                                    token.claim(sess, draft=fields is not None,
                                        wait_readers=2 if request.path in ('/api/module-f/sub/graph',
                                            '/api/module-f/system/extract','/api/module-f/machineroom/extract') else 0)
                                    if fields is not None:
                                        token.draft(sess, fields)
                                    else:
                                        token.snapshot(sess, background_snapshot_fields(request.path))
                                except ValueError:
                                    pass  # Original endpoint returns its usual 410.
                    checkpoint()
                    result = fn(*args, **kwargs)
                    if token.workspaces:
                        response = app.make_response(result)
                        payload = response.get_json(silent=True) if response.is_json else None
                        if response.status_code >= 400 or (isinstance(payload, dict) and payload.get('ok') is False):
                            token.failed = True
                        return response
                    return result
            except OperationCancelled:
                return jsonify(ok=False, cancelled=True, message="작업을 중지했습니다."), 499
            except ValueError as exc:
                if identifier is not None and not re.fullmatch(r"[a-zA-Z0-9_-]{16,80}", identifier):
                    return jsonify(ok=False, message=str(exc)), 400
                raise
            except SessionConflict as exc:
                return jsonify(ok=False, message=str(exc)), 409

        @functools.wraps(fn)
        def guarded(*args, **kwargs):
            try:
                with reading():
                    return run(*args, **kwargs)
            except SessionConflict as exc:
                return jsonify(ok=False, message=str(exc)), 409
        return guarded

    for rule in list(app.url_map.iter_rules()):
        if endpoint_prefix and not rule.endpoint.startswith(endpoint_prefix):
            continue
        if rule.rule.startswith("/api/module-f/") and rule.rule not in _CONTROL_PATHS:
            fn = app.view_functions[rule.endpoint]
            if not getattr(fn, "_f_cancellable", False):
                wrapped = wrap(fn)
                wrapped._f_cancellable = True
                app.view_functions[rule.endpoint] = wrapped
