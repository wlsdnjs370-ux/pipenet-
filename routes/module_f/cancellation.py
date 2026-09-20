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


class OperationCancelled(BaseException):
    """Control flow, deliberately not swallowed by engine Exception handlers."""


_LOCAL = threading.local()
_LOCK = threading.RLock()
_OPERATIONS: dict[str, "Operation"] = {}
_ROOT = str(Path(__file__).resolve().parents[2]).lower().replace("\\", "/") + "/"
_MONITOR_ID = None
_MONITORING = getattr(sys, "monitoring", None)
_PROTECTED_MODULES = ("/pick/io.py", "/edit/io.py", "/pipeline/handoff.py",
                      "/pipeline/disp_cache.py")
_INFRA_MODULES = ("/module_f/cancellation.py", "/module_f/jobs.py")
_CONTROL_PATHS = {"/api/module-f/job", "/api/module-f/job/stream",
                  "/api/module-f/job/cancel", "/api/module-f/job/cancel-state"}


def _interruptible(frame) -> bool:
    """Keep persistence and infrastructure out of computational checkpoints."""
    path = frame.f_code.co_filename.lower().replace("\\", "/")
    if not (path.startswith(_ROOT) or "/ezdxf/" in path):
        return False
    if path.endswith(_INFRA_MODULES):
        return False
    while frame is not None:
        filename = frame.f_code.co_filename.lower().replace("\\", "/")
        name = frame.f_code.co_name.lower().lstrip("_")
        if filename.endswith(_PROTECTED_MODULES):
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
        frame = frame.f_back
    return True


def _line(_code, _line_number) -> None:
    op = getattr(_LOCAL, "operation", None)
    if (op is not None and op.cancelled and not getattr(_LOCAL, "delivered", False)
            and _interruptible(sys._getframe(1))):
        _LOCAL.delivered = True  # Do not interrupt exception/finally cleanup again.
        raise OperationCancelled("작업을 중지했습니다.")


def _trace(frame, event, arg):
    # Python 3.11 fallback. Current local server uses 3.13 monitoring instead.
    if event == "line":
        op = getattr(_LOCAL, "operation", None)
        if op is not None and op.cancelled and not getattr(_LOCAL, "delivered", False):
            if _interruptible(frame):
                _LOCAL.delivered = True
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


def _session_copy(sess: dict) -> dict:
    """Copy editable state; retain immutable DXF geometry and spatial indexes."""
    memo = {}
    states = [sess, *(sess.get("slots") or {}).values()]
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
    return copy.deepcopy({k: v for k, v in sess.items()
                          if k not in {"job", "log", "touched"}}, memo)


class Operation:
    """One UI Working operation, possibly spanning a request and a worker."""

    def __init__(self, identifier: str):
        self.id = identifier
        self.cancelled = False
        self.active = 0
        self.touched = time.monotonic()
        self.snapshots: dict[int, tuple[dict, dict]] = {}
        self.rollback_callbacks: list = []
        self.rollback_errors: list[str] = []

    def reserve(self) -> None:
        with _LOCK:
            if self.cancelled:
                raise OperationCancelled("작업을 중지했습니다.")
            self.active += 1
            self.touched = time.monotonic()
            _monitor_refresh()

    def snapshot(self, sess: dict) -> None:
        if id(sess) not in self.snapshots:
            self.snapshots[id(sess)] = (sess, _session_copy(sess))

    def release(self) -> None:
        with _LOCK:
            self.active -= 1
            self.touched = time.monotonic()
            if self.active == 0:
                if self.cancelled:
                    for rollback in reversed(self.rollback_callbacks):
                        try:
                            rollback()
                        except OSError as exc:
                            self.rollback_errors.append(str(exc))
                    for sess, before in self.snapshots.values():
                        runtime = {k: sess[k] for k in ("job", "log", "touched") if k in sess}
                        sess.update(before)
                        for key in set(sess) - set(before) - set(runtime):
                            sess.pop(key, None)
                        sess.update(runtime)
                self.snapshots.clear()
                self.rollback_callbacks.clear()
            _monitor_refresh()

    @contextmanager
    def scope(self, *, reserved: bool = False):
        """Bind cancellation to one thread; native work is never killed."""
        if not reserved:
            self.reserve()
        previous = current_operation()
        previous_delivered = getattr(_LOCAL, "delivered", False)
        previous_trace = sys.gettrace()
        _LOCAL.operation, _LOCAL.delivered = self, False
        use_trace = _MONITOR_ID is None
        if use_trace:
            sys.settrace(_trace)
        try:
            checkpoint()
            yield self
            checkpoint()
        finally:
            _LOCAL.operation, _LOCAL.delivered = previous, previous_delivered
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


def install(app) -> None:
    """Wrap every F request carrying the UI operation id, including uploads."""
    @app.post("/api/module-f/job/cancel")
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

    @app.get("/api/module-f/job/cancel-state")
    def module_f_cancel_state():
        identifier = request.args.get("operation", "")
        try:
            return jsonify(operation(identifier).status())
        except ValueError as exc:
            return jsonify(ok=False, message=str(exc)), 400

    def wrap(fn):
        @functools.wraps(fn)
        def guarded(*args, **kwargs):
            identifier = request.headers.get("X-Module-F-Operation")
            if not identifier:
                return fn(*args, **kwargs)
            try:
                token = operation(identifier)
                with token.scope():
                    # Bodies are read inside the scope so uploads can be stopped
                    # before parsing/starting a worker, including late arrivals.
                    if request.method == "POST":
                        body = (request.get_json(silent=True) or {}) if request.is_json else request.form
                        if isinstance(body, dict) or hasattr(body, "get"):
                            sid = body.get("sid")
                            if sid:
                                from routes.module_f.jobs import _sess
                                try:
                                    token.snapshot(_sess(sid))
                                except ValueError:
                                    pass  # Original endpoint returns its usual 410.
                    checkpoint()
                    return fn(*args, **kwargs)
            except OperationCancelled:
                return jsonify(ok=False, cancelled=True, message="작업을 중지했습니다."), 499
            except ValueError as exc:
                if not re.fullmatch(r"[a-zA-Z0-9_-]{16,80}", identifier):
                    return jsonify(ok=False, message=str(exc)), 400
                raise
        return guarded

    for rule in list(app.url_map.iter_rules()):
        if rule.rule.startswith("/api/module-f/") and rule.rule not in _CONTROL_PATHS:
            fn = app.view_functions[rule.endpoint]
            if not getattr(fn, "_f_cancellable", False):
                wrapped = wrap(fn)
                wrapped._f_cancellable = True
                app.view_functions[rule.endpoint] = wrapped
