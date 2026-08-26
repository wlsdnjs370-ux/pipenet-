# -*- coding: utf-8 -*-
"""세션과 무거운 작업 — 진행 표시는 «실제로 찍힌 줄» 로만 한다."""
from __future__ import annotations

import sys
import threading
import time
import traceback
import uuid

from routes.module_f.common import LOG_TAIL, SESSION_TTL_SECONDS
from routes.module_f.slots import _slot_blank, _slot_init

_SESSIONS: dict[str, dict] = {}
_SESSIONS_LOCK = threading.Lock()
# 무거운 단계(도면 파싱·망 구성·평면 그래프)는 한 번에 하나만 돈다.
# docs/import 캐시와 stdout 을 공유하므로 겹치면 로그가 섞이고 캐시가 깨진다.
_HEAVY_LOCK = threading.Lock()

class _Tee:
    """파이프라인이 print 로 뱉는 단계 문구를 잡아 화면 진행표시로 쓴다.

    지어낸 퍼센트를 그리지 않기 위해서다 — 실제로 찍힌 줄만 보여준다.
    서버 로그도 그대로 유지해야 하므로 원본으로도 흘려보낸다.

    sys.stdout 은 프로세스 전역이라, 잡이 도는 동안 다른 요청이 찍은 줄까지
    이 잡의 진행표시로 새어 들어간다. 그래서 **작업 스레드가 쓴 줄만** 담는다.
    """

    def __init__(self, real, sink, owner):
        self._real = real
        self._sink = sink
        self._owner = owner
        self._buf = ""

    def write(self, s):
        if self._real is not None:
            try:
                self._real.write(s)
            except Exception:  # noqa: BLE001 — 로그 실패가 작업을 죽이면 안 된다
                pass
        if threading.current_thread() is not self._owner:
            return len(s)
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            line = line.strip()
            if line:
                self._sink(line)
        return len(s)

    def flush(self):
        if self._real is not None:
            try:
                self._real.flush()
            except Exception:  # noqa: BLE001
                pass

    def isatty(self):
        return False


# ─────────────────────────────────────────────────────────── 세션


# ─────────────────────────────────────────────────────────── 세션
def _sweep() -> None:
    now = time.time()
    with _SESSIONS_LOCK:
        dead = [k for k, s in _SESSIONS.items()
                if now - s.get("touched", 0) > SESSION_TTL_SECONDS]
        for k in dead:
            _SESSIONS.pop(k, None)


def _new_session(**kw) -> dict:
    """세션 하나 = 도면 슬롯 세 칸(S650). `slot=` 으로 첫 활성 슬롯을 고른다.

    평면 dict 는 그대로 둔다 — 그 내용이 곧 활성 슬롯의 도면 상태다(slots.py).
    """
    _sweep()
    sid = uuid.uuid4().hex[:16]
    sess = {
        "id": sid, "created": time.time(), "touched": time.time(),
        "job": None, "log": [],
    }
    sess.update(_slot_blank())
    _slot_init(sess, kw.pop("slot", "plan"))
    sess.update(kw)
    with _SESSIONS_LOCK:
        _SESSIONS[sid] = sess
    return sess


def _sess(sid: str) -> dict:
    with _SESSIONS_LOCK:
        found = _SESSIONS.get(str(sid or ""))
    if found is None:
        raise ValueError("작업이 만료되었습니다. 도면을 다시 여세요.")
    found["touched"] = time.time()
    return found


def _run_job(sess: dict, phase: str, fn) -> dict:
    """무거운 단계 하나를 백그라운드로 돌린다. 진행은 실제 출력 줄로만 보고."""
    job = {"state": "run", "phase": phase, "started": time.time(),
           "ended": None, "error": None, "result": None}
    sess["job"] = job
    sess["log"] = []

    def sink(line: str) -> None:
        log = sess["log"]
        log.append(line)
        if len(log) > 400:
            del log[:200]

    def worker() -> None:
        me = threading.current_thread()
        with _HEAVY_LOCK:
            old_out, old_err = sys.stdout, sys.stderr
            sys.stdout = _Tee(old_out, sink, me)
            sys.stderr = _Tee(old_err, sink, me)
            try:
                job["result"] = fn()
                job["state"] = "done"
            # ★BaseException 까지다. 엔진은 CLI 태생이라 실패를 SystemExit 로
            #   던지는 곳이 있다(실측: 원본 DXF 없는 키 reopen →
            #   `raise SystemExit("DXF를 못 찾음: apt")`). Exception 만 잡으면
            #   워커가 소리 없이 죽고 잡이 영원히 «run» 으로 남아, 사용자는
            #   멈춘 진행바만 보게 된다 — 실패가 있으면 실패라고 말해야 한다.
            except BaseException as exc:  # noqa: BLE001 — 무엇이 나든 화면에 알린다
                job["state"] = "error"
                job["error"] = f"{type(exc).__name__}: {exc}"
                sink("!! " + job["error"])
                for ln in traceback.format_exc().splitlines()[-6:]:
                    sink("   " + ln)
            finally:
                sys.stdout, sys.stderr = old_out, old_err
                job["ended"] = time.time()
                sess["touched"] = time.time()

    threading.Thread(target=worker, daemon=True).start()
    return job


def _job_view(sess: dict) -> dict:
    job = sess.get("job")
    if job is None:
        return {"state": "idle", "phase": "", "elapsed": 0.0, "lines": []}
    end = job["ended"] or time.time()
    return {
        "state": job["state"], "phase": job["phase"],
        "elapsed": round(end - job["started"], 1),
        "error": job["error"],
        "lines": sess["log"][-LOG_TAIL:],
        "queued": _HEAVY_LOCK.locked() and job["state"] == "run",
    }


# ─────────────────────────────────────────────────────────── 도형 직렬화


def _job_running(sess: dict) -> bool:
    """이 세션에서 무거운 작업이 아직 돌고 있나.

    가림막이 캔버스만 덮어 옆 패널 단추는 작업 중에도 눌린다. 실측 — 자동
    이음을 두 번 밀어넣으면 **낡은 후보로 한 번 더 붙어** 다리가 겹치고
    (간선 +49 → +81) 되돌리기 한 번으로 원상복구가 안 되며 보고 수치도
    두 번째 것만 남아 거짓말이 된다. 화면도 고치지만 진짜 방벽은 여기다.
    """
    job = sess.get("job")
    return bool(job) and job.get("state") == "run"
