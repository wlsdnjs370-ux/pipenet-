# -*- coding: utf-8 -*-
"""세션과 무거운 작업 — 진행 표시는 «실제로 찍힌 줄» 로만 한다."""
from __future__ import annotations

import functools
import sys
import threading
import time
import traceback
import uuid

from flask import request

from routes.module_f.common import LOG_TAIL, SESSION_TTL_SECONDS, _fail
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
# 한 세션이 붙드는 것이 가볍지 않다 — 슬롯마다 도면 하나치 도형 목록이 앉는다
# (실측: LH306동 21,715개 · B1F 는 그 몇 배). 만료로만 회수하면 아무도 새 도면을
# 안 여는 동안 계속 쌓인다.
MAX_SESSIONS = 24
_last_sweep = 0.0
SWEEP_EVERY_SECONDS = 60.0


def _sweep(force: bool = False) -> None:
    """만료 세션 회수 + 수 상한. 자주 불리므로 1분에 한 번만 실제로 돈다."""
    global _last_sweep
    now = time.time()
    if not force and (now - _last_sweep) < SWEEP_EVERY_SECONDS:
        return
    _last_sweep = now
    with _SESSIONS_LOCK:
        # ★도는 작업은 만료로도 걷지 않는다. 아래 수 상한 분기는 이미 그렇게
        #   하는데 여기만 안 봐줬다 — 같은 함수 안에서 규칙이 갈려 있었다.
        #   걷어 버리면 워커는 계속 돌지만 그 결과를 받을 세션이 없어, 일이
        #   «끝난 뒤에» 화면은 「작업이 만료되었습니다」만 본다.
        #   (SSE 로 보는 동안 touched 가 안 갱신되던 것도 같이 고쳤다 —
        #    api_open 의 진행 스트림 참조.)
        dead = [k for k, s in _SESSIONS.items()
                if now - s.get("touched", 0) > SESSION_TTL_SECONDS
                and not ((s.get("job") or {}).get("state") == "run")]
        for k in dead:
            _SESSIONS.pop(k, None)
        # 그래도 넘치면 오래 안 만진 것부터 — 단, 도는 작업은 건드리지 않는다.
        if len(_SESSIONS) > MAX_SESSIONS:
            idle = sorted(
                (s for s in _SESSIONS.values()
                 if not ((s.get("job") or {}).get("state") == "run")),
                key=lambda s: s.get("touched", 0))
            for s in idle[:len(_SESSIONS) - MAX_SESSIONS]:
                _SESSIONS.pop(s.get("id"), None)


def _new_session(**kw) -> dict:
    """세션 하나 = 도면 슬롯 세 칸(S650). `slot=` 으로 첫 활성 슬롯을 고른다.

    평면 dict 는 그대로 둔다 — 그 내용이 곧 활성 슬롯의 도면 상태다(slots.py).
    """
    _sweep(force=True)          # 새 세션을 만들 때는 반드시 한 번 훑는다
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
    # 회수를 «새 도면 열 때» 에만 걸면, 아무도 안 여는 동안 만료된 세션이
    # 도형 목록을 그대로 붙들고 있다. 여기서도 훑는다(1분 스로틀).
    _sweep()
    with _SESSIONS_LOCK:
        found = _SESSIONS.get(str(sid or ""))
    if found is None:
        raise ValueError("작업이 만료되었습니다. 도면을 다시 여세요.")
    found["touched"] = time.time()
    return found


def route_session(resolve=None, *, post: bool = False, why_code: int = 409):
    """라우트 앞머리 — 요청 꺼내기 · 세션 찾기 · 실패 응답을 한 자리에.

    라우트 60개가 저마다 이 네댓 줄을 손으로 적고 있었다(실측 43곳 · 188줄):

        body = request.get_json(silent=True) or {}
        try:
            sess, why = _need_auto(body)
        except ValueError as exc:
            return _fail(str(exc), 410)
        if why:
            return _fail(why, 409)

    베끼는 자리는 «빠뜨릴 수 있는 자리» 다. 한곳에 모아 두면 새 라우트가 가드를
    잊을 수 없고, 만료 응답(410)·충돌 응답(409)의 뜻이 한 군데서만 정해진다.

    핸들러는 `(sess, body)` 를 받는다 — body 는 POST 면 JSON dict, GET 이면
    `request.args`. 둘 다 `.get()` 이 되므로 쓰는 쪽은 같다.

    resolve: 안 주면 `_sess(body["sid"])`. 주면 «(sess, 사유)» 를 돌려주는
        함수로 보고(`_need_auto` 꼴), 사유가 있으면 막는다. 세션만 돌려주는
        함수도 그대로 받는다.
    why_code: 사유가 있을 때의 상태 코드. 기본 409(«지금은 안 된다»)지만,
        옮겨 오기 전에 400 으로 답하던 자리는 400 을 그대로 준다 — 리팩터가
        화면이 보는 코드를 바꾸면 안 된다.

    ★`@app.get(...)` 아래에 붙인다. Flask 는 함수 이름을 엔드포인트로 쓰므로
      `functools.wraps` 로 이름을 지켜야 한다.
    """
    def deco(fn):
        @functools.wraps(fn)
        def wrapper(*a, **kw):
            body = (request.get_json(silent=True) or {}) if post else request.args
            try:
                got = resolve(body) if resolve is not None else _sess(body.get("sid"))
            except ValueError as exc:
                return _fail(str(exc), 410)
            if isinstance(got, tuple):
                sess, why = got
                if why:
                    # 사유는 문장 하나, 또는 «(문장, 코드)» — 라우트마다 코드가
                    # 다른 자리를 옮겨 올 때 그 코드를 그대로 지키기 위해서다.
                    if isinstance(why, tuple):
                        return _fail(why[0], why[1])
                    return _fail(why, why_code)
            else:
                sess = got
            return fn(sess, body, *a, **kw)
        return wrapper
    return deco


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
        return {"state": "idle", "phase": "", "elapsed": 0.0, "lines": [],
                "world_ready": sess.get("world") is not None}
    end = job["ended"] or time.time()
    return {
        "state": job["state"], "phase": job["phase"],
        "elapsed": round(end - job["started"], 1),
        "error": job["error"],
        "lines": sess["log"][-LOG_TAIL:],
        "queued": _HEAVY_LOCK.locked() and job["state"] == "run",
        # ★잡이 끝나기 «전에» 도면을 그릴 수 있는가. `_open_job` 은 찍기판을
        #   세우자마자 sess["world"] 를 앉히고 그 뒤에 정찰을 덤으로 돌린다
        #   — 도면은 그 사이 내내 준비되어 있다. 그 사실을 화면에 말해 주지
        #   않으면 화면은 잡이 다 끝날 때까지 빈 캔버스를 보여 준다.
        #   실측(B1F 110.6MB · 처음 여는 도면):
        #       찍기 6.8s + 도형 2.2s = 9.0s   ← 여기서 이미 그릴 수 있다
        #       정찰(덤)             +33.2s
        #   즉 42초 중 33초가 «이미 있는 도면» 을 안 그린 채 흘렀다.
        "world_ready": sess.get("world") is not None,
    }


def _job_running(sess: dict) -> bool:
    """이 세션에서 무거운 작업이 아직 돌고 있나.

    가림막이 캔버스만 덮어 옆 패널 단추는 작업 중에도 눌린다. 실측 — 자동
    이음을 두 번 밀어넣으면 **낡은 후보로 한 번 더 붙어** 다리가 겹치고
    (간선 +49 → +81) 되돌리기 한 번으로 원상복구가 안 되며 보고 수치도
    두 번째 것만 남아 거짓말이 된다. 화면도 고치지만 진짜 방벽은 여기다.
    """
    job = sess.get("job")
    return bool(job) and job.get("state") == "run"
