# -*- coding: utf-8 -*-
"""인증 도메인 라우트 — 로그인/로그아웃.

`대조 서버.py` 에서 `register(app, login_password=...)` 로 등록한다.
엔드포인트명(login_page/login_submit/logout)은 접두사 없이 보존된다.
"""
from __future__ import annotations

import ipaddress
import secrets
import time
from threading import Lock

from flask import (make_response, redirect, render_template, request,
                   session, url_for)

# Brute-force 방어 — 4자리 PIN 은 무제한 시도면 1만 회로 뚫린다.
# 단순 in-memory dict — waitress 단일 프로세스 + 16 thread 환경에 적합.
_FAIL_WINDOW_S = 15 * 60
_FAIL_THRESHOLD = 5
_LOCK_S = 15 * 60
_login_state: dict[str, dict] = {}   # ip → {"fails": [ts, ...], "lock_until": ts}
_login_lock = Lock()


def _locked_until(ip: str) -> float:
    with _login_lock:
        st = _login_state.get(ip)
        return float(st.get("lock_until", 0.0)) if st else 0.0


def _record_failure(ip: str) -> tuple[int, float]:
    """실패 1회 기록 → (윈도우 내 실패횟수, lock_until_ts)."""
    now = time.time()
    with _login_lock:
        st = _login_state.setdefault(ip, {"fails": [], "lock_until": 0.0})
        st["fails"] = [t for t in st["fails"] if now - t < _FAIL_WINDOW_S]
        st["fails"].append(now)
        if len(st["fails"]) >= _FAIL_THRESHOLD:
            st["lock_until"] = now + _LOCK_S
        return len(st["fails"]), st["lock_until"]


def _hint_allowed() -> bool:
    """비밀번호 힌트를 띄워도 되는 요청인가 — 루프백/사설망에서만.

    이 페이지는 cloudflared 로 인터넷에 노출돼 있어, 힌트를 무조건 렌더하면
    로그인 게이트가 사실상 무력화된다. ProxyFix 가 X-Forwarded-For 를 반영하므로
    외부 접속은 공인 IP 로 잡힌다.
    """
    try:
        ip = ipaddress.ip_address(request.remote_addr or "")
    except ValueError:
        return False
    return ip.is_loopback or ip.is_private


def register(app, *, login_password: str) -> None:
    @app.get("/login")
    def login_page():
        if session.get("authed"):
            nxt = request.args.get("next", "/")
            return redirect(nxt or "/")
        response = make_response(render_template(
            "login.html", error=None, next_path=request.args.get("next", "/"),
            password_hint=login_password if _hint_allowed() else None))
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.post("/login")
    def login_submit():
        pw = (request.form.get("password") or "").strip()
        nxt = (request.form.get("next") or "/").strip() or "/"
        # safety — open redirect 방지 (외부 URL 금지)
        if not nxt.startswith("/") or nxt.startswith("//"):
            nxt = "/"

        def _deny(msg, status=200):
            response = make_response(render_template(
                "login.html", error=msg, next_path=nxt,
                password_hint=login_password if _hint_allowed() else None))
            response.headers["Cache-Control"] = "no-store"
            return response, status

        # ProxyFix 적용 후라 remote_addr 이 진짜 클라이언트 IP.
        ip = request.remote_addr or "unknown"
        now = time.time()
        lock_until = _locked_until(ip)
        if lock_until > now:
            wait = int(lock_until - now)
            return _deny(f"로그인 시도 횟수 초과. {wait // 60}분 {wait % 60}초 후 "
                         f"다시 시도하세요.", 429)
        if secrets.compare_digest(pw, login_password):  # 상수 시간 — timing attack 방지
            with _login_lock:
                _login_state.pop(ip, None)
            session["authed"] = True
            session.permanent = True  # 세션 영구 (기본 31일)
            return redirect(nxt)
        n_fail, locked = _record_failure(ip)
        if locked > now:
            return _deny(f"로그인 시도 {_FAIL_THRESHOLD}회 실패. "
                         f"{int(locked - now) // 60}분 동안 잠금됩니다.")
        return _deny(f"비밀번호가 틀립니다. (남은 시도 {_FAIL_THRESHOLD - n_fail}회)")

    @app.get("/logout")
    def logout():
        session.pop("authed", None)
        return redirect(url_for("login_page"))
