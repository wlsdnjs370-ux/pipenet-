# -*- coding: utf-8 -*-
"""인증 도메인 라우트 — 로그인/로그아웃.

`대조 서버.py` 에서 `register(app, login_password=...)` 로 등록한다.
엔드포인트명(login_page/login_submit/logout)은 접두사 없이 보존된다.
"""
from __future__ import annotations

from flask import (make_response, redirect, render_template, request,
                   session, url_for)


def register(app, *, login_password: str) -> None:
    @app.get("/login")
    def login_page():
        if session.get("authed"):
            nxt = request.args.get("next", "/")
            return redirect(nxt or "/")
        response = make_response(render_template(
            "login.html", error=None, next_path=request.args.get("next", "/")))
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.post("/login")
    def login_submit():
        pw = (request.form.get("password") or "").strip()
        nxt = (request.form.get("next") or "/").strip() or "/"
        # safety — open redirect 방지 (외부 URL 금지)
        if not nxt.startswith("/") or nxt.startswith("//"):
            nxt = "/"
        if pw == login_password:
            session["authed"] = True
            session.permanent = True  # 세션 영구 (기본 31일)
            return redirect(nxt)
        response = make_response(render_template(
            "login.html", error="Incorrect password.", next_path=nxt))
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.get("/logout")
    def logout():
        session.pop("authed", None)
        return redirect(url_for("login_page"))
