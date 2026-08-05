# -*- coding: utf-8 -*-
"""모듈 C 설계 자동화 API — 지시서 §11.

`register(app, ...)` 패턴. 기존 `/api/remote30/*` 는 건드리지 않으며, 이 모듈은
`DESIGN_WORKBENCH_ENABLED` 가 꺼져 있으면 라우트를 아예 등록하지 않는다(§C.5).

PR-2 범위는 세션과 게이트뿐이다. C2 이후 엔드포인트는 **게이트 강제만** 걸어
자리를 잡아 두고, 게이트를 통과한 뒤에는 아직 구현되지 않았음을 501 로 밝힌다.
통과한 것처럼 200 을 돌려주면 화면이 진행된 것으로 오해한다.
"""
from __future__ import annotations

from flask import jsonify, request

from core.design import gate as G
from core.design import session as S
from core.design.schema import BuildingDraft

# 게이트 뒤에 서는 단계들. 지시서 §4.1 — `/api/design/c2/*` 이후 전부.
_GATED_STAGES = (
    ("/api/design/c2/constraints", "POST", "C2"),
    ("/api/design/c2b/esfr", "POST", "C2B"),
    ("/api/design/c3/valves", "POST", "C3"),
    ("/api/design/c3/zones", "POST", "C3-zones"),
    ("/api/design/c4/heads", "POST", "C4"),
    ("/api/design/c5/route", "POST", "C5"),
    ("/api/design/emit", "POST", "EMIT"),
    ("/api/design/checks/<sid>", "GET", "CHECKS"),
)


def _fail(code: str, message: str, status: int, **extra):
    return jsonify({"ok": False, "code": code, "message": message, **extra}), status


def register(app, *, DESIGN_SESSION_DIR, enabled: bool = False):
    if not enabled:
        return

    DESIGN_SESSION_DIR.mkdir(parents=True, exist_ok=True)

    def _open(sid: str):
        """반환 `(세션, None)` 또는 `(None, 오류응답)`."""
        try:
            return S.DesignSession.open(DESIGN_SESSION_DIR, sid), None
        except S.SessionNotFound:
            return None, _fail("SESSION_NOT_FOUND", "세션을 찾을 수 없습니다.", 404)

    def _load_draft(sess):
        """반환 `(draft, version, None)` 또는 `(None, 0, 오류응답)`."""
        try:
            raw, version = sess.read("building.json")
        except S.ArtifactNotFound:
            return None, 0, _fail(
                "C1_NOT_DONE", "C1 인식이 끝나지 않아 확정할 대상이 없습니다.", 409)
        return BuildingDraft.from_dict(raw), version, None

    # ── 세션 ────────────────────────────────────────────────────────────
    @app.post("/api/design/session")
    def design_session_create():
        body = request.get_json(silent=True) or {}
        operator = (body.get("operator") or "").strip() or None
        sess = S.DesignSession.create(DESIGN_SESSION_DIR, operator=operator)
        return jsonify({"ok": True, "session_id": sess.sid}), 201

    @app.get("/api/design/session/<sid>")
    def design_session_status(sid):
        sess, err = _open(sid)
        if err:
            return err
        return jsonify({"ok": True, **sess.status()})

    # ── GATE ────────────────────────────────────────────────────────────
    @app.get("/api/design/c1/gate_items/<sid>")
    def design_gate_items(sid):
        sess, err = _open(sid)
        if err:
            return err
        draft, _version, err = _load_draft(sess)
        if err:
            return err
        return jsonify(G.gate_items(draft))

    @app.post("/api/design/gate/confirm")
    def design_gate_confirm():
        body = request.get_json(silent=True) or {}
        sess, err = _open(str(body.get("session_id") or ""))
        if err:
            return err
        draft, version, err = _load_draft(sess)
        if err:
            return err

        operator = (body.get("operator") or "").strip() or None
        if_version = body.get("if_version")
        try:
            changes = G.apply_values(draft, body.get("values") or {})
        except (ValueError, TypeError) as exc:
            return _fail("INVALID_VALUE", str(exc), 400)

        defaults = G.apply_defaults(draft)
        edits = list(body.get("edits") or [])
        draft.gate.edits.extend(edits)
        draft.gate.operator = operator
        draft.gate.unresolved = G.unresolved(draft)
        draft.gate.passed = not draft.gate.unresolved
        draft.gate.passed_at = S.now_iso() if draft.gate.passed else None

        try:
            new_version = sess.write(
                "building.json", draft.to_dict(),
                if_version=int(if_version) if if_version is not None else version)
        except S.VersionConflict as conflict:
            return _fail("VERSION_CONFLICT",
                         "다른 곳에서 먼저 저장했습니다. 현재 내용을 확인하세요.", 409,
                         current_version=conflict.current, current=conflict.data)

        actor = operator or "unknown"
        for change in changes:
            event = ("use_override"
                     if change["field"] == "use" and change.get("suggested") else "value_confirmed")
            sess.audit(actor, "GATE", event, change)
        for applied in defaults:
            sess.audit("system", "GATE", "default_applied", applied)
        for edit in edits:
            sess.audit(actor, "GATE", "room_edit", edit)

        if not draft.gate.passed:
            sess.audit(actor, "GATE", "confirm_incomplete",
                       {"unresolved": len(draft.gate.unresolved)})
            return jsonify({
                "ok": False, "code": "GATE_INCOMPLETE",
                "message": "확정되지 않은 필수 항목이 남아 있습니다.",
                "unresolved": draft.gate.unresolved, "version": new_version,
            }), 422

        sess.update_meta(stage="c2", gate_passed=True, operator=operator)
        sess.audit(actor, "GATE", "passed",
                   {"rooms": len(draft.rooms), "edits": len(draft.gate.edits)})
        return jsonify({"ok": True, "passed": True, "version": new_version,
                        "passed_at": draft.gate.passed_at})

    # ── 게이트 뒤 단계 (자리만) ─────────────────────────────────────────
    def _make_gated(stage: str):
        def view(sid=None):
            session_id = sid or str((request.get_json(silent=True) or {}).get("session_id") or "")
            sess, err = _open(session_id)
            if err:
                return err
            draft, _version, err = _load_draft(sess)
            if err:
                return err
            try:
                G.require_gate(draft)
            except G.GateNotPassed as exc:
                return _fail("GATE_NOT_PASSED", str(exc), 409, unresolved=exc.unresolved)
            return _fail("STAGE_NOT_IMPLEMENTED",
                         f"{stage} 단계는 아직 구현되지 않았습니다.", 501)
        return view

    for rule, method, stage in _GATED_STAGES:
        app.add_url_rule(rule, endpoint=f"design_stage_{stage.lower().replace('-', '_')}",
                         view_func=_make_gated(stage), methods=[method])
