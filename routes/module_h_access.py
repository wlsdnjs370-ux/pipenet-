"""Read-only readiness for H's input circuit; never infer completion from visits.

Ready means material exists for integration, not hydraulic/code validation.
F retains its optional system/machine inputs and its original navigation.
"""
from __future__ import annotations

from flask import Flask, jsonify

from routes.module_f.jobs import _job_running, route_session
from routes.module_f.selection import _design_stale
from routes.module_f.slots import (
    SYSTEM_MAX, _slot_active, slot_kinds, slot_label, slot_role, system_kinds,
)


def access_state(sess: dict) -> dict:
    """Describe live and parked slots without switching or changing any slot."""
    active = _slot_active(sess)
    slots = []
    for kind in slot_kinds(sess):
        state = sess if kind == active else (sess.get("slots") or {}).get(kind, {})
        role = slot_role(kind)
        if role == "plan":
            tables = (state.get("design") or {}).get("tables")
            ready = bool(getattr(tables, "nodes", None) and getattr(tables, "pipes", None))
            stale = _design_stale(state) if ready else None
            ready = ready and not stale
            reason = "입력표 준비 완료" if ready else "속성 정의 · 입력표 준비 필요"
            if stale:
                reason = " · ".join(stale["why"])
        else:
            path = state.get("riser" if role == "system" else "machineroom") or {}
            ready = bool(path.get("nodes") and path.get("pipes"))
            reason = "추출 경로 준비 완료" if ready else "연결점 지정 · 경로 추출 필요"
        slots.append(dict(kind=kind, role=role, label=slot_label(sess, kind),
                          key=str(state.get("key") or ""), active=kind == active,
                          opened=bool(state.get("world") or state.get("entities")),
                          ready=ready, reason=reason, removable=kind.startswith("system") and kind != "system"))
    groups = {role: all(s["ready"] for s in slots if s["role"] == role)
              for role in ("plan", "system", "machineroom")}
    return dict(ok=True, active=active, slots=slots, groups=groups,
                can_merge=all(groups.values()) and not _job_running(sess),
                can_add_system=len(system_kinds(sess)) < SYSTEM_MAX)


def register(app: Flask) -> None:
    """Expose H-only readiness under the existing application auth gate."""
    @app.get("/api/module-h/access-state")
    @route_session()
    def module_h_access_state(sess: dict, body: dict):
        return jsonify(access_state(sess))
