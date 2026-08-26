# -*- coding: utf-8 -*-
"""[H-0] 모듈 F 라우트 — 도면 슬롯(특허 S650).

세 라우트뿐이다. 슬롯을 **열고 · 바꾸고 · 들여다본다.** 도면을 실제로 여는 일은
`api_open._open_job` 을 그대로 쓴다 — 특허 S650 이 «같은 절차를 반복 적용» 하라고
하므로, 계통도·기계실이 평면도와 다른 제1국면을 밟으면 그 자체가 오구현이다.

계통도·기계실의 **추출**(A 엔진 접합)은 H-2 · H-3 의 일이다. 여기서는 슬롯이
서로를 덮지 않는다는 계약까지만 세운다.
"""
from __future__ import annotations

import os

from flask import jsonify, request

from routes.module_f.api_open import _open_job
from routes.module_f.common import _boot, _fail
from routes.module_f.jobs import _job_running, _new_session, _run_job, _sess
from routes.module_f.slots import (
    SLOT_LABELS, _check_slot_kind, _slot_state, _slot_switch)


def register(app, *, _save_upload):
    @app.get("/api/module-f/slot/state")
    def module_f_slot_state():
        """세 슬롯의 진행 한 장 — S650 이 «남은 도면이 있나» 를 묻는 자리."""
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        out = _slot_state(sess)
        out["ok"] = True
        return jsonify(out)

    @app.post("/api/module-f/slot/switch")
    def module_f_slot_switch():
        """활성 슬롯을 바꾼다. 작업이 도는 중에는 거절한다.

        ★잡이 도는 중에 슬롯을 바꾸면 워커가 **다른 슬롯의 평면 dict** 에 결과를
          쓴다. `_open_job` 의 클로저가 붙잡은 것은 세션이지 슬롯이 아니기
          때문이다 — 계통도를 읽던 잡이 평면도의 찍기 상태를 덮어쓴다.
        """
        body = request.get_json(silent=True) or {}
        try:
            sess = _sess(body.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if _job_running(sess):
            return _fail("작업이 끝난 뒤에 도면을 바꿀 수 있습니다.", 409)
        try:
            kind = _slot_switch(sess, body.get("kind"))
        except ValueError as exc:
            return _fail(str(exc))
        out = _slot_state(sess)
        out["ok"] = True
        out["switched"] = kind
        return jsonify(out)

    @app.post("/api/module-f/slot/open")
    def module_f_slot_open():
        """도면 종류를 지정해 DXF 를 연다(S650 의 회귀 한 바퀴).

        `sid` 가 있으면 그 세션의 해당 슬롯으로, 없으면 새 세션을 그 슬롯으로
        시작한다. 열기 자체는 평면도와 **같은 잡**이다.
        """
        try:
            kind = _check_slot_kind(request.form.get("kind"))
        except ValueError as exc:
            return _fail(str(exc))

        sid = (request.form.get("sid") or "").strip()
        sess = None
        if sid:
            try:
                sess = _sess(sid)
            except ValueError as exc:
                return _fail(str(exc), 410)
            if _job_running(sess):
                return _fail("작업이 끝난 뒤에 도면을 열 수 있습니다.", 409)

        try:
            _boot()
            dxf = _save_upload("dxf_file", {".dxf"}, required=True)
        except ValueError as exc:
            return _fail(str(exc))
        except Exception as exc:  # noqa: BLE001
            return _fail(f"도면을 저장하지 못했습니다: {exc}", 500)

        if sess is None:
            sess = _new_session(slot=kind, dxf=str(dxf))
        else:
            _slot_switch(sess, kind)
            sess["dxf"] = str(dxf)

        _run_job(sess, f"{SLOT_LABELS[kind]} 읽기", _open_job(sess, dxf))
        return jsonify({"ok": True, "sid": sess["id"], "kind": kind,
                        "filename": os.path.basename(str(dxf))})
