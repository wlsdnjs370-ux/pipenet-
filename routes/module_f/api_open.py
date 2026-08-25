# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 페이지·설명 그림·열기/이어열기·진행·도면."""
from __future__ import annotations

import os
import time

from flask import jsonify, render_template, request, send_file

from routes.module_f.common import (
    DIAGRAMS, IMPORT_WORK_ROOT, _boot, _check_key, _fail)
from routes.module_f.jobs import _job_view, _new_session, _run_job, _sess
from routes.module_f.remote30 import _sheet_frames
from routes.module_f.views import _pick_state
from routes.module_f.world import _saved_keys, _world_payload


def register(app, *, _save_upload):
    @app.get("/module-f")
    def module_f_page():
        return render_template("module_f.html")

    @app.get("/api/module-f/diagram/<key>")
    def module_f_diagram(key):
        """설명 그림 — 모듈 E 가 대화상자에 띄우는 파일 그대로.

        키는 위 표에 있는 것만 받는다(경로가 아니라 키다 — 도면 폴더 아래
        아무 파일이나 내보내지 않는다).
        """
        name = DIAGRAMS.get(str(key))
        if name is None:
            return _fail(f"그런 그림이 없습니다: {key}", 404)
        path = IMPORT_WORK_ROOT / name
        if not path.is_file():
            return _fail(f"그림 파일이 없습니다: {name}", 404)
        return send_file(str(path), mimetype="image/png")

    # ─────────────────────────────────────────── 0. 열기
    @app.get("/api/module-f/saved")
    def module_f_saved():
        try:
            _boot()
            return jsonify({"ok": True, "items": _saved_keys()})
        except Exception as exc:  # noqa: BLE001
            return _fail(f"저장된 찍기 목록을 읽지 못했습니다: {exc}", 500)

    @app.post("/api/module-f/open")
    def module_f_open():
        """DXF 를 올려 찍기 세션을 연다. 파싱이 길어 잡으로 돌린다."""
        try:
            _boot()
            dxf = _save_upload("dxf_file", {".dxf"}, required=True)
        except ValueError as exc:
            return _fail(str(exc))
        except Exception as exc:  # noqa: BLE001
            return _fail(f"도면을 저장하지 못했습니다: {exc}", 500)

        sess = _new_session(dxf=str(dxf))

        def job():
            from services.cad_import.pick.session import PickSession
            t0 = time.perf_counter()
            print(f"[찍기] DXF 읽는 중 — {os.path.basename(str(dxf))}")
            ps = PickSession.open(str(dxf))
            _check_key(ps.key)   # 올린 파일 이름에서 딴 키도 같은 자를 통과시킨다
            # E 의 찍기판은 열자마자 armed 가 아니다 — Qt 대화상자도 "배관 선택"
            # 을 눌러야 클릭이 먹는다. 웹에서는 첫 할 일이 어차피 그것뿐이라
            # 같은 호출을 미리 해 둔다(엔진 상태는 단추를 누른 것과 동일).
            ps.select_pipe()
            sess["pick"] = ps
            sess["key"] = ps.key
            payload = _world_payload(ps.world)
            sess["world"] = payload
            print(f"[찍기] 완료 {time.perf_counter() - t0:.1f}s · "
                  f"선분 {payload['counts']['segs']} · "
                  f"원 {payload['counts']['circles']} · "
                  f"호 {payload['counts']['arcs']}")
            return {"key": ps.key}

        _run_job(sess, "도면 읽기", job)
        return jsonify({"ok": True, "sid": sess["id"],
                        "filename": os.path.basename(str(dxf))})

    @app.post("/api/module-f/reopen")
    def module_f_reopen():
        """이미 찍어 둔 키로 손질부터 시작한다. 찍기 단계를 건너뛴다."""
        body = request.get_json(silent=True) or {}
        try:
            key = _check_key(body.get("key"))
        except ValueError as exc:
            return _fail(str(exc))
        try:
            _boot()
        except Exception as exc:  # noqa: BLE001
            return _fail(str(exc), 500)
        sess = _new_session(key=key)

        def job():
            from services.cad_import.edit.session import EditSession
            t0 = time.perf_counter()
            print(f"[손질] 저장본으로 배관망을 여는 중 — {key}")
            es = EditSession.open(key, out_dir=None, load_saved=True,
                                  use_cache=True)
            sess["edit"] = es
            sess["sheets"] = _sheet_frames(es.board)
            print(f"[손질] 완료 {time.perf_counter() - t0:.1f}s · "
                  f"노드 {len(es.board.pts)} · 간선 {len(es.board.edges)} · "
                  f"헤드 {len(es.board.disks)}")
            return {"key": key}

        _run_job(sess, "배관망 열기", job)
        return jsonify({"ok": True, "sid": sess["id"], "key": key})

    @app.get("/api/module-f/job")
    def module_f_job():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        view = _job_view(sess)
        view["ok"] = True
        view["stage"] = ("edit" if sess.get("edit") is not None
                         else ("pick" if sess.get("pick") is not None else ""))
        view["key"] = sess.get("key")
        return jsonify(view)

    @app.get("/api/module-f/world")
    def module_f_world():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("world") is None:
            return _fail("도면이 아직 준비되지 않았습니다.")
        return jsonify({"ok": True, "world": sess["world"],
                        "key": sess["key"], "state": _pick_state(sess)})
