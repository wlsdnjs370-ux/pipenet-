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


def _open_job(sess: dict, dxf):
    """DXF 한 장을 찍기 세션으로 여는 잡을 만든다.

    [H-0] 슬롯 열기(`/api/module-f/slot/open`)와 같은 것을 쓴다 — 계통도·기계실도
    특허 S650 대로 «같은 절차» 를 밟으므로, 여기가 갈라지면 도면 종류에 따라
    제1국면이 달라진다.
    """
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
    return job


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
        _run_job(sess, "도면 읽기", _open_job(sess, dxf))
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

    @app.get("/api/module-f/job/stream")
    def module_f_job_stream():
        """[F-6] 진행 스트리밍 — 잡 상태·로그 줄을 SSE 로 흘린다.

        r30_prototype 의 SSE 패턴을 참조하되 세션·잡 규약은 F 것 그대로다:
        상태의 원천은 여전히 `_run_job` 의 sess["job"]/sess["log"] 이고, 이
        스트림은 그것을 읽어 보내기만 한다. 폴링(/api/module-f/job)은
        하위호환으로 남는다 — EventSource 가 없는 환경은 그리로 돌아간다.
        """
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)

        def gen():
            import json as _json
            sent = 0
            idle_beats = 0
            while True:
                view = _job_view(sess)
                lines = sess.get("log") or []
                # 새 로그 줄 — 줄 단위 이벤트로.
                while sent < len(lines):
                    yield ("event: line\ndata: "
                           + _json.dumps(lines[sent], ensure_ascii=False)
                           + "\n\n")
                    sent += 1
                yield ("event: state\ndata: "
                       + _json.dumps(view, ensure_ascii=False) + "\n\n")
                if view.get("state") in ("done", "error"):
                    return
                if view.get("state") == "idle":
                    # 잡이 아직 안 붙었을 수 있다 — 몇 박자는 기다려 준다.
                    idle_beats += 1
                    if idle_beats > 25:
                        return
                time.sleep(0.4)

        from flask import Response
        return Response(gen(), mimetype="text/event-stream",
                        headers={"Cache-Control": "no-cache",
                                 "X-Accel-Buffering": "no"})

    @app.get("/api/module-f/world")
    def module_f_world():
        try:
            sess = _sess(request.args.get("sid"))
        except ValueError as exc:
            return _fail(str(exc), 410)
        if sess.get("world") is None:
            return _fail("도면이 아직 준비되지 않았습니다.")
        # [H-2] 계통도·기계실 슬롯에는 찍기판이 없다 — 도면만 내려보낸다.
        # `_pick_state` 는 sess["pick"] 을 전제하므로 여기서 갈라야 한다.
        state = _pick_state(sess) if sess.get("pick") is not None else None
        return jsonify({"ok": True, "world": sess["world"],
                        "key": sess["key"], "state": state})
