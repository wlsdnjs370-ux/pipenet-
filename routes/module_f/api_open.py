# -*- coding: utf-8 -*-
"""모듈 F 라우트 — 페이지·설명 그림·열기/이어열기·진행·도면."""
from __future__ import annotations

import os
import time

from flask import jsonify, render_template, request, send_file

from routes.module_f.common import (
    DIAGRAMS, IMPORT_WORK_ROOT, _boot, _check_key, _fail)
from routes.module_f.jobs import (_job_view, _new_session, _run_job, route_session)
from routes.module_f.remote30 import _sheet_frames
from routes.module_f.views import _pick_state
from routes.module_f.world import _saved_keys, _world_payload


def _open_job(sess: dict, dxf, *, kind: str = "plan"):
    """DXF 한 장을 찍기 세션으로 여는 잡을 만든다.

    [H-0] 슬롯 열기(`/api/module-f/slot/open`)와 같은 것을 쓴다 — 계통도·기계실도
    특허 S650 대로 «같은 절차» 를 밟으므로, 여기가 갈라지면 도면 종류에 따라
    제1국면이 달라진다.

    [F-8a] 평면도는 찍기판을 연 «뒤에» 정찰까지 같은 잡에서 이어 돌린다
    (D-F8-1 — 버튼이 아니라 업로드 직후 자동). 순서가 요점이다: `sess["world"]`
    는 정찰 전에 이미 앉으므로 화면은 정찰이 도는 동안에도 도면을 그린다.
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
        t_open = time.perf_counter() - t0
        print(f"[찍기] 완료 {t_open:.1f}s · "
              f"선분 {payload['counts']['segs']} · "
              f"원 {payload['counts']['circles']} · "
              f"호 {payload['counts']['arcs']}")
        # ★여기부터가 덤이다 — 위까지로 도면은 이미 화면에 그려진다.
        #   [D-F8-2] 계통도·기계실은 두 점 경로가 전부라 헤드 검출이 무의미하다.
        if str(kind) == "plan":
            _recon_into(sess, dxf, payload, t_open)
        return {"key": ps.key}
    return job


def _recon_into(sess: dict, dxf, payload: dict, t_open: float) -> None:
    """[F-8a] 정찰을 돌려 `sess["recon"]` 에 남긴다.

    ★정찰 실패는 열기 실패가 아니다(F-8a-4). A 를 못 부르든 도면이 이상하든
      찍기는 종전대로 돌아야 하므로, 무엇이 나오든 여기서 삼키고 사유만 남긴다.
      `_run_job` 이 BaseException 까지 잡는 것과 같은 이유로 여기도 그렇게 한다
      — 엔진 계열 코드는 실패를 SystemExit 로 던지는 곳이 있다.
    """
    from routes.module_f.recon import (
        BAND_HIGH, BAND_LOW, BAND_MID, run_recon)
    print(f"[정찰] 자동 인식 중… — 도면은 이미 화면에 있습니다(+{t_open:.1f}s)")
    try:
        rec = run_recon(dxf, world=payload)
    except BaseException as exc:  # noqa: BLE001 — 열기를 죽이지 않는다
        why = f"{type(exc).__name__}: {exc}"
        sess["recon"] = {"error": why}
        print(f"[정찰] 건너뜀 — {why} (찍기는 종전대로 쓸 수 있습니다)")
        return
    sess["recon"] = rec
    b, n = rec["bundles"], rec["bands"]
    print(f"[정찰] 배관 묶음 {b.get('PIPE', 0)} · 헤드 후보 {len(rec['heads'])} "
          f"(높음 {n[BAND_HIGH]}·중간 {n[BAND_MID]}·낮음 {n[BAND_LOW]}) "
          f"— {rec['elapsed_ms'] / 1000:.1f}s")


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
        # 이 문은 평면도 전용이다 — 계통도·기계실은 `/slot/open` 으로 들어온다.
        _run_job(sess, "도면 읽기", _open_job(sess, dxf, kind="plan"))
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
    @route_session()
    def module_f_job(sess, body):
        view = _job_view(sess)
        view["ok"] = True
        view["stage"] = ("edit" if sess.get("edit") is not None
                         else ("pick" if sess.get("pick") is not None else ""))
        view["key"] = sess.get("key")
        return jsonify(view)

    @app.get("/api/module-f/job/stream")
    @route_session()
    def module_f_job_stream(sess, body):
        """[F-6] 진행 스트리밍 — 잡 상태·로그 줄을 SSE 로 흘린다.

        r30_prototype 의 SSE 패턴을 참조하되 세션·잡 규약은 F 것 그대로다:
        상태의 원천은 여전히 `_run_job` 의 sess["job"]/sess["log"] 이고, 이
        스트림은 그것을 읽어 보내기만 한다. 폴링(/api/module-f/job)은
        하위호환으로 남는다 — EventSource 가 없는 환경은 그리로 돌아간다.
        """
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

    @app.get("/api/module-f/recon")
    @route_session()
    def module_f_recon(sess, body):
        """[F-8a] 정찰 결과 조회 — 새로고침해도 카드가 다시 채워지게.

        수치만 준다. 후보 좌표 수천 개는 `heads=1` 로 따로 청한다 — 카드를
        그릴 때마다 3천 점을 내려보내면 새로고침이 그만큼 무거워진다.
        """
        from routes.module_f.recon import recon_view
        rec = sess.get("recon")
        out = {"ok": True, "recon": recon_view(rec)}
        if request.args.get("heads") in ("1", "true", "yes"):
            out["heads"] = (rec or {}).get("heads") or []
        return jsonify(out)

    @app.get("/api/module-f/world")
    @route_session()
    def module_f_world(sess, body):
        if sess.get("world") is None:
            return _fail("도면이 아직 준비되지 않았습니다.")
        # [H-2] 계통도·기계실 슬롯에는 찍기판이 없다 — 도면만 내려보낸다.
        # `_pick_state` 는 sess["pick"] 을 전제하므로 여기서 갈라야 한다.
        state = _pick_state(sess) if sess.get("pick") is not None else None
        return jsonify({"ok": True, "world": sess["world"],
                        "key": sess["key"], "state": state})
