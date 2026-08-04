# -*- coding: utf-8 -*-
"""Remote 30 프로토타입 도메인 — DXF → 4-stage 파이프라인 + SSE 실시간 진행.

`대조 서버.py` 에서 `register(app, ...)` 로 등록. 공유 헬퍼·전역 상태
(_save_upload/_serve_run_file/_PROTOTYPE_JOBS 등)는 main 에 그대로 두고 참조로
주입한다 — 라우트 본문은 원본 그대로(수정 0), 엔드포인트명 보존.
"""
from __future__ import annotations

import json
import threading
from pathlib import Path

from flask import Response, jsonify, make_response, render_template, request

# 같은 job_id 를 두 번 구독하면(재구독·새로고침) 두 제너레이터가 하나의 job
# 딕셔너리에 스테이지별로 나눠 쓴다. 한쪽의 layers 와 다른쪽의 detected_heads 가
# 섞이면 finalize 가 어긋난 상태를 본다.
_JOB_LOCK = threading.Lock()


def register(app, *, _save_upload, _register_job, _serve_run_file,
             _sweep_old_run_dirs, _PROTOTYPE_JOBS,
             PROTOTYPE_OUTPUT_DIR, OVERALL_OUTPUT_DIR, COMBINED_OUTPUT_DIR):

    @app.get("/remote30-prototype")
    def remote30_prototype():
        response = make_response(render_template("remote30_prototype.html"))
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return response

    @app.post("/api/remote30/prototype/run")
    def remote30_prototype_run():
        """DXF 업로드 → 백그라운드 잡 시작 → job_id 반환. 진행은 /stream/<job_id> 으로 구독.

        Form fields (옵션):
            alarm_x, alarm_y: 알람밸브 좌표 (둘 다 또는 둘 다 없음 — 없으면 auto)
        """
        import secrets
        try:
            dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        alarm_x = request.form.get("alarm_x", "").strip()
        alarm_y = request.form.get("alarm_y", "").strip()
        alarm_xy: tuple[float, float] | None = None
        if alarm_x and alarm_y:
            try:
                alarm_xy = (float(alarm_x), float(alarm_y))
            except ValueError:
                return jsonify({"ok": False, "message": "alarm_x/alarm_y 는 숫자여야 합니다."}), 400
        elif alarm_x or alarm_y:
            return jsonify({"ok": False, "message": "alarm_x 와 alarm_y 는 함께 입력하거나 둘 다 비워야 합니다."}), 400
        job_id = secrets.token_hex(6)
        out_dir = PROTOTYPE_OUTPUT_DIR / job_id
        out_dir.mkdir(parents=True, exist_ok=True)
        _sweep_old_run_dirs(PROTOTYPE_OUTPUT_DIR, OVERALL_OUTPUT_DIR, COMBINED_OUTPUT_DIR)
        _register_job(_PROTOTYPE_JOBS, job_id, {
            "dxf_path": str(dxf_path),
            "out_dir": str(out_dir),
            "dxf_filename": dxf_path.name,
            "alarm_xy": alarm_xy,
        })
        return jsonify({"ok": True, "job_id": job_id, "dxf_filename": dxf_path.name,
                        "alarm_xy": list(alarm_xy) if alarm_xy else None})

    @app.get("/api/remote30/prototype/stream/<job_id>")
    def remote30_prototype_stream(job_id: str):
        """Stage 0~2 만 진행 — 헤드 인식까지 마치고 stream 종료. 그 시점에서 사용자 편집 대기."""
        job = _PROTOTYPE_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404

        from remote30_prototype import run_stages_0_2

        # 분기영역 — 사용자가 Stage-3 에서 드래그 후 스트림 재구독 시 쿼리로 전달.
        # job 에 저장해 이후 finalize 에서도 재사용. 미지정이면 종전대로 전체 렌더.
        bz_raw = request.args.get("branch_zones")
        if bz_raw:
            try:
                job["branch_zones"] = [tuple(z) for z in json.loads(bz_raw)]
            except (ValueError, TypeError):
                pass

        # 알람밸브 좌표 override — 분기영역 corridor 는 source(알람밸브) 가 필수인데,
        # 사용자가 빌드 후 캔버스/입력필드로 지정한 좌표는 job["alarm_xy"](빌드 시점값)
        # 에 반영돼 있지 않다. rebuild 재구독 시 쿼리로 받아 job 에 갱신해 반영한다.
        ax_raw = request.args.get("alarm_x")
        ay_raw = request.args.get("alarm_y")
        if ax_raw and ay_raw:
            try:
                job["alarm_xy"] = (float(ax_raw), float(ay_raw))
            except (ValueError, TypeError):
                pass

        def _gen():
            # 스테이지 산출은 이 스트림의 로컬 값으로만 만들고 한 번에 반영한다.
            stream_layers: list = []
            try:
                for evt in run_stages_0_2(Path(job["dxf_path"]), job_id,
                                           alarm_xy=job.get("alarm_xy"),
                                           branch_zones=job.get("branch_zones")):
                    # 마지막 awaiting_finalize 이벤트 직전에 detected_heads 데이터를 job 에 저장
                    patch = None
                    if evt.get("type") == "entities" and evt.get("stage") == 1:
                        patch = {"pipe_ents": evt["entities"]}
                    elif evt.get("type") == "entities" and evt.get("stage") == 0:
                        stream_layers = evt["layers"]
                        patch = {"layers": stream_layers, "bbox": evt["bbox"]}
                    elif evt.get("type") == "entities" and evt.get("stage") == 2:
                        # bbox entity 들에서 detected_heads 위치 + bbox 추출
                        detected = []
                        for be in evt["entities"]:
                            if be.get("t") == "B":
                                p = be["p"]
                                cx = (p[0] + p[2]) / 2; cy = (p[1] + p[3]) / 2
                                detected.append({"pos": [cx, cy], "bbox": p,
                                                 "k": be.get("k", ""), "c": be.get("c", 0),
                                                 "i": be.get("i", 0)})
                        patch = {"detected_heads": detected,
                                 "layer_cat": {l["name"]: l["auto_category"]
                                               for l in stream_layers}}
                    if patch is not None:
                        with _JOB_LOCK:
                            job.update(patch)
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
            except Exception as exc:  # noqa: BLE001
                err = {"type": "error", "message": str(exc)[:500]}
                yield f"data: {json.dumps(err, ensure_ascii=False)}\n\n"

        response = Response(_gen(), mimetype="text/event-stream")
        response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @app.post("/api/remote30/prototype/finalize/<job_id>")
    def remote30_prototype_finalize(job_id: str):
        """사용자 편집 데이터 (added/deleted heads + zones + alarm_xy) 수신.

        body (JSON):
            added_heads: [[x,y], ...]
            deleted_indices: [int, ...]
            zones: [[x1,y1,x2,y2], ...]
            alarm_x, alarm_y: float | null (선택 — 비우면 자동)
        """
        job = _PROTOTYPE_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "detected_heads" not in job:
            return jsonify({"ok": False, "message": "Stage 2 가 아직 끝나지 않았습니다."}), 400
        body = request.get_json(silent=True) or {}
        job["edit"] = {
            "added_heads": [tuple(p) for p in body.get("added_heads", [])],
            "deleted_indices": [int(i) for i in body.get("deleted_indices", [])],
            "zones": [tuple(z) for z in body.get("zones", [])],
            # 분기영역 — body 우선, 없으면 스트림 재구독 때 저장해둔 job 값 재사용.
            "branch_zones": [tuple(z) for z in body.get("branch_zones",
                                                        job.get("branch_zones") or [])],
        }
        # alarm_xy 갱신 (사용자가 후속으로 변경했을 수 있음)
        ax, ay = body.get("alarm_x"), body.get("alarm_y")
        if ax is not None and ay is not None:
            try:
                job["alarm_xy"] = (float(ax), float(ay))
            except (TypeError, ValueError):
                pass
        # 불리한 헤드 개수 N (미지정 시 기본 30 — run_stages_3_5 default)
        n_heads = body.get("n_heads")
        if n_heads is not None:
            try:
                n_val = int(n_heads)
                if n_val >= 1:
                    job["k_heads"] = n_val
            except (TypeError, ValueError):
                pass
        return jsonify({"ok": True, "job_id": job_id,
                        "added": len(job["edit"]["added_heads"]),
                        "deleted": len(job["edit"]["deleted_indices"]),
                        "zones": len(job["edit"]["zones"])})

    @app.get("/api/remote30/prototype/finalize_stream/<job_id>")
    def remote30_prototype_finalize_stream(job_id: str):
        """Stage 3~5 SSE — finalize() 호출 후에 구독."""
        job = _PROTOTYPE_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "edit" not in job:
            return jsonify({"ok": False, "message": "finalize() 먼저 호출하세요."}), 400

        from remote30_prototype import run_stages_3_5

        def _gen():
            try:
                detected_pos = [tuple(d["pos"]) for d in job.get("detected_heads", [])]
                for evt in run_stages_3_5(
                    Path(job["dxf_path"]), Path(job["out_dir"]), job_id,
                    pipe_ents=job.get("pipe_ents", []),
                    layer_categories=job.get("layer_cat", {}),
                    detected_heads_pos=detected_pos,
                    k_heads=job.get("k_heads", 30),
                    alarm_xy=job.get("alarm_xy"),
                    user_added_heads=job["edit"]["added_heads"],
                    user_deleted_indices=job["edit"]["deleted_indices"],
                    zones=job["edit"]["zones"],
                    branch_zones=job["edit"].get("branch_zones") or None,
                ):
                    # 신축배관(FX) 검토 게이트 — stage2_complete 저장 방식과 동일하게
                    # tables/fx_review 를 job state 에 저장하고 이벤트는 그대로 프론트에 전달.
                    if evt.get("type") == "stage5_complete":
                        job["tables"] = evt["tables"]
                        job["project_title"] = evt.get("project_title", "")
                        job["fx_review"] = evt.get("fx_review", {})
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
            except Exception as exc:  # noqa: BLE001
                err = {"type": "error", "message": str(exc)[:500]}
                yield f"data: {json.dumps(err, ensure_ascii=False)}\n\n"

        response = Response(_gen(), mimetype="text/event-stream")
        response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @app.post("/api/remote30/prototype/fx/finalize/<job_id>")
    def remote30_prototype_fx_finalize(job_id: str):
        """신축배관(FX) 검토 편집 결과 수신 — 헤드 편집 finalize 와 동일 컨벤션.

        body (JSON):
            equipment: [ {label, spec_ref, eq_len, source, override_flag, override_note, ...}, ... ]
                       생략/null 이면 원본 그대로 emit (편집 없이 확정 = 회귀 없음).
        저장만 하고 실제 emit 은 fx/finalize_stream 에서 스트리밍.
        """
        job = _PROTOTYPE_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "tables" not in job:
            return jsonify({"ok": False, "message": "stage5 (테이블 생성) 가 아직 끝나지 않았습니다."}), 400
        body = request.get_json(silent=True) or {}
        edited = body.get("equipment")
        job["fx_edit"] = edited if isinstance(edited, list) else None
        return jsonify({"ok": True, "job_id": job_id,
                        "edited": (len(job["fx_edit"]) if job["fx_edit"] is not None else 0),
                        "mode": ("edited" if job["fx_edit"] is not None else "original")})

    @app.get("/api/remote30/prototype/fx/finalize_stream/<job_id>")
    def remote30_prototype_fx_finalize_stream(job_id: str):
        """Stage 6 (SDF emit) SSE — fx/finalize() 호출 후 구독."""
        job = _PROTOTYPE_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "tables" not in job:
            return jsonify({"ok": False, "message": "stage5 결과가 없습니다. finalize_stream 을 먼저 완료하세요."}), 400

        from remote30_prototype import run_stage_6_emit, PipeTables

        def _gen():
            try:
                tables = PipeTables.from_dict(job["tables"])
                for evt in run_stage_6_emit(
                    Path(job["out_dir"]), job_id, tables,
                    edited_equipment=job.get("fx_edit"),
                    project_title=job.get("project_title") or Path(job["dxf_path"]).stem,
                ):
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
            except Exception as exc:  # noqa: BLE001
                err = {"type": "error", "message": str(exc)[:500]}
                yield f"data: {json.dumps(err, ensure_ascii=False)}\n\n"

        response = Response(_gen(), mimetype="text/event-stream")
        response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @app.get("/api/remote30/prototype/result/<job_id>/<path:filename>")
    def remote30_prototype_result(job_id: str, filename: str):
        return _serve_run_file(PROTOTYPE_OUTPUT_DIR, job_id, filename)
