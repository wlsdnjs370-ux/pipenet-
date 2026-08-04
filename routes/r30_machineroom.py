# -*- coding: utf-8 -*-
"""Remote 30 기계실(machineroom) 도메인 라우트.

대조 서버.py 에서 register(app, ...) 로 등록. 공유 헬퍼·전역은 main 에 두고 참조 주입 — 라우트 본문 원본 그대로, 엔드포인트명 보존."""
from __future__ import annotations

import json

from flask import jsonify, request


def register(app, *, COMBINED_OUTPUT_DIR, MACHINEROOM_OUTPUT_DIR, OVERALL_OUTPUT_DIR, PROTOTYPE_OUTPUT_DIR, SYSTEM_OUTPUT_DIR, _emit_subnetwork_bundle, _err500, _load_cached_view_entities, _save_upload, _serve_run_file, _sweep_old_run_dirs, _to_float):

    @app.post("/api/remote30/machineroom/emit")
    def remote30_machineroom_emit():
        """기계실(옥상수조) 경로 단독 수리계산 파일 생성 — machine_room dict → SDF/SLF/KFP/ZIP.

        Body(JSON): { machine_room: extract_machine_room_path 출력 dict }
        """
        import secrets
        body = request.get_json(silent=True) or {}
        mr = body.get("machine_room")
        if not mr or not mr.get("nodes") or not mr.get("pipes"):
            return jsonify({"ok": False,
                            "message": "machine_room (기계실 추출 결과) 가 필요합니다 — 기계실 추출을 먼저 실행하세요."}), 400
        from remote30_full_network import CombinedTables, insert_source_pump
        try:
            net = CombinedTables(
                nodes=list(mr["nodes"]),
                pipes=list(mr["pipes"]),
            )
            # 펌프 노드 추가 (옵션) — 수원(Input 경계) 직후에 가압펌프 삽입.
            # 기계실 단독망은 기본이 옥상수조 자연낙차라 펌프가 없어 입력검사에서
            # "펌프 누락" 으로 뜬다. 사용자가 켜면 화재안전기준 3점 성능곡선 펌프를 넣는다.
            if body.get("add_pump"):
                try:
                    q = float(body.get("pump_q_lpm", 2400.0))
                    h = float(body.get("pump_h_m", 100.0))
                    n = int(body.get("pump_count", 1))
                except (TypeError, ValueError):
                    q, h, n = 2400.0, 100.0, 1
                insert_source_pump(net, rated_q_lpm=q, rated_h_m=h, count=max(1, n))
            job_id = secrets.token_hex(6)
            _sweep_old_run_dirs(SYSTEM_OUTPUT_DIR, MACHINEROOM_OUTPUT_DIR,
                                COMBINED_OUTPUT_DIR, PROTOTYPE_OUTPUT_DIR, OVERALL_OUTPUT_DIR)
            out_dir = MACHINEROOM_OUTPUT_DIR / job_id
            files = _emit_subnetwork_bundle(
                net, out_dir, job_id, "machineroom",
                f"Remote 30 기계실 — {mr.get('title', 'MachineRoom')}",
                coord_scale=min(max(_to_float(body.get("kfp_coord_scale"), 1.0), 0.05), 20.0))
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)
        base = f"/api/remote30/machineroom/result/{job_id}/"
        return jsonify({
            "ok": True, "job_id": job_id,
            "nodes": len(net.nodes), "pipes": len(net.pipes),
            "pumps": len(net.pumps),
            "download_url_sdf": base + files["sdf"],
            "download_url_slf": (base + files["slf"]) if files["slf"] else None,
            "download_url_kfp": (base + files["kfp"]) if files["kfp"] else None,
            "download_url_zip": base + files["zip"],
        })

    @app.get("/api/remote30/machineroom/result/<job_id>/<path:filename>")
    def remote30_machineroom_result(job_id: str, filename: str):
        return _serve_run_file(MACHINEROOM_OUTPUT_DIR, job_id, filename)

    @app.post("/api/remote30/machineroom/parse")
    def remote30_machineroom_parse():
        """기계실 모드용 DXF 파싱 — 캔버스 표시용 (system/parse 와 동형)."""
        try:
            dxf_path = _save_upload("machineroom_dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        from remote30_prototype import parse_dxf_for_view
        try:
            result = parse_dxf_for_view(dxf_path, include_hidden_layers=True)
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)
        result["ok"] = True
        result["filename"] = dxf_path.name
        return jsonify(result)

    @app.post("/api/remote30/machineroom/extract")
    def remote30_machineroom_extract():
        """기계실(옥상수조) 경로 추출 — 탱크(수원) → 입상관 연결점 2점 클릭.

        Multipart form (또는 JSON):
            machineroom_dxf_file   — 기계실 .dxf (필수)
            source_x, source_y     — 탱크 토출구(수원) 좌표 (mm, 필수)
            conn_x,   conn_y       — 입상관 연결점 좌표 (mm, 필수)
            snap_tolerance_mm      — 클릭 ↔ 그래프 노드 허용 거리 (기본 3000)
            machine_room_ceiling_m — 기계실 천장고 (m). 수면→천장 아래 배관 낙차로
                                     첫 구간에 적용. 없으면 그 구간 표고는 미확정.

        동작: 계통도 추출(extract_system_path)과 동형 — DXF LINE 으로 그래프 빌드 →
            탱크/연결점 클릭점을 가장 가까운 노드에 snap → Dijkstra 경로 → 기계실 dict.
            결과는 combined/build 의 machine_room 입력으로 전달.
        """
        sx = sy = cx = cy = None
        snap_tol = 3000.0
        if request.is_json:
            body = request.get_json(silent=True) or {}
            try:
                sx = float(body["source_x"]); sy = float(body["source_y"])
                cx = float(body["conn_x"]);   cy = float(body["conn_y"])
            except (KeyError, TypeError, ValueError) as exc:
                return jsonify({"ok": False, "message": f"source_x/y, conn_x/y 좌표 필요: {exc}"}), 400
            try:
                snap_tol = float(body.get("snap_tolerance_mm", 3000.0))
            except (TypeError, ValueError):
                snap_tol = 3000.0
        else:
            try:
                sx = float(request.form["source_x"]); sy = float(request.form["source_y"])
                cx = float(request.form["conn_x"]);   cy = float(request.form["conn_y"])
            except (KeyError, TypeError, ValueError) as exc:
                return jsonify({"ok": False, "message": f"source_x/y, conn_x/y 좌표 필요: {exc}"}), 400
            try:
                snap_tol = float(request.form.get("snap_tolerance_mm", "3000"))
            except (TypeError, ValueError):
                snap_tol = 3000.0

        # 천장고는 비우면 미확정(None)으로 남긴다. 0 은 "천장고가 0" 이라는 주장이라
        # 미입력과 같지 않고, 숫자가 아니면 조용히 흡수하지 않고 되돌려준다.
        _ceiling_raw = (request.get_json(silent=True) or {}).get("machine_room_ceiling_m") \
            if request.is_json else request.form.get("machine_room_ceiling_m")
        ceiling_m = None
        if _ceiling_raw not in (None, ""):
            try:
                ceiling_m = float(_ceiling_raw)
            except (TypeError, ValueError):
                return jsonify({"ok": False,
                                "message": "machine_room_ceiling_m 는 숫자여야 합니다."}), 400

        # 사용자가 지정한 배관 레이어 — 다계통(PIPE/PIPE_MAIN/PIPE_SUB/소화수관…) 도면에서
        # '어떤 레이어가 이 기계실 배관인가'를 명시. 없으면 자동(SP 레이어→키워드) 추론.
        pipe_layers = None
        _pl_raw = (request.get_json(silent=True) or {}).get("pipe_layers") if request.is_json \
            else request.form.get("pipe_layers", "")
        if _pl_raw:
            try:
                _pl = json.loads(_pl_raw) if isinstance(_pl_raw, str) else _pl_raw
                if isinstance(_pl, list):
                    pipe_layers = {str(x) for x in _pl if str(x).strip()} or None
            except (json.JSONDecodeError, TypeError):
                pipe_layers = None

        try:
            dxf_path = _save_upload("machineroom_dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": f"기계실 DXF 파일 필요: {exc}"}), 400

        from remote30_prototype import parse_dxf_for_view, extract_machine_room_path
        try:
            entities = _load_cached_view_entities(dxf_path)
            if entities is None:
                entities = parse_dxf_for_view(dxf_path, include_hidden_layers=True)["entities"]
            mr = extract_machine_room_path(entities, (sx, sy), (cx, cy),
                                           layer_filter=pipe_layers,
                                           snap_tolerance_mm=snap_tol,
                                           ceiling_m=ceiling_m)
            return jsonify({"ok": True, "machine_room": mr, "algorithm": "machineroom_path_v1"})
        except ValueError as exc:
            # 사용자 입력 오류 (snap 실패 / disconnected). 프론트는 본문 ok 로 분기한다.
            return jsonify({"ok": False, "message": str(exc),
                            "algorithm": "machineroom_path_v1"}), 400
        except Exception as exc:  # noqa: BLE001
            return _err500(exc, algorithm="machineroom_path_v1")
