# -*- coding: utf-8 -*-
"""Remote 30 계통도(system) 도메인 라우트.

대조 서버.py 에서 register(app, ...) 로 등록. 공유 헬퍼·전역은 main 에 두고 참조 주입 — 라우트 본문 원본 그대로, 엔드포인트명 보존. connection_review 의 __file__ 는 BASE_DIR 로 대체(동일 위치)."""
from __future__ import annotations

import json
import os
from pathlib import Path

from flask import jsonify, request


def register(app, *, BASE_DIR, COMBINED_OUTPUT_DIR, MACHINEROOM_OUTPUT_DIR, OVERALL_OUTPUT_DIR, PROTOTYPE_OUTPUT_DIR, SYSTEM_OUTPUT_DIR, _emit_subnetwork_bundle, _err500, _load_cached_view_entities, _save_upload, _serve_run_file, _sweep_old_run_dirs, _to_float):

    @app.post("/api/remote30/system/parse")
    def remote30_system_parse():
        """Remote 30 프로토타입 — 계통도 모드용 DXF 파싱.

        parse_dxf_for_view 사용:
          - hidden layer (is_off/is_frozen/color<0) 모두 포함 (도면 다 보이게)
          - POINT, LEADER, MLEADER, RAY, XLINE, WIPEOUT 등 추가 entity type 처리
          - 알 수 없는 type 은 virtual_entities 로 explode 시도
          - skipped/error 통계 반환
        """
        try:
            dxf_path = _save_upload("system_dxf_file", {".dxf", ".dwg"}, required=True)
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

    @app.post("/api/remote30/system/emit")
    def remote30_system_emit():
        """계통도(라이저) 단독 수리계산 파일 생성 — riser dict → SDF/SLF/KFP/ZIP.

        통합(combined) 을 거치지 않고 계통도 추출 결과만으로 부분 SDF 를 받기 위함.
        Body(JSON): { riser: extract_system_path 출력 dict }
        """
        import secrets
        body = request.get_json(silent=True) or {}
        riser = body.get("riser")
        if not riser or not riser.get("nodes") or not riser.get("pipes"):
            return jsonify({"ok": False,
                            "message": "riser (계통도 추출 결과) 가 필요합니다 — 계통도 추출을 먼저 실행하세요."}), 400
        from remote30_full_network import CombinedTables
        try:
            net = CombinedTables(
                nodes=list(riser["nodes"]),
                pipes=list(riser["pipes"]),
                pumps=list(riser.get("pumps", [])),
                valves=list(riser.get("valves", [])),
            )
            job_id = secrets.token_hex(6)
            _sweep_old_run_dirs(SYSTEM_OUTPUT_DIR, MACHINEROOM_OUTPUT_DIR,
                                COMBINED_OUTPUT_DIR, PROTOTYPE_OUTPUT_DIR, OVERALL_OUTPUT_DIR)
            out_dir = SYSTEM_OUTPUT_DIR / job_id
            files = _emit_subnetwork_bundle(
                net, out_dir, job_id, "system",
                f"Remote 30 계통도 — {riser.get('title', 'System')}",
                coord_scale=min(max(_to_float(body.get("kfp_coord_scale"), 1.0), 0.05), 20.0))
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)
        base = f"/api/remote30/system/result/{job_id}/"
        return jsonify({
            "ok": True, "job_id": job_id,
            "nodes": len(net.nodes), "pipes": len(net.pipes),
            "download_url_sdf": base + files["sdf"],
            "download_url_slf": (base + files["slf"]) if files["slf"] else None,
            "download_url_kfp": (base + files["kfp"]) if files["kfp"] else None,
            "download_url_zip": base + files["zip"],
        })

    @app.get("/api/remote30/system/result/<job_id>/<path:filename>")
    def remote30_system_result(job_id: str, filename: str):
        return _serve_run_file(SYSTEM_OUTPUT_DIR, job_id, filename)

    @app.post("/api/remote30/system/extract")
    def remote30_system_extract():
        """계통도 라이저 추출 — v1 (DXF 토폴로지) + legacy fallback.

        Multipart form:
            system_dxf_file        — 계통도 .dxf (v1 알고리즘 사용 시 필수)
            pump_x, pump_y         — 사용자 픽 펌프 좌표 (mm, 필수)
            av_x,   av_y           — 사용자 픽 알람밸브 좌표 (mm, 필수)
            use_legacy_template    — "true" 면 옛 affine template 사용 (DXF 불필요)
            snap_tolerance_mm      — 클릭 ↔ 그래프 노드 허용 거리 (기본 2500)

        v1 동작: DXF LINE 들로 그래프 빌드 → 펌프/AV 클릭점 → 가장 가까운 노드 매핑
            → Dijkstra → 경로를 PIPENET 호환 dict 로 반환.
        Legacy: extract_riser_msp_28f — 정답 28F 토폴로지 affine 변환 (DXF 무시).
        """
        # 좌표는 form 또는 JSON 둘 다 받기 (legacy JSON 호출자 호환)
        px = py = ax = ay = None
        use_legacy = False
        snap_tol = 2500.0
        waypoints: list[tuple[float, float]] = []

        def _parse_waypoints(raw):
            """waypoints 는 [[x,y], ...] JSON 문자열. 잘못된 형식은 무시(빈 리스트)."""
            if not raw:
                return []
            try:
                data = raw if isinstance(raw, list) else json.loads(raw)
                return [(float(p[0]), float(p[1])) for p in data]
            except (TypeError, ValueError, KeyError, IndexError):
                return []

        if request.is_json:
            body = request.get_json(silent=True) or {}
            try:
                px = float(body["pump_x"]); py = float(body["pump_y"])
                ax = float(body["av_x"]);   ay = float(body["av_y"])
            except (KeyError, TypeError, ValueError) as exc:
                return jsonify({"ok": False, "message": f"pump_x/y, av_x/y 좌표 필요: {exc}"}), 400
            use_legacy = bool(body.get("use_legacy_template"))
            try:
                snap_tol = float(body.get("snap_tolerance_mm", 2500.0))
            except (TypeError, ValueError):
                snap_tol = 2500.0
            waypoints = _parse_waypoints(body.get("waypoints"))
        else:
            try:
                px = float(request.form["pump_x"]); py = float(request.form["pump_y"])
                ax = float(request.form["av_x"]);   ay = float(request.form["av_y"])
            except (KeyError, TypeError, ValueError) as exc:
                return jsonify({"ok": False, "message": f"pump_x/y, av_x/y 좌표 필요: {exc}"}), 400
            use_legacy = request.form.get("use_legacy_template", "").lower() == "true"
            try:
                snap_tol = float(request.form.get("snap_tolerance_mm", "2500"))
            except (TypeError, ValueError):
                snap_tol = 2500.0
            waypoints = _parse_waypoints(request.form.get("waypoints"))

        if use_legacy:
            from remote30_prototype import extract_riser_msp_28f
            try:
                riser = extract_riser_msp_28f((px, py), (ax, ay))
                return jsonify({"ok": True, "riser": riser, "algorithm": "legacy_template"})
            except Exception as exc:  # noqa: BLE001
                return _err500(exc)

        # v1 — DXF 기반 path 추출 (DXF 파일 필수)
        try:
            dxf_path = _save_upload("system_dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False,
                            "message": f"DXF 파일 필요 (v1 알고리즘). legacy 사용하려면 use_legacy_template=true. ({exc})"}), 400

        # 선택: 주석 도면 (관경·층 라벨 TEXT 만 끌어와 결합).
        # 깨끗한 배관망 파일은 geometry 만 (TEXT 0) 갖고, 풀 도면은 annotation 을 갖되
        # geometry 가 파편화돼 있어 둘을 합친다 — geometry 는 primary, TEXT 는 annotation.
        anno_path = None
        try:
            anno_path = _save_upload("system_annotation_dxf_file", {".dxf", ".dwg"}, required=False)
        except ValueError:
            anno_path = None

        from remote30_prototype import parse_dxf_for_view, extract_system_path
        try:
            entities = _load_cached_view_entities(dxf_path)
            if entities is None:
                entities = parse_dxf_for_view(dxf_path, include_hidden_layers=True)["entities"]
            if anno_path is not None:
                anno_ents = _load_cached_view_entities(anno_path)
                if anno_ents is None:
                    anno_ents = parse_dxf_for_view(anno_path, include_hidden_layers=True)["entities"]
                anno_text = [e for e in anno_ents if e.get("t") == "T"]
                entities = entities + anno_text
            riser = extract_system_path(entities, (px, py), (ax, ay),
                                        snap_tolerance_mm=snap_tol,
                                        waypoints=waypoints or None)
            return jsonify({"ok": True, "riser": riser, "algorithm": "dxf_path_v1"})
        except ValueError as exc:
            # 사용자 입력 오류 (snap 실패 / disconnected). 상태코드 200 + suggest_legacy 표시.
            return jsonify({"ok": False, "message": str(exc),
                            "algorithm": "dxf_path_v1", "suggest_legacy": True}), 200
        except Exception as exc:  # noqa: BLE001
            return _err500(exc, algorithm="dxf_path_v1")

    @app.post("/api/remote30/system/connection_review")
    def remote30_system_connection_review():
        """연결복원 검수 오버레이 — 휴리스틱 × ML 합의 등급(A/CONFLICT/B/C).

        같은 계통도 DXF 를 받아 추출 계산망은 건드리지 않고(advisory), 복원 연결 후보를
        신뢰등급으로 분류해 좌표 JSON 으로 반환한다(프론트 점선 오버레이용).

          · A        : 휴리스틱∧ML 같은 끝단·같은 위치 → 고신뢰
          · CONFLICT : 둘 다 그 끝단을 잇지만 목표 다름 → 최우선 검수
          · B        : 휴리스틱 단독(거리 bridge, ML 침묵)
          · C        : ML 단독(T분기 포함, 휴리스틱이 못 만드는 연결)

        Multipart form:
            system_dxf_file — 계통도 .dxf/.dwg (필수)
            ml_cut          — ML top-1 채택 임계 (기본 0.45)
            mode            — 모델 코퍼스 (remote/all/allt, 기본 allt)
        """
        try:
            dxf_path = _save_upload("system_dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": f"DXF 파일 필요. ({exc})"}), 400

        if request.is_json:
            body = request.get_json(silent=True) or {}
            raw_cut, mode = body.get("ml_cut"), body.get("mode", "allt")
        else:
            raw_cut, mode = request.form.get("ml_cut"), request.form.get("mode", "allt")
        try:
            ml_cut = float(raw_cut) if raw_cut is not None else 0.45
        except (TypeError, ValueError):
            ml_cut = 0.45
        if mode not in ("remote", "all", "allt"):
            mode = "allt"

        import sys as _sys
        # 원본은 _Path(__file__).parent — 라우트 분리 후 위치가 바뀌므로 주입된
        # BASE_DIR(=대조 서버.py 의 부모, 동일 위치)로 대체.
        _cal = str((BASE_DIR / "calibration"))
        if _cal not in _sys.path:
            _sys.path.insert(0, _cal)
        import linkpred_integrate as li
        from remote30_prototype import parse_dxf_for_view

        pair = li.load_model(mode)
        if pair is None:
            return jsonify({"ok": False,
                            "message": f"연결복원 모델 없음(mode={mode}). "
                                       f"linkpred_train_v2.py {mode} 먼저 실행 필요."}), 503
        model, feats = pair
        try:
            entities = _load_cached_view_entities(dxf_path)
            if entities is None:
                entities = parse_dxf_for_view(dxf_path, include_hidden_layers=True)["entities"]
            res = li.reconcile_entities(entities, model, feats, ml_cut=ml_cut)
            payload = li.serialize_result(res)
            payload["mode"] = mode
            payload["ml_cut"] = ml_cut
            return jsonify(payload)
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

    @app.post("/api/remote30/system/clean_network")
    def remote30_system_clean_network():
        """임시 stopgap — 깨끗한(손작도) 배관망 DXF 전체를 그대로 길이와 함께 추출.

        풀 계통도가 조각나 강제 bridge 로 경로가 튀는 문제를 우회. 펌프/AV 클릭 없이
        파일에 그려진 단일 연결망을 그대로 pipe + 길이로 띄운다.

        Form/JSON (모두 선택):
            scale_mm_per_unit — 도면 1단위 = 실제 mm (기본 1.0, 용지 스케일이면 작게 나옴).
        파일 경로: env REMOTE30_CLEAN_SYSTEM_DXF 우선, 없으면
            samples/dxf/계통도_LH_306_배관망추출.dxf.
        """
        scale = 1.0
        raw_scale = None
        if request.is_json:
            raw_scale = (request.get_json(silent=True) or {}).get("scale_mm_per_unit")
        else:
            raw_scale = request.form.get("scale_mm_per_unit")
        try:
            if raw_scale is not None:
                scale = float(raw_scale)
        except (TypeError, ValueError):
            scale = 1.0

        clean_path = os.environ.get("REMOTE30_CLEAN_SYSTEM_DXF")
        clean_file = Path(clean_path) if clean_path else (BASE_DIR / "samples" / "dxf" / "계통도_LH_306_배관망추출.dxf")
        if not clean_file.is_file():
            return jsonify({"ok": False,
                            "message": f"깨끗한 배관망 파일 없음: {clean_file}. "
                                       f"REMOTE30_CLEAN_SYSTEM_DXF 로 경로 지정 가능."}), 200

        from remote30_prototype import extract_clean_system_network
        try:
            riser = extract_clean_system_network(clean_file, scale_mm_per_unit=scale)
            return jsonify({"ok": True, "riser": riser, "algorithm": "clean_network"})
        except Exception as exc:  # noqa: BLE001
            return _err500(exc, algorithm="clean_network")
