# -*- coding: utf-8 -*-
"""Remote 30 전체 배관망 총괄 (10번 모듈) 도메인 라우트.

대조 서버.py 에서 register(app, ...) 로 등록. 공유 헬퍼·전역 상태는
main 에 두고 참조 주입 — 라우트 본문 원본 그대로, 엔드포인트명 보존.
_save_pressure_table_upload 는 overall 전용이라 이 모듈로 이동."""
from __future__ import annotations

import json
from pathlib import Path

from flask import Response, jsonify, make_response, render_template, request


def _save_pressure_table_upload(field_name: str, out_dir: Path) -> Path | None:
    """선택적 압력표 파일 업로드 (csv/xlsx) — 파일이 있으면 out_dir 로 저장 후 경로 반환."""
    from werkzeug.utils import secure_filename as _sec
    f = request.files.get(field_name)
    if not f or not f.filename:
        return None
    suffix = Path(f.filename).suffix.lower()
    if suffix not in {".csv", ".xlsx"}:
        raise ValueError(f"{field_name} 은 .csv 또는 .xlsx 만 허용합니다.")
    safe = _sec(f.filename) or f"upload{suffix}"
    dst = out_dir / safe
    f.save(dst)
    return dst


def register(app, *, _err500, _register_job, _save_upload, _serve_run_file, _sweep_old_run_dirs, _OVERALL_JOBS, OVERALL_OUTPUT_DIR, PROTOTYPE_OUTPUT_DIR, COMBINED_OUTPUT_DIR):

    @app.get("/remote30-overall")
    def remote30_overall():
        """10번 모듈 — Remote 30 전체 배관망 총괄.

        Stage A (헤드망 추출 — remote30_prototype 재사용) + Stage B (zone 별 라이저 템플릿)
        + Stage C (stitch) + Stage D (PIPENET-native 후처리) 로 완성 SDF 생성.
        현재는 모듈 자리만 마련된 상태이고 API 는 task #14~#16 에서 추가.
        """
        response = make_response(render_template("remote30_overall.html"))
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        return response

    @app.post("/api/remote30/overall/run")
    def remote30_overall_run():
        """DXF + zone_spec(form) + (선택) 압력표 업로드 → job_id 반환.

        Form fields (필수):
            dxf_file: 평면도 DXF
            zone_type: "hsp_pump" / "lsp_1stage" / "llsp_2stage" / "lsp_gravity"
            target_floor: "16층" 등 (라이저 AV elevation 결정용)

        Form fields (zone 별 필수):
            prv1_target_kgf 또는 prv1_target_m  (HSP_PUMP, LSP_1STAGE, LLSP_2STAGE)
            prv2_target_kgf 또는 prv2_target_m  (LLSP_2STAGE 만)

        Form fields (옵션):
            alarm_x, alarm_y: 알람밸브 좌표 (둘 다 또는 둘 다 없음)
            pump_library_name: HSP_PUMP 의 Library-pump 이름 (기본 SP_162M_2900LPM)
            pump_count: HSP_PUMP 의 Pump-fan 개수 (기본 2)
            building_name: 빌딩 식별자
            pressure_table_csv: 압력표 CSV 파일
            pressure_table_xlsx: 압력표 엑셀 파일
            pressure_table_json: 압력표 직접 입력 (JSON 문자열)
        """
        import secrets
        from remote30_full_network import zone_spec_from_form, ZoneType

        try:
            dxf_path = _save_upload("dxf_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400

        # zone_spec 검증
        zone_type_raw = request.form.get("zone_type", "").strip()
        try:
            ZoneType(zone_type_raw)
        except ValueError:
            return jsonify({"ok": False,
                            "message": f"zone_type 필요 — 허용값: {[z.value for z in ZoneType]}"}), 400

        spec = zone_spec_from_form(request.form)
        if not spec.target_floor:
            return jsonify({"ok": False, "message": "target_floor (예: '16층') 가 필요합니다."}), 400
        if spec.zone_type in {ZoneType.HSP_PUMP, ZoneType.LSP_1STAGE, ZoneType.LLSP_2STAGE} \
                and spec.prv1_target_pa is None:
            return jsonify({"ok": False,
                            "message": "prv1_target_kgf 또는 prv1_target_m 이 필요합니다."}), 400
        if spec.zone_type == ZoneType.LLSP_2STAGE and spec.prv2_target_pa is None:
            return jsonify({"ok": False,
                            "message": "LLSP_2STAGE 는 prv2_target_kgf 또는 prv2_target_m 도 필요합니다."}), 400

        # 알람밸브 좌표 (옵션) — prototype 와 동일 패턴
        alarm_x = request.form.get("alarm_x", "").strip()
        alarm_y = request.form.get("alarm_y", "").strip()
        alarm_xy: tuple[float, float] | None = None
        if alarm_x and alarm_y:
            try:
                alarm_xy = (float(alarm_x), float(alarm_y))
            except ValueError:
                return jsonify({"ok": False, "message": "alarm_x/alarm_y 는 숫자여야 합니다."}), 400
        elif alarm_x or alarm_y:
            return jsonify({"ok": False,
                            "message": "alarm_x 와 alarm_y 는 함께 또는 둘 다 비워야 합니다."}), 400

        # 잡 등록 + 출력 디렉토리
        job_id = secrets.token_hex(6)
        out_dir = OVERALL_OUTPUT_DIR / job_id
        out_dir.mkdir(parents=True, exist_ok=True)

        # 선택적 압력표 파일 저장 (있으면)
        try:
            pressure_csv = _save_pressure_table_upload("pressure_table_csv", out_dir)
            pressure_xlsx = _save_pressure_table_upload("pressure_table_xlsx", out_dir)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400

        _sweep_old_run_dirs(PROTOTYPE_OUTPUT_DIR, OVERALL_OUTPUT_DIR, COMBINED_OUTPUT_DIR)
        _register_job(_OVERALL_JOBS, job_id, {
            "dxf_path": str(dxf_path),
            "dxf_filename": dxf_path.name,
            "out_dir": str(out_dir),
            "alarm_xy": alarm_xy,
            "spec_form": dict(request.form),  # finalize 시 ZoneSpec 재구성
            "pressure_table_csv": str(pressure_csv) if pressure_csv else None,
            "pressure_table_xlsx": str(pressure_xlsx) if pressure_xlsx else None,
        })
        return jsonify({
            "ok": True, "job_id": job_id, "dxf_filename": dxf_path.name,
            "zone_type": spec.zone_type.value, "target_floor": spec.target_floor,
            "prv1_target_pa": spec.prv1_target_pa,
            "prv2_target_pa": spec.prv2_target_pa,
            "alarm_xy": list(alarm_xy) if alarm_xy else None,
            "pressure_table": (pressure_csv and pressure_csv.name) or (pressure_xlsx and pressure_xlsx.name),
        })

    @app.get("/api/remote30/overall/stream/<job_id>")
    def remote30_overall_stream(job_id: str):
        """Stage A SSE — prototype 의 stream 과 동일 로직 (run_stages_0_2 재사용)."""
        job = _OVERALL_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404

        from remote30_prototype import run_stages_0_2

        bz_raw = request.args.get("branch_zones")
        if bz_raw:
            try:
                job["branch_zones"] = [tuple(z) for z in json.loads(bz_raw)]
            except (ValueError, TypeError):
                pass

        def _gen():
            try:
                for evt in run_stages_0_2(Path(job["dxf_path"]), job_id,
                                           alarm_xy=job.get("alarm_xy"),
                                           branch_zones=job.get("branch_zones")):
                    if evt.get("type") == "entities" and evt.get("stage") == 1:
                        job["pipe_ents"] = evt["entities"]
                    elif evt.get("type") == "entities" and evt.get("stage") == 0:
                        job["layers"] = evt["layers"]
                        job["bbox"] = evt["bbox"]
                    elif evt.get("type") == "entities" and evt.get("stage") == 2:
                        detected = []
                        for be in evt["entities"]:
                            if be.get("t") == "B":
                                p = be["p"]
                                cx = (p[0] + p[2]) / 2; cy = (p[1] + p[3]) / 2
                                detected.append({"pos": [cx, cy], "bbox": p,
                                                 "k": be.get("k", ""), "c": be.get("c", 0),
                                                 "i": be.get("i", 0)})
                        job["detected_heads"] = detected
                        job["layer_cat"] = {l["name"]: l["auto_category"] for l in job.get("layers", [])}
                    yield f"data: {json.dumps(evt, ensure_ascii=False)}\n\n"
            except Exception as exc:  # noqa: BLE001
                err = {"type": "error", "message": str(exc)[:500]}
                yield f"data: {json.dumps(err, ensure_ascii=False)}\n\n"

        response = Response(_gen(), mimetype="text/event-stream")
        response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @app.post("/api/remote30/overall/finalize/<job_id>")
    def remote30_overall_finalize(job_id: str):
        """사용자 편집(added/deleted heads + zones + alarm_xy 갱신) 수신 — prototype 과 동일."""
        job = _OVERALL_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "detected_heads" not in job:
            return jsonify({"ok": False, "message": "Stage A (스트림) 가 아직 끝나지 않았습니다."}), 400
        body = request.get_json(silent=True) or {}
        job["edit"] = {
            "added_heads": [tuple(p) for p in body.get("added_heads", [])],
            "deleted_indices": [int(i) for i in body.get("deleted_indices", [])],
            "zones": [tuple(z) for z in body.get("zones", [])],
            "branch_zones": [tuple(z) for z in body.get("branch_zones",
                                                        job.get("branch_zones") or [])],
        }
        ax, ay = body.get("alarm_x"), body.get("alarm_y")
        if ax is not None and ay is not None:
            try:
                job["alarm_xy"] = (float(ax), float(ay))
            except (TypeError, ValueError):
                pass
        return jsonify({"ok": True, "job_id": job_id,
                        "added": len(job["edit"]["added_heads"]),
                        "deleted": len(job["edit"]["deleted_indices"]),
                        "zones": len(job["edit"]["zones"])})

    @app.get("/api/remote30/overall/finalize_stream/<job_id>")
    def remote30_overall_finalize_stream(job_id: str):
        """Stage A 헤드 선정·테이블 생성 + Stage B 라이저 + Stage C stitch + Stage D emit_full_sdf SSE.

        run_stages_3_5 는 호출하지 않음 — 그 함수는 prototype 자체 SDF/CSV/zip 까지 만드는데,
        우리는 emit_full_sdf 가 새 SDF 를 생성하므로 중복. 대신 select_worst30_heads +
        build_input_tables 만 직접 호출하여 head_tables 객체를 얻는다.
        """
        job = _OVERALL_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "edit" not in job:
            return jsonify({"ok": False, "message": "finalize() 먼저 호출하세요."}), 400

        from remote30_prototype import select_worst30_heads, build_input_tables

        def _emit(payload: dict) -> str:
            return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

        def _gen():
            try:
                # ── Stage A 마무리: 헤드 선정 (사용자 edit 반영) + PipeTables 생성
                yield _emit({"type": "overall_progress", "phase": "stage_a_select",
                             "msg": "헤드 선정 (사용자 편집 반영)"})

                detected_pos = [tuple(d["pos"]) for d in job.get("detected_heads", [])]
                deleted = set(job["edit"]["deleted_indices"])
                manual_heads: list[tuple[float, float]] = [
                    p for i, p in enumerate(detected_pos) if i not in deleted
                ]
                manual_heads.extend(job["edit"]["added_heads"])

                selection = select_worst30_heads(
                    pipe_entities=job.get("pipe_ents", []),
                    layer_categories=job.get("layer_cat", {}),
                    manual_source=job.get("alarm_xy"),
                    manual_heads=manual_heads if (manual_heads or job["edit"]["deleted_indices"]
                                                  or job["edit"]["added_heads"]) else None,
                    zones=job["edit"]["zones"] if job["edit"]["zones"] else None,
                    branch_zones=job["edit"].get("branch_zones") or None,
                )
                yield _emit({"type": "overall_progress", "phase": "stage_a_select_done",
                             "heads": len(selection.heads),
                             "edges": len(selection.edges),
                             "nodes_in_subgraph": len(selection.nodes_in_subgraph)})

                from remote30_prototype import build_stage4_entities
                stage4_ents = build_stage4_entities(selection)
                yield _emit({
                    "type": "entities", "stage": 4, "entities": stage4_ents,
                    "summary": {
                        "selected_heads": len(selection.heads),
                        "subgraph_edges": len(selection.edges),
                        "off_tree_edges": sum(1 for e in stage4_ents if e.get("x")),
                        "source_pos": list(selection.source_pos) if selection.source_pos else None,
                        "source_kind": selection.source_kind,
                    },
                })
                yield _emit({"type": "stage", "stage": 4, "status": "done",
                             "label": f"30 헤드 선정 완료 — 헤드 {len(selection.heads)}개 / "
                                      f"경로 {len(selection.edges)} edge"})

                head_tables = build_input_tables(
                    selection,
                    pipe_entities=job.get("pipe_ents", []),
                    project_title=Path(job["dxf_path"]).stem,
                )
                yield _emit({"type": "overall_progress", "phase": "stage_a_tables_done",
                             "head_nodes": len(head_tables.nodes),
                             "head_pipes": len(head_tables.pipes),
                             "head_nozzles": len(head_tables.nozzles)})

                # ── 신축배관(FX) 검토 게이트 — prototype 과 동일한 2단 게이트.
                # 라이저/stitch/emit 은 fx/finalize_stream 으로 분리했다. head_tables 를
                # job state 에 저장 → 웹 편집기(FX 검토 패널) → POST .../fx/finalize.
                from core.remote30_constants import FX_SPEC_PROFILES, FX_DEFAULT_PROFILE
                job["head_tables"] = head_tables.as_dict()
                job["overall_project_stem"] = Path(job["dxf_path"]).stem
                job["fx_review"] = {
                    "equipment": head_tables.equipment,   # 헤드망 FX/AV — 전량, 편집 대상
                    "profiles": FX_SPEC_PROFILES,
                    "default_profile": FX_DEFAULT_PROFILE,
                }
                yield _emit({"type": "stage5_complete",
                             "tables": job["head_tables"],
                             "project_title": job["overall_project_stem"],
                             "fx_review": job["fx_review"]})

            except Exception as exc:  # noqa: BLE001
                import traceback
                err = {"type": "error", "message": str(exc)[:500],
                       "traceback": traceback.format_exc()[-1500:]}
                yield _emit(err)

        response = Response(_gen(), mimetype="text/event-stream")
        response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @app.post("/api/remote30/overall/fx/finalize/<job_id>")
    def remote30_overall_fx_finalize(job_id: str):
        """신축배관(FX) 검토 편집 결과 수신 (통합 플로우) — prototype 과 동일 컨벤션.

        body (JSON): equipment: [...] | null (null/생략 = 원본 그대로 emit)
        저장만 하고 실제 라이저·stitch·emit_full_sdf 는 fx/finalize_stream 에서 스트리밍.
        """
        job = _OVERALL_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "head_tables" not in job:
            return jsonify({"ok": False, "message": "stage5 (헤드망 테이블) 가 아직 끝나지 않았습니다."}), 400
        body = request.get_json(silent=True) or {}
        edited = body.get("equipment")
        job["fx_edit"] = edited if isinstance(edited, list) else None
        return jsonify({"ok": True, "job_id": job_id,
                        "edited": (len(job["fx_edit"]) if job["fx_edit"] is not None else 0),
                        "mode": ("edited" if job["fx_edit"] is not None else "original")})

    @app.get("/api/remote30/overall/fx/finalize_stream/<job_id>")
    def remote30_overall_fx_finalize_stream(job_id: str):
        """Stage B(라이저)~D(emit_full_sdf) SSE — fx/finalize() 호출 후 구독.

        head_tables 를 job state 에서 복원하고, FX 편집이 있으면 검증 후 equipment 교체.
        spec/profile/riser 는 job form 에서 결정적으로 재계산한다.
        """
        job = _OVERALL_JOBS.get(job_id)
        if not job:
            return jsonify({"ok": False, "message": f"unknown job_id {job_id}"}), 404
        if "head_tables" not in job:
            return jsonify({"ok": False, "message": "stage5 결과가 없습니다. finalize_stream 을 먼저 완료하세요."}), 400

        from remote30_prototype import PipeTables, _validate_edited_equipment
        from remote30_full_network import (
            zone_spec_from_form, profile_from_form, BuildingPressureProfile,
            build_riser, stitch_riser_and_heads, emit_full_sdf,
        )

        def _emit(payload: dict) -> str:
            return f"data: {json.dumps(payload, ensure_ascii=False)}\n\n"

        def _gen():
            try:
                head_tables = PipeTables.from_dict(job["head_tables"])
                # ── FX 편집 적용 (헤드망 equipment 만) + 검증
                fx_edit = job.get("fx_edit")
                if fx_edit is not None:
                    try:
                        new_equipment, warns = _validate_edited_equipment(fx_edit, head_tables.equipment)
                    except ValueError as _ve:
                        yield _emit({"type": "error", "message": str(_ve)})
                        return
                    for w in warns:
                        yield _emit(dict(w))
                    head_tables.equipment = new_equipment

                spec = zone_spec_from_form(job["spec_form"])
                # 압력표 우선순위: csv → xlsx → form JSON
                profile: BuildingPressureProfile | None = None
                if job.get("pressure_table_csv"):
                    profile = BuildingPressureProfile.from_csv(
                        Path(job["pressure_table_csv"]),
                        building_name=job["spec_form"].get("building_name", ""))
                elif job.get("pressure_table_xlsx"):
                    profile = BuildingPressureProfile.from_xlsx(
                        Path(job["pressure_table_xlsx"]),
                        building_name=job["spec_form"].get("building_name", ""))
                else:
                    profile = profile_from_form(job["spec_form"])

                # ── Stage B: 라이저
                yield _emit({"type": "overall_progress", "phase": "stage_b",
                             "msg": f"라이저 템플릿 생성 ({spec.zone_type.value})"})
                riser = build_riser(spec, profile)
                yield _emit({"type": "overall_progress", "phase": "stage_b_done",
                             "riser_nodes": len(riser.nodes),
                             "riser_pipes": len(riser.pipes),
                             "pumps": len(riser.pumps),
                             "valves": len(riser.valves)})

                # ── Stage C: stitch
                yield _emit({"type": "overall_progress", "phase": "stage_c",
                             "msg": "라이저 ↔ 헤드망 결합"})
                combined = stitch_riser_and_heads(riser, head_tables)
                yield _emit({"type": "overall_progress", "phase": "stage_c_done",
                             "combined_nodes": len(combined.nodes),
                             "combined_pipes": len(combined.pipes)})

                # ── Stage D: emit_full_sdf
                yield _emit({"type": "overall_progress", "phase": "stage_d",
                             "msg": "완성 SDF 직렬화 (PIPENET-native 후처리)"})
                out_sdf = Path(job["out_dir"]) / f"overall_{job_id}.sdf"
                emit_full_sdf(combined, out_sdf,
                              project_title=f"Remote 30 전체 — {spec.zone_type.value} {spec.target_floor}")

                yield _emit({"type": "overall_result",
                             "sdf": out_sdf.name, "job_id": job_id,
                             "nodes": len(combined.nodes),
                             "pipes": len(combined.pipes),
                             "pumps": len(combined.pumps),
                             "valves": len(combined.valves),
                             "nozzles": len(combined.nozzles)})

                # ── 통합 배관망 geometry 송신 — 클라이언트가 2D/Z/3D 토글로 시각화
                riser_labels = {str(n["label"]) for n in riser.nodes}
                yield _emit({
                    "type": "overall_geometry",
                    "av_node_label": riser.av_node_label,
                    "riser_labels": list(riser_labels),  # 라이저(Stage B)에 속하는 노드 라벨
                    "nodes": [
                        {"label": str(n["label"]),
                         "x": float(n.get("x", 0)), "y": float(n.get("y", 0)),
                         "z": float(n.get("elevation", 0)),
                         "io": n.get("io_node", "No")}
                        for n in combined.nodes
                    ],
                    "pipes": [
                        {"label": str(p.get("label", "")),
                         "in": str(p["in"]), "out": str(p["out"]),
                         "dia": p.get("dia", 0)}
                        for p in combined.pipes
                    ],
                    "pumps": [
                        {"label": str(p["label"]), "in": str(p["in"]), "out": str(p["out"]),
                         "library_pump": p.get("library_pump", "")}
                        for p in combined.pumps
                    ],
                    "valves": [
                        {"label": str(v["label"]), "in": str(v["in"]), "out": str(v["out"]),
                         "target_pa": v["target_value"]}
                        for v in combined.valves
                    ],
                })

                # ── 완료 신호 — prototype 의 finalize handler 가 done 또는 error 이벤트 시 버튼을 "✓ 완료" 로 전환
                yield _emit({"type": "done", "message": "전체 배관망 SDF 생성 완료",
                             "job_id": job_id, "sdf": out_sdf.name})

            except Exception as exc:  # noqa: BLE001
                import traceback
                err = {"type": "error", "message": str(exc)[:500],
                       "traceback": traceback.format_exc()[-1500:]}
                yield _emit(err)

        response = Response(_gen(), mimetype="text/event-stream")
        response.headers["Cache-Control"] = "no-cache"
        response.headers["X-Accel-Buffering"] = "no"
        return response

    @app.post("/api/remote30/overall/parse-system-diagram")
    def remote30_overall_parse_system_diagram():
        """계통도 DXF 업로드 → 텍스트 라벨 파싱 → JSON floors 배열 반환.

        Form fields:
            system_diagram_file: 계통도 .dxf (required)
            default_height_m: 표준 층고 (선택, 기본 2.9)
            roof_height_m: 옥상층 층고 (선택, 기본 6.0)

        Response:
            {ok, building_name, floors: [{floor_label, height_m, head_drop_m, note}, ...]}
        """
        from remote30_full_network import parse_system_diagram_dxf
        try:
            dxf_path = _save_upload("system_diagram_file", {".dxf", ".dwg"}, required=True)
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        default_h = float(request.form.get("default_height_m", "2.9") or "2.9")
        roof_h = float(request.form.get("roof_height_m", "6.0") or "6.0")
        try:
            profile = parse_system_diagram_dxf(dxf_path,
                                                default_height_m=default_h,
                                                roof_height_m=roof_h)
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)
        return jsonify({
            "ok": True,
            "building_name": profile.building_name,
            "floors": [
                {"floor_label": r.floor_label, "height_m": r.height_m,
                 "head_drop_m": r.head_drop_m, "note": r.note}
                for r in profile.floors
            ],
            "n_floors": len(profile.floors),
        })

    @app.get("/api/remote30/overall/result/<job_id>/<path:filename>")
    def remote30_overall_result(job_id: str, filename: str):
        return _serve_run_file(OVERALL_OUTPUT_DIR, job_id, filename)
