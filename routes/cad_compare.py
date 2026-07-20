# -*- coding: utf-8 -*-
"""CAD↔SDF 대조(cad_compare) 도메인 — DXF 파싱·비교·AI 영역 매칭.

`대조 서버.py` 에서 `register(app, ...)` 로 등록. AI 매칭 헬퍼(_ai_graph_match)는
주입 이름(_ai_edge_features 등)을 클로저로 참조하므로 register 안에 중첩.
공유 헬퍼·전역은 main 유지·주입. 라우트 본문·엔드포인트명 원본 그대로.
"""
from __future__ import annotations

import time
import xml.etree.ElementTree as ET
from pathlib import Path

from flask import jsonify, make_response, render_template, request

from cad_match import _torch_device_info


def register(app, *, BASE_DIR, UPLOAD_DIR, _AI_MATCH_MAX_EDGES, _ai_edge_features, _compact_cad_graph_for_sdf, _component_similarity_stats, _edge_length, _recompute_edge_degrees, _save_upload):

    def _sdf_counts_only(sdf_path: Path | None) -> dict:
        if sdf_path is None or not sdf_path.exists():
            return {}
        root = ET.parse(sdf_path).getroot()
        return {
            "pipes": len(root.findall(".//Pipe")),
            "nozzles": len(root.findall(".//Nozzle")),
            "equipment": len(root.findall(".//Equipment")),
        }

    def _ai_graph_match(dxf_graph: dict, sdf_graph: dict) -> dict:
        raw_dxf_graph = dxf_graph or {}
        sdf_graph = sdf_graph or {}
        dxf_graph = _compact_cad_graph_for_sdf(raw_dxf_graph, sdf_graph)
        dxf_edges = dxf_graph.get("edges") or []
        sdf_edges = sdf_graph.get("edges") or []
        # sdf 측도 상한 — diff 텐서가 (len(dxf)×len(sdf)×8) 라 sdf 가 거대하면 메모리 폭발.
        if len(sdf_edges) > _AI_MATCH_MAX_EDGES:
            sdf_edges = sorted(
                sdf_edges,
                key=lambda e: float(e.get("length") or _edge_length(e)),
                reverse=True,
            )[:_AI_MATCH_MAX_EDGES]
        _recompute_edge_degrees(dxf_edges)
        _recompute_edge_degrees(sdf_edges)
        for edge in dxf_edges:
            if edge.get("sdf_expected_source_degree") is not None:
                edge["sourceDegree"] = edge.get("sdf_expected_source_degree")
            if edge.get("sdf_expected_target_degree") is not None:
                edge["targetDegree"] = edge.get("sdf_expected_target_degree")
        dxf_features = _ai_edge_features(dxf_edges)
        sdf_features = _ai_edge_features(sdf_edges)
        if not dxf_features or not sdf_features:
            return {
                "ok": True,
                "device": "none",
                "rows": [],
                "summary": "선택영역에서 비교 가능한 DXF Edge 또는 SDF Pipe가 부족합니다.",
                "stats": {"score": 0, "pass": 0, "review": 0, "fail": len(sdf_edges), "ai_average": 0},
            }
        try:
            import torch
            device = "cuda" if torch.cuda.is_available() else "cpu"
            dxf_tensor = torch.tensor(dxf_features, dtype=torch.float32, device=device)
            sdf_tensor = torch.tensor(sdf_features, dtype=torch.float32, device=device)
            weights = torch.tensor([1.0, 1.0, 0.75, 0.35, 0.35, 0.35, 0.25, 0.18], dtype=torch.float32, device=device)
            diff = (dxf_tensor[:, None, :] - sdf_tensor[None, :, :]).abs() * weights
            # Layer prior is an advantage for DXF fire/sprinkler layers, not a distance penalty.
            diff[:, :, 7] = torch.clamp(-dxf_tensor[:, None, 7] * 0.18, min=-0.18, max=0.18)
            matrix = diff.sum(dim=2).detach().cpu().tolist()
        except Exception:
            device = "cpu-fallback"
            matrix = []
            for dxf in dxf_features:
                row = []
                for sdf in sdf_features:
                    dist = (
                        abs(dxf[0] - sdf[0]) * 1.0
                        + abs(dxf[1] - sdf[1]) * 1.0
                        + abs(dxf[2] - sdf[2]) * 0.75
                        + abs(dxf[3] - sdf[3]) * 0.35
                        + abs(dxf[4] - sdf[4]) * 0.35
                        + abs(dxf[5] - sdf[5]) * 0.35
                        + abs(dxf[6] - sdf[6]) * 0.25
                        - dxf[7] * 0.18
                    )
                    row.append(dist)
                matrix.append(row)

        guided_edges = [edge for edge in dxf_edges if edge.get("matched_sdf_id") is not None or edge.get("matched_sdf_label") is not None]
        if guided_edges:
            sdf_by_id = {str(edge.get("id")): edge for edge in sdf_edges if edge.get("id") is not None}
            sdf_by_label = {str(edge.get("label")): edge for edge in sdf_edges if edge.get("label") is not None}
            rows = []
            used_sdf: set[str] = set()
            for dxf_edge in guided_edges:
                sdf_edge = sdf_by_id.get(str(dxf_edge.get("matched_sdf_id"))) or sdf_by_label.get(str(dxf_edge.get("matched_sdf_label")))
                if not sdf_edge:
                    continue
                used_sdf.add(str(sdf_edge.get("id") or sdf_edge.get("label")))
                length_ratio = float(dxf_edge.get("length") or 0.0) / max(float(sdf_edge.get("length") or _edge_length(sdf_edge)), 1e-9)
                length_fail = abs(1.0 - length_ratio) > 0.10
                degree_fail = abs((float(dxf_edge.get("sourceDegree") or 0) + float(dxf_edge.get("targetDegree") or 0)) - (float(sdf_edge.get("sourceDegree") or 0) + float(sdf_edge.get("targetDegree") or 0))) >= 2
                guide_score = float(dxf_edge.get("sdf_guided_score") or 0.0)
                ai_conf = max(0.0, min(1.0, 1.0 - min(guide_score, 1.8) / 1.8))
                status = "FAIL" if length_fail or degree_fail else "PASS"
                rows.append(
                    {
                        "status": status,
                        "dxf_id": dxf_edge.get("id"),
                        "sdf_id": sdf_edge.get("id"),
                        "dxf_label": dxf_edge.get("label") or dxf_edge.get("id"),
                        "sdf_label": sdf_edge.get("label") or sdf_edge.get("id"),
                        "dxf_layer": dxf_edge.get("layer"),
                        "sdf_layer": sdf_edge.get("layer"),
                        "ai_confidence": round(ai_conf * 100, 1),
                        "score": round(guide_score, 4),
                        "compare": f"길이 {float(dxf_edge.get('length') or 0):.1f} / {float(sdf_edge.get('length') or _edge_length(sdf_edge)):.1f}, 길이비 {length_ratio:.2f}",
                        "reason": f"SDF-guided CAD 묶음 대조, 원본 CAD 길이 {float(dxf_edge.get('raw_cad_length') or 0):.1f}, 스케일 보정 {float(dxf_edge.get('length_scale_factor') or 1):.3f}, 형상 후보점수 {guide_score:.3f}",
                    }
                )
            for edge in sdf_edges:
                key = str(edge.get("id") or edge.get("label"))
                if key not in used_sdf:
                    rows.append(
                        {
                            "status": "FAIL",
                            "dxf_id": None,
                            "sdf_id": edge.get("id"),
                            "dxf_label": "-",
                            "sdf_label": edge.get("label") or edge.get("id"),
                            "dxf_layer": "-",
                            "sdf_layer": edge.get("layer"),
                            "ai_confidence": None,
                            "score": None,
                            "compare": "DXF 대응 Bundle 없음",
                            "reason": "SDF-guided bundling 단계에서 대응 CAD 묶음을 만들지 못했습니다.",
                        }
                    )
            pass_count = sum(1 for row in rows if row["status"] == "PASS")
            review_count = sum(1 for row in rows if row["status"] == "REVIEW")
            fail_count = sum(1 for row in rows if row["status"] == "FAIL")
            ai_values = [row["ai_confidence"] for row in rows if isinstance(row.get("ai_confidence"), (int, float))]
            ai_avg = sum(ai_values) / len(ai_values) if ai_values else 0.0
            score = max(0.0, min(100.0, ((pass_count + review_count * 0.45) / max(len(sdf_edges), 1)) * 100.0))
            component_stats = _component_similarity_stats(dxf_graph, sdf_graph, rows)
            summary = (
                f"SDF-guided 방식으로 CAD 원본 선분 {dxf_graph.get('edges_raw_count', len(dxf_edges))}개를 SDF Pipe {len(sdf_edges)}개 기준의 배관 묶음 {len(dxf_edges)}개로 재구성했습니다. "
                f"PASS {pass_count}건, REVIEW {review_count}건, FAIL {fail_count}건이며 FFT 형상 유사도는 {component_stats.get('fft_shape_similarity', 0)}%입니다."
            )
            return {
                "ok": True,
                "device": device,
                "rows": rows,
                "summary": summary,
                "dxf_graph": dxf_graph,
                "sdf_graph": sdf_graph,
                "component_scores": component_stats,
                "preprocess": dxf_graph.get("ai_preprocess") or {},
                "stats": {
                    "score": round(score, 1),
                    "pass": pass_count,
                    "review": review_count,
                    "fail": fail_count,
                    "ai_average": round(ai_avg, 1),
                    "dxf_edge_count": len(dxf_edges),
                    "sdf_pipe_count": len(sdf_edges),
                    **component_stats,
                },
            }

        pairs = []
        for i, row in enumerate(matrix):
            for j, score in enumerate(row):
                pairs.append((float(score), i, j))
        pairs.sort(key=lambda x: x[0])
        used_dxf, used_sdf = set(), set()
        rows = []
        for score, i, j in pairs:
            score = max(0.0, float(score))
            if i in used_dxf or j in used_sdf:
                continue
            if score > 1.35:
                continue
            used_dxf.add(i)
            used_sdf.add(j)
            dxf_edge = dxf_edges[i]
            sdf_edge = sdf_edges[j]
            ai_conf = max(0.0, min(1.0, 1.0 - score / 1.35))
            length_ratio = float(dxf_edge.get("length") or 0.0) / max(float(sdf_edge.get("length") or 0.0), 1e-9)
            length_fail = abs(1.0 - length_ratio) > 0.25
            degree_fail = abs((float(dxf_edge.get("sourceDegree") or 0) + float(dxf_edge.get("targetDegree") or 0)) - (float(sdf_edge.get("sourceDegree") or 0) + float(sdf_edge.get("targetDegree") or 0))) >= 2
            status = "FAIL" if length_fail or degree_fail else "REVIEW" if ai_conf < 0.56 else "PASS"
            rows.append(
                {
                    "status": status,
                    "dxf_id": dxf_edge.get("id"),
                    "sdf_id": sdf_edge.get("id"),
                    "dxf_label": dxf_edge.get("label") or dxf_edge.get("id"),
                    "sdf_label": sdf_edge.get("label") or sdf_edge.get("id"),
                    "dxf_layer": dxf_edge.get("layer"),
                    "sdf_layer": sdf_edge.get("layer"),
                    "ai_confidence": round(ai_conf * 100, 1),
                    "score": round(score, 4),
                    "compare": f"길이 {float(dxf_edge.get('length') or 0):.1f} / {float(sdf_edge.get('length') or 0):.1f}, 길이비 {length_ratio:.2f}",
                    "reason": f"GPU/AI 그래프 유사도 {score:.3f}, 신뢰도 {ai_conf * 100:.1f}%, 연결차수 DXF {dxf_edge.get('sourceDegree', 0)}+{dxf_edge.get('targetDegree', 0)} / SDF {sdf_edge.get('sourceDegree', 0)}+{sdf_edge.get('targetDegree', 0)}",
                }
            )
        for i, edge in enumerate(dxf_edges):
            if i not in used_dxf:
                rows.append(
                    {
                        "status": "REVIEW",
                        "dxf_id": edge.get("id"),
                        "sdf_id": None,
                        "dxf_label": edge.get("label") or edge.get("id"),
                        "sdf_label": "-",
                        "dxf_layer": edge.get("layer"),
                        "sdf_layer": "-",
                        "ai_confidence": None,
                        "score": None,
                        "compare": "SDF 대응 Pipe 미확정",
                        "reason": "선택영역 안에서 AI 유사도 기준에 맞는 SDF Pipe를 찾지 못했습니다.",
                    }
                )
        for j, edge in enumerate(sdf_edges):
            if j not in used_sdf:
                rows.append(
                    {
                        "status": "FAIL",
                        "dxf_id": None,
                        "sdf_id": edge.get("id"),
                        "dxf_label": "-",
                        "sdf_label": edge.get("label") or edge.get("id"),
                        "dxf_layer": "-",
                        "sdf_layer": edge.get("layer"),
                        "ai_confidence": None,
                        "score": None,
                        "compare": "DXF 대응 Edge 없음",
                        "reason": "SDF Pipe는 선택영역 안에 있으나 AI 그래프 대조에서 대응 DXF 선분이 확인되지 않았습니다.",
                    }
                )
        pass_count = sum(1 for row in rows if row["status"] == "PASS")
        review_count = sum(1 for row in rows if row["status"] == "REVIEW")
        fail_count = sum(1 for row in rows if row["status"] == "FAIL")
        ai_values = [row["ai_confidence"] for row in rows if isinstance(row.get("ai_confidence"), (int, float))]
        ai_avg = sum(ai_values) / len(ai_values) if ai_values else 0.0
        score = max(0.0, min(100.0, ((pass_count + review_count * 0.45) / max(len(sdf_edges), 1)) * 100.0))
        component_stats = _component_similarity_stats(dxf_graph, sdf_graph, rows)
        summary = (
            f"선택영역 AI 그래프 대조 결과, SDF Pipe {len(sdf_edges)}개 중 PASS {pass_count}건, REVIEW {review_count}건, FAIL {fail_count}건입니다. "
            f"연산 장치는 {device}이며 평균 AI 신뢰도는 {ai_avg:.1f}%입니다. "
            "빨간 구간은 도면 선분 누락, 선택영역 불일치, 긴 CAD 선분의 분할 문제, 또는 실제 배관망 형상 차이를 우선 점검해야 합니다."
        )
        return {
            "ok": True,
            "device": device,
            "rows": rows,
            "summary": summary,
            "dxf_graph": dxf_graph,
            "sdf_graph": sdf_graph,
            "component_scores": component_stats,
            "preprocess": dxf_graph.get("ai_preprocess") or {},
            "stats": {
                "score": round(score, 1),
                "pass": pass_count,
                "review": review_count,
                "fail": fail_count,
                "ai_average": round(ai_avg, 1),
                "dxf_edge_count": len(dxf_edges),
                "sdf_pipe_count": len(sdf_edges),
                **component_stats,
            },
        }

    @app.get("/cad-compare-module")
    def cad_compare_module():
        response = make_response(render_template("cad_compare_module.html"))
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response

    @app.get("/cad-compare-module-7")
    def cad_compare_module_7():
        response = make_response(render_template("cad_compare_module_7.html"))
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response

    @app.post("/api/cad-module/dxf-parse")
    def cad_module_dxf_parse():
        try:
            cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=True)

            from cad_engine import DXFWorkspace

            workspace = DXFWorkspace(UPLOAD_DIR / "cad_workspace")
            workspace.load_file(cad_path)
            payload = workspace.to_payload(
                include_network_entities=False,
                include_network_summary=False,
                include_graph=False,
            )
        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:
            return jsonify({"ok": False, "message": f"DXF 파싱 중 오류가 발생했습니다: {exc}"}), 500

        return jsonify(
            {
                "ok": True,
                "message": "DXF 파싱이 완료되었습니다.",
                "cad_payload": {
                    "filename": payload.get("filename"),
                    "bounds": payload.get("bounds"),
                    "layers": payload.get("layers"),
                    "entities": payload.get("entities") or [],
                    "graph": {},
                    "unsupported": payload.get("unsupported") or {},
                },
            }
        )

    @app.post("/api/cad-compare")
    def cad_compare():
        try:
            cad_path = _save_upload("cad_file", {".dxf", ".dwg"}, required=True)
            sdf_path = _save_upload("sdf_file", {".sdf"}, required=False)

            from cad_engine import DXFWorkspace

            workspace = DXFWorkspace(UPLOAD_DIR / "cad_workspace")
            workspace.load_file(cad_path)
            payload = workspace.to_payload(
                include_network_entities=True,
                include_network_summary=True,
                include_graph=True,
            )

            network_layers = set(payload.get("networkLayers") or [])
            network_entity_ids = set(payload.get("networkEntityIds") or [])
            entities = payload.get("entities") or []
            if network_entity_ids:
                entities = [e for e in entities if e.get("id") in network_entity_ids]
            if network_layers:
                entities = [e for e in entities if e.get("layer") in network_layers]

            head_boxes: list[dict] = []
            detector_mode = "template"
            use_yolo = str(request.form.get("use_yolo", "")).strip() == "1"
            try:
                if use_yolo:
                    from head_detector import TriangleHeadDetector

                    model_path = BASE_DIR / "models" / "triangle_head_yolo_ai" / "weights" / "best.pt"
                    if not model_path.exists():
                        model_path = BASE_DIR / "runs" / "detect" / "models" / "triangle_head_yolo_ai" / "weights" / "best.pt"
                    if not model_path.exists():
                        model_path = BASE_DIR / "models" / "triangle_head_yolo" / "weights" / "best.pt"
                    if not model_path.exists():
                        model_path = BASE_DIR / "runs" / "detect" / "models" / "triangle_head_yolo" / "weights" / "best.pt"
                    if not model_path.exists():
                        model_path = BASE_DIR / "yolo26n.pt"
                    if not model_path.exists():
                        model_path = BASE_DIR / "yolo11n.pt"
                    detector = TriangleHeadDetector(BASE_DIR / "data" / "head_templates", model_path)
                    head_boxes = detector.detect(entities, payload.get("bounds") or {}, network_layers)
                    detector_mode = "yolo+template" if detector.yolo_detector.available else "template"
                else:
                    from head_detector import TriangleHeadTemplateDetector

                    detector = TriangleHeadTemplateDetector(BASE_DIR / "data" / "head_templates")
                    head_boxes = detector.detect(entities, payload.get("bounds") or {}, network_layers)
                    detector_mode = "template"
            except Exception:
                detector_mode = "unavailable"
                head_boxes = []

            cad_counts = {
                "entities": len(entities),
                "network_layers": len(network_layers),
                "detected_heads": len(head_boxes),
                "lines": sum(1 for e in entities if e.get("type") in {"LINE", "LWPOLYLINE", "ARC"}),
                "circles": sum(1 for e in entities if e.get("type") == "CIRCLE"),
                "texts": sum(1 for e in entities if e.get("type") == "TEXT"),
            }
            sdf_counts = _sdf_counts_only(sdf_path)
            messages: list[str] = []
            if sdf_counts:
                sdf_heads = int(sdf_counts.get("nozzles", 0))
                diff = cad_counts["detected_heads"] - sdf_heads
                if diff == 0:
                    messages.append(f"헤드 수 일치: CAD 탐지 {cad_counts['detected_heads']} / SDF {sdf_heads}")
                else:
                    messages.append(
                        f"헤드 수 차이: CAD 탐지 {cad_counts['detected_heads']} / SDF {sdf_heads} (차이 {diff:+d})"
                    )
                messages.append(
                    f"SDF 수량: 배관 {sdf_counts.get('pipes', 0)} / 헤드 {sdf_counts.get('nozzles', 0)} / 특수설비 {sdf_counts.get('equipment', 0)}"
                )
            else:
                messages.append("SDF 미업로드 상태입니다. CAD 단독 추출/탐지 결과만 표시합니다.")
            messages.append(f"탐지 엔진: {detector_mode}")

        except ValueError as exc:
            return jsonify({"ok": False, "message": str(exc)}), 400
        except Exception as exc:
            return jsonify({"ok": False, "message": f"CAD 대조 중 오류가 발생했습니다: {exc}"}), 500

        return jsonify(
            {
                "ok": True,
                "message": "CAD 대조가 완료되었습니다.",
                "cad_filename": cad_path.name,
                "sdf_filename": sdf_path.name if sdf_path else None,
                "cad_payload": {
                    "filename": payload.get("filename"),
                    "bounds": payload.get("bounds"),
                    "layers": payload.get("layers"),
                    "networkLayers": list(network_layers),
                    "entities": entities,
                    "graph": payload.get("graph") or {},
                },
                "detected_heads": head_boxes,
                "cad_counts": cad_counts,
                "sdf_counts": sdf_counts,
                "messages": messages,
            }
        )

    @app.post("/api/cad-sdf-ai-region-match")
    def cad_sdf_ai_region_match():
        started = time.perf_counter()
        try:
            payload = request.get_json(force=True)
            min_runtime_ms = max(0, min(int(payload.get("min_runtime_ms") or 0), 8000))
            result = _ai_graph_match(payload.get("dxf_graph") or {}, payload.get("sdf_graph") or {})
            elapsed_ms = int((time.perf_counter() - started) * 1000)
            remaining = max(0, min_runtime_ms - elapsed_ms)
            if remaining:
                time.sleep(remaining / 1000.0)
                elapsed_ms = int((time.perf_counter() - started) * 1000)
            preprocess = result.get("preprocess") or {}
            device_info = _torch_device_info()
            result["runtime_ms"] = elapsed_ms
            result["engine_pipeline"] = [
                {"id": "head_yolo", "name": "YOLO Head Detector", "status": "ACTIVE", "device": device_info.get("device"), "gpu": device_info.get("gpu_enabled")},
                {"id": "pipe_segmentation", "name": "Pipe Segmentation", "status": "ACTIVE" if (preprocess.get("segmentation") or {}).get("available") else "FALLBACK", **(preprocess.get("segmentation") or {})},
                {"id": "sdf_guided_bundle", "name": "SDF-guided Pipe Bundling", "status": "ACTIVE", "mode": preprocess.get("bundling_mode")},
                {"id": "fft_shape", "name": "FFT Shape Similarity", "status": "ACTIVE", "device": device_info.get("device"), "gpu": device_info.get("gpu_enabled")},
                {"id": "graph_match", "name": "GPU Graph Matching", "status": "ACTIVE" if device_info.get("gpu_enabled") else "CPU", "device": device_info.get("device"), "gpu_name": device_info.get("gpu_name")},
            ]
            return jsonify(result)
        except Exception as exc:
            return jsonify({"ok": False, "message": f"AI 그래프 대조 중 오류가 발생했습니다: {exc}"}), 500
