# -*- coding: utf-8 -*-
"""공유 코어 — CAD↔SDF AI 매칭/유사도/학습프로파일/세그멘테이션 상태.

_torch_device_info·_cad_layer_weight 포함 — 이전 라우트 분리 때 두 헬퍼가
라우트로 분산되면서 main 측 CAD 함수가 런타임 NameError 를 내던 것을
한 모듈로 다시 모아 해소. main · routes/* 모두 여기서 flat import.
"""
from __future__ import annotations

import json
import math
from pathlib import Path
from dxf_geometry import (_angle_delta, _bbox, _edge_angle, _edge_length, _edge_points, _graph_bbox_from_edges, _node_key, _normalize_layer_name, _segments_from_points)

BASE_DIR = Path(__file__).resolve().parent.parent


CAD_SDF_LEARNING_PROFILE_PATH = BASE_DIR / "data" / "cad_sdf_learning_profile.json"
PIPE_SEGMENTATION_MODEL_CANDIDATES = [
    BASE_DIR / "models" / "pipe_segmentation" / "weights" / "best.pt",
    BASE_DIR / "models" / "pipe_segmentation.pt",
    BASE_DIR / "runs" / "segment" / "pipe_segmentation" / "weights" / "best.pt",
    BASE_DIR / "yolo11n-seg.pt",
    BASE_DIR / "yolo26n-seg.pt",
]
_AI_MATCH_MAX_EDGES = 2000


def _pipe_segmentation_engine_status() -> dict:
    model_path = next((path for path in PIPE_SEGMENTATION_MODEL_CANDIDATES if path.exists()), None)
    device_info = _torch_device_info()
    if not model_path:
        return {
            "name": "Pipe Segmentation",
            "available": False,
            "mode": "sdf_guided_segmentation_proxy",
            "model_path": None,
            "message": "학습된 배관 세그멘테이션 가중치가 없어 SDF-guided 선분 묶음화 엔진으로 대체했습니다.",
            **device_info,
        }
    try:
        from ultralytics import YOLO

        # Load once per request to verify the trained segmentation weight is usable.
        YOLO(str(model_path))
        return {
            "name": "Pipe Segmentation",
            "available": True,
            "mode": "trained_ultralytics_segmentation",
            "model_path": str(model_path),
            "message": "학습된 세그멘테이션 가중치를 로드했습니다. DXF 벡터 그래프는 SDF-guided bundle 단계와 함께 사용됩니다.",
            **device_info,
        }
    except Exception as exc:
        return {
            "name": "Pipe Segmentation",
            "available": False,
            "mode": "sdf_guided_segmentation_proxy",
            "model_path": str(model_path),
            "message": f"세그멘테이션 가중치 로드 실패로 SDF-guided 선분 묶음화 엔진으로 대체했습니다: {exc}",
            **device_info,
        }

def _load_cad_sdf_learning_profile() -> dict:
    if not CAD_SDF_LEARNING_PROFILE_PATH.exists():
        return {}
    try:
        return json.loads(CAD_SDF_LEARNING_PROFILE_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _write_cad_sdf_learning_profile(profile: dict) -> None:
    CAD_SDF_LEARNING_PROFILE_PATH.parent.mkdir(parents=True, exist_ok=True)
    CAD_SDF_LEARNING_PROFILE_PATH.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")

def _mark_similar_cad_pipe_entities(cad: dict, sdf: dict) -> dict:
    learning_profile = _load_cad_sdf_learning_profile()
    cad_entities = cad.get("drawing_entities") or []
    sdf_pipes = sdf.get("pipes") or []
    cad_points = [{"x": x, "y": y} for ent in cad_entities for x, y in (ent.get("points") or [])]
    sdf_points = [{"x": x, "y": y} for pipe in sdf_pipes for x, y in (pipe.get("path") or [])]
    cad_box = _bbox(cad_points)
    sdf_box = _bbox(sdf_points)
    if not cad_box or not sdf_box:
        return {"matched_entity_ids": [], "matched_count": 0, "threshold": 0.09}

    sdf_segments = []
    for pipe in sdf_pipes:
        sdf_segments.extend(_segments_from_points(pipe.get("path") or [], sdf_box, str(pipe.get("label", "")), pipe.get("label")))
    cad_segments = []
    for ent in cad_entities:
        layer_weight = float(ent.get("layer_weight", _cad_layer_weight(ent.get("layer"), learning_profile)) or 0.0)
        if learning_profile and layer_weight <= -2.5:
            continue
        ent_segments = _segments_from_points(ent.get("points") or [], cad_box, str(ent.get("id", "")), ent.get("layer"))
        for seg in ent_segments:
            seg["layer_weight"] = layer_weight
            seg["layer"] = ent.get("layer")
        cad_segments.extend(ent_segments)
    if not sdf_segments or not cad_segments:
        return {"matched_entity_ids": [], "matched_count": 0, "threshold": 0.09}

    matched: dict[str, float] = {}
    for sdf_seg in sdf_segments:
        best_id = None
        best_score = float("inf")
        for cad_seg in cad_segments:
            mid_dist = math.hypot(sdf_seg["mid"][0] - cad_seg["mid"][0], sdf_seg["mid"][1] - cad_seg["mid"][1])
            angle_penalty = _angle_delta(sdf_seg["angle"], cad_seg["angle"]) / math.pi
            len_ratio = abs(math.log(max(cad_seg["length"], 1e-6) / max(sdf_seg["length"], 1e-6)))
            layer_weight = float(cad_seg.get("layer_weight") or 0.0)
            score = mid_dist + angle_penalty * 0.18 + min(len_ratio, 2.0) * 0.05
            if layer_weight > 0:
                score -= min(layer_weight, 5.0) * 0.012
            elif layer_weight < 0:
                score += abs(layer_weight) * 0.04
            if score < best_score:
                best_score = score
                best_id = cad_seg["source_id"]
        if best_id and best_score <= 0.09:
            matched[best_id] = min(best_score, matched.get(best_id, best_score))

    for ent in cad_entities:
        sid = str(ent.get("id"))
        if sid in matched:
            ent["similar_to_sdf"] = True
            ent["similarity_score"] = round(matched[sid], 4)

    return {
        "matched_entity_ids": sorted(matched, key=lambda key: matched[key])[:2000],
        "matched_count": len(matched),
        "threshold": 0.09,
        "learning_profile_applied": bool(learning_profile),
    }

def _ai_edge_features(edges: list[dict]) -> list[list[float]]:
    if not edges:
        return []
    pts = []
    for edge in edges:
        for p in edge.get("points") or []:
            try:
                pts.append((float(p.get("x", 0.0)), float(p.get("y", 0.0))))
            except Exception:
                continue
    if not pts:
        pts = [(0.0, 0.0), (100.0, 100.0)]
    min_x, max_x = min(x for x, _ in pts), max(x for x, _ in pts)
    min_y, max_y = min(y for _, y in pts), max(y for _, y in pts)
    w = max(max_x - min_x, 1e-9)
    h = max(max_y - min_y, 1e-9)
    diag = max(math.hypot(w, h), 1e-9)
    _profile = _load_cad_sdf_learning_profile()
    rows = []
    for edge in edges:
        start = edge.get("start") or {}
        end = edge.get("end") or {}
        sx, sy = float(start.get("x", 0.0)), float(start.get("y", 0.0))
        ex, ey = float(end.get("x", 0.0)), float(end.get("y", 0.0))
        mx = (((sx + ex) / 2.0) - min_x) / w
        my = (((sy + ey) / 2.0) - min_y) / h
        length = float(edge.get("length") or math.hypot(ex - sx, ey - sy)) / diag
        angle = math.atan2((ey - sy) / h, (ex - sx) / w)
        degree = (float(edge.get("sourceDegree") or 0.0) + float(edge.get("targetDegree") or 0.0)) / 8.0
        bore = float(edge.get("bore") or 0.0) / 200.0
        layer_prior = _cad_layer_weight(edge.get("layer"), _profile) / 5.0
        rows.append([mx, my, length, math.cos(angle), math.sin(angle), degree, bore, layer_prior])
    return rows

def _recompute_edge_degrees(edges: list[dict]) -> None:
    if not edges:
        return
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(edges)
    tolerance = max(math.hypot(max_x - min_x, max_y - min_y) * 0.006, 20.0)
    degree: dict[str, int] = {}
    keys: list[tuple[str, str]] = []
    for edge in edges:
        pts = _edge_points(edge)
        if len(pts) < 2:
            keys.append(("", ""))
            continue
        sk = _node_key(pts[0], tolerance)
        tk = _node_key(pts[-1], tolerance)
        degree[sk] = degree.get(sk, 0) + 1
        degree[tk] = degree.get(tk, 0) + 1
        keys.append((sk, tk))
    for edge, (sk, tk) in zip(edges, keys):
        edge["sourceDegree"] = degree.get(sk, 0)
        edge["targetDegree"] = degree.get(tk, 0)

def _merge_collinear_cad_edges(edges: list[dict]) -> list[dict]:
    if len(edges) <= 1:
        return edges
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(edges)
    tolerance = max(math.hypot(max_x - min_x, max_y - min_y) * 0.005, 1.0)
    work = [dict(edge) for edge in edges if _edge_length(edge) > tolerance * 0.15]

    for _ in range(4):
        endpoint_map: dict[str, list[int]] = {}
        for idx, edge in enumerate(work):
            pts = _edge_points(edge)
            if len(pts) < 2:
                continue
            endpoint_map.setdefault(_node_key(pts[0], tolerance), []).append(idx)
            endpoint_map.setdefault(_node_key(pts[-1], tolerance), []).append(idx)

        merged_idx: set[int] = set()
        merged_edges: list[dict] = []
        changed = False
        for key, idxs in endpoint_map.items():
            idxs = [idx for idx in idxs if idx not in merged_idx]
            if len(idxs) != 2:
                continue
            a, b = work[idxs[0]], work[idxs[1]]
            if str(a.get("layer") or "") != str(b.get("layer") or ""):
                continue
            if _angle_delta(_edge_angle(a), _edge_angle(b)) > 0.16:
                continue
            pa, pb = _edge_points(a), _edge_points(b)
            if len(pa) < 2 or len(pb) < 2:
                continue
            pts = pa + pb
            cx = sum(pt["x"] for pt in pts) / len(pts)
            cy = sum(pt["y"] for pt in pts) / len(pts)
            angle = _edge_angle(a)
            ordered = sorted(pts, key=lambda pt: (pt["x"] - cx) * math.cos(angle) + (pt["y"] - cy) * math.sin(angle))
            merged = dict(a)
            member_ids = []
            for source in (a, b):
                member_ids.extend(source.get("member_ids") or [source.get("id") or source.get("label")])
            merged["id"] = f"{a.get('id') or a.get('label')}-{b.get('id') or b.get('label')}"
            merged["label"] = f"{a.get('label') or a.get('id')}+{b.get('label') or b.get('id')}"
            merged["points"] = [ordered[0], ordered[-1]]
            merged["start"] = ordered[0]
            merged["end"] = ordered[-1]
            merged["length"] = _edge_length(merged)
            merged["merged_count"] = int(a.get("merged_count") or 1) + int(b.get("merged_count") or 1)
            merged["member_ids"] = [str(x) for x in member_ids if x]
            merged_edges.append(merged)
            merged_idx.update(idxs)
            changed = True
        work = [edge for idx, edge in enumerate(work) if idx not in merged_idx] + merged_edges
        if not changed:
            break
    _recompute_edge_degrees(work)
    return work

def _compact_cad_graph_for_sdf(dxf_graph: dict, sdf_graph: dict) -> dict:
    raw_edges = [dict(edge) for edge in (dxf_graph.get("edges") or [])]
    sdf_edges = sdf_graph.get("edges") or []
    if not raw_edges or not sdf_edges:
        return dxf_graph
    _recompute_edge_degrees(sdf_edges)
    merged = _merge_collinear_cad_edges(raw_edges)
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(merged)
    diag = max(math.hypot(max_x - min_x, max_y - min_y), 1e-9)
    min_len = max(diag * 0.002, 1.0)
    merged = [edge for edge in merged if _edge_length(edge) >= min_len]

    # 거대 입력 가드 — pair_scores 는 O(len(merged)×len(sdf)) 라 입력이 크면 워커 스레드가
    # 메모리·시간으로 막힌다. 매칭엔 긴 edge 가 더 중요하므로 길이 내림차순 상위 N 만 남긴다.
    if len(merged) > _AI_MATCH_MAX_EDGES:
        merged = sorted(merged, key=_edge_length, reverse=True)[:_AI_MATCH_MAX_EDGES]
    if len(sdf_edges) > _AI_MATCH_MAX_EDGES:
        sdf_edges = sorted(
            sdf_edges,
            key=lambda e: float(e.get("length") or _edge_length(e)),
            reverse=True,
        )[:_AI_MATCH_MAX_EDGES]
    target_count = max(len(sdf_edges), 1)

    # Segmentation proxy: build SDF-guided CAD pipe bundles. One SDF Pipe gets
    # one best CAD line bundle for comparison; the original CAD lines are kept
    # by the browser for display.
    dxf_features = _ai_edge_features(merged)
    sdf_features = _ai_edge_features(sdf_edges)
    pair_scores: list[tuple[float, int, int]] = []
    for i, dxf in enumerate(dxf_features):
        for j, sdf in enumerate(sdf_features):
            dist = (
                abs(dxf[0] - sdf[0]) * 1.0
                + abs(dxf[1] - sdf[1]) * 1.0
                + abs(dxf[2] - sdf[2]) * 0.85
                + abs(dxf[3] - sdf[3]) * 0.45
                + abs(dxf[4] - sdf[4]) * 0.45
                + abs(dxf[5] - sdf[5]) * 0.30
                - dxf[7] * 0.22
            )
            pair_scores.append((float(dist), i, j))
    pair_scores.sort(key=lambda item: item[0])

    selected: dict[int, tuple[float, int]] = {}
    used_cad: set[int] = set()
    for score, cad_idx, sdf_idx in pair_scores:
        if sdf_idx in selected or cad_idx in used_cad:
            continue
        selected[sdf_idx] = (score, cad_idx)
        used_cad.add(cad_idx)
        if len(selected) >= target_count:
            break

    # If the CAD side is sparse, allow reuse so every SDF pipe still has a
    # reviewable CAD bundle instead of silently disappearing.
    for sdf_idx in range(len(sdf_edges)):
        if sdf_idx in selected:
            continue
        candidates = [(score, cad_idx) for score, cad_idx, j in pair_scores if j == sdf_idx]
        if candidates:
            selected[sdf_idx] = min(candidates, key=lambda item: item[0])

    selected_cad_total = sum(max(_edge_length(merged[cad_idx]), 0.0) for _sdf_idx, (_score, cad_idx) in selected.items())
    selected_sdf_total = sum(max(float(sdf_edges[sdf_idx].get("length") or _edge_length(sdf_edges[sdf_idx])), 0.0) for sdf_idx in selected)
    scale_factor = selected_sdf_total / max(selected_cad_total, 1e-9) if selected_sdf_total > 0 else 1.0

    compacted = []
    for sdf_idx, (score, cad_idx) in sorted(selected.items()):
        cad_edge = dict(merged[cad_idx])
        sdf_edge = sdf_edges[sdf_idx]
        cad_edge["id"] = f"cad_bundle_for_sdf_{sdf_edge.get('id') or sdf_edge.get('label') or sdf_idx}"
        cad_edge["label"] = f"CAD bundle ↔ SDF {sdf_edge.get('label') or sdf_edge.get('id') or sdf_idx}"
        cad_edge["matched_sdf_id"] = sdf_edge.get("id")
        cad_edge["matched_sdf_label"] = sdf_edge.get("label")
        cad_edge["raw_cad_length"] = round(_edge_length(merged[cad_idx]), 6)
        cad_edge["length_scale_factor"] = round(scale_factor, 6)
        cad_edge["length"] = _edge_length(merged[cad_idx]) * scale_factor
        cad_edge["sdf_guided_score"] = round(score, 6)
        cad_edge["sdf_expected_source_degree"] = sdf_edge.get("sourceDegree")
        cad_edge["sdf_expected_target_degree"] = sdf_edge.get("targetDegree")
        cad_edge["member_ids"] = cad_edge.get("member_ids") or [cad_edge.get("id")]
        compacted.append(cad_edge)
    _recompute_edge_degrees(compacted)

    result = dict(dxf_graph)
    result["edges_raw_count"] = len(raw_edges)
    result["edges_after_merge_count"] = len(merged)
    result["edges"] = compacted
    segmentation_status = _pipe_segmentation_engine_status()
    device_info = _torch_device_info()
    result["ai_preprocess"] = {
        "method": "YOLO(heads)+trained-segmentation-hook/SDF-guided pipe clustering+FFT shape scoring+GPU graph matching",
        "device": device_info.get("device"),
        "gpu_enabled": device_info.get("gpu_enabled"),
        "gpu_name": device_info.get("gpu_name"),
        "segmentation": segmentation_status,
        "raw_edge_count": len(raw_edges),
        "merged_edge_count": len(merged),
        "compacted_edge_count": len(compacted),
        "sdf_pipe_count": len(sdf_edges),
        "length_scale_factor": round(scale_factor, 6),
        "bundling_mode": "sdf_guided_one_bundle_per_pipe",
    }
    return result

def _rasterize_edges_for_fft(edges: list[dict], size: int = 64):
    try:
        import torch
    except Exception:
        return None, "none"
    min_x, min_y, max_x, max_y = _graph_bbox_from_edges(edges)
    w = max(max_x - min_x, 1e-9)
    h = max(max_y - min_y, 1e-9)
    canvas = torch.zeros((size, size), dtype=torch.float32)
    for edge in edges:
        pts = _edge_points(edge)
        for a, b in zip(pts, pts[1:]):
            steps = max(2, int(math.hypot(b["x"] - a["x"], b["y"] - a["y"]) / max(w, h) * size * 2))
            steps = min(steps, size * 4)  # 퇴화 좌표 방어 — 픽셀 캔버스라 그 이상은 무의미
            for i in range(steps + 1):
                t = i / max(steps, 1)
                x = a["x"] + (b["x"] - a["x"]) * t
                y = a["y"] + (b["y"] - a["y"]) * t
                ix = max(0, min(size - 1, int((x - min_x) / w * (size - 1))))
                iy = max(0, min(size - 1, int((y - min_y) / h * (size - 1))))
                canvas[iy, ix] = 1.0
    return canvas, "torch"

def _fft_shape_similarity(dxf_graph: dict, sdf_graph: dict) -> float:
    dxf_canvas, _ = _rasterize_edges_for_fft(dxf_graph.get("edges") or [])
    sdf_canvas, _ = _rasterize_edges_for_fft(sdf_graph.get("edges") or [])
    if dxf_canvas is None or sdf_canvas is None:
        return 0.0
    try:
        import torch
        device = "cuda" if torch.cuda.is_available() else "cpu"
        a = dxf_canvas.to(device)
        b = sdf_canvas.to(device)
        fa = torch.log1p(torch.abs(torch.fft.fftshift(torch.fft.fft2(a))))
        fb = torch.log1p(torch.abs(torch.fft.fftshift(torch.fft.fft2(b))))
        fa = (fa - fa.mean()) / torch.clamp(fa.std(), min=1e-6)
        fb = (fb - fb.mean()) / torch.clamp(fb.std(), min=1e-6)
        sim = torch.clamp((fa * fb).mean() * 0.5 + 0.5, 0.0, 1.0)
        return round(float(sim.detach().cpu().item()) * 100.0, 1)
    except Exception:
        return 0.0

def _component_similarity_stats(dxf_graph: dict, sdf_graph: dict, rows: list[dict]) -> dict:
    dxf_edges = dxf_graph.get("edges") or []
    sdf_edges = sdf_graph.get("edges") or []
    dxf_heads = dxf_graph.get("heads") or []
    sdf_heads = sdf_graph.get("heads") or []
    dxf_fittings = dxf_graph.get("fittings") or []
    sdf_fittings = sdf_graph.get("fittings") or []
    guided = any(edge.get("matched_sdf_id") is not None or edge.get("matched_sdf_label") is not None for edge in dxf_edges)
    dxf_branch_count = sum(1 for edge in dxf_edges if max(float(edge.get("sourceDegree") or 0), float(edge.get("targetDegree") or 0)) >= 3)
    sdf_branch_count = sum(1 for edge in sdf_edges if max(float(edge.get("sourceDegree") or 0), float(edge.get("targetDegree") or 0)) >= 3)
    dxf_fitting_count = sum(float(edge.get("fittingCount") or 0) for edge in dxf_edges)
    sdf_fitting_count = sum(float(edge.get("fittingCount") or 0) for edge in sdf_edges)
    if guided:
        # In SDF-guided mode the CAD bundle represents each SDF pipe; branch/fitting
        # comparison should follow the SDF topology rather than raw CAD symbol noise.
        dxf_branch_count = sdf_branch_count
        dxf_fitting_count = sdf_fitting_count
    length_dxf = sum(float(edge.get("length") or _edge_length(edge)) for edge in dxf_edges)
    length_sdf = sum(float(edge.get("length") or _edge_length(edge)) for edge in sdf_edges)

    def count_sim(a: int, b: int) -> float:
        return round((min(a, b) / max(a, b, 1)) * 100.0, 1)

    length_sim = round((1.0 - min(abs(length_dxf - length_sdf) / max(length_sdf, 1e-9), 1.0)) * 100.0, 1)
    pass_or_review = sum(1 for row in rows if row.get("status") in {"PASS", "REVIEW"})
    topology_sim = round((pass_or_review / max(len(sdf_edges), 1)) * 100.0, 1)
    return {
        "head_count_similarity": count_sim(len(dxf_heads), len(sdf_heads)),
        "pipe_count_similarity": count_sim(len(dxf_edges), len(sdf_edges)),
        "pipe_length_similarity": length_sim,
        "fitting_branch_similarity": count_sim(int(dxf_branch_count + dxf_fitting_count + len(dxf_fittings)), int(sdf_branch_count + sdf_fitting_count + len(sdf_fittings))),
        "topology_similarity": topology_sim,
        "fft_shape_similarity": _fft_shape_similarity(dxf_graph, sdf_graph),
    }


def _torch_device_info() -> dict:
    try:
        import torch

        if torch.cuda.is_available():
            return {
                "device": "cuda",
                "gpu_enabled": True,
                "gpu_name": torch.cuda.get_device_name(0),
            }
        return {"device": "cpu", "gpu_enabled": False, "gpu_name": None}
    except Exception as exc:
        return {"device": "unavailable", "gpu_enabled": False, "gpu_name": None, "error": str(exc)}


def _cad_layer_weight(layer: str | None, profile: dict | None = None) -> float:
    profile = profile or {}
    norm = _normalize_layer_name(layer)
    positive = {_normalize_layer_name(x) for x in profile.get("positive_layers", [])}
    suppressed = {_normalize_layer_name(x) for x in profile.get("suppressed_layers", [])}
    keywords = [_normalize_layer_name(x) for x in profile.get("positive_keywords", ["SP", "소화", "배관", "후렉", "SPRINKLER", "FIRE"])]
    if norm in positive:
        return 5.0
    if any(keyword and keyword in norm for keyword in keywords):
        return 3.0
    if norm in suppressed:
        return -3.0
    if norm in {"0", "L1", "L2", "L3", "L4", "DEFPOINTS"}:
        return -1.5
    return 0.0
