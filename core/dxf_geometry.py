# -*- coding: utf-8 -*-
"""공유 코어 — DXF/그래프 기하 헬퍼 (순수 함수, 전역 미참조).

대조 서버.py 및 routes/* 에서 flat import. 함수 본문·시그니처는 원본 그대로.
"""
from __future__ import annotations

import math


def _to_float(v, default: float = 0.0) -> float:
    try:
        if v is None or v == "":
            return default
        return float(v)
    except (TypeError, ValueError):
        return default

def _point_on_polyline(path: list[tuple[float, float]], ratio: float) -> tuple[float, float] | None:
    if len(path) < 2:
        return path[0] if path else None
    ratio = max(0.0, min(1.0, ratio))
    seg_lengths: list[float] = []
    total = 0.0
    for i in range(len(path) - 1):
        x1, y1 = path[i]
        x2, y2 = path[i + 1]
        d = math.hypot(x2 - x1, y2 - y1)
        seg_lengths.append(d)
        total += d
    if total <= 0:
        return path[0]
    target = total * ratio
    acc = 0.0
    for i, d in enumerate(seg_lengths):
        if acc + d >= target:
            t = (target - acc) / d if d > 0 else 0.0
            x1, y1 = path[i]
            x2, y2 = path[i + 1]
            return (x1 + (x2 - x1) * t, y1 + (y2 - y1) * t)
        acc += d
    return path[-1]

def _cad_entity_points(ent: dict) -> list[list[float]]:
    if ent.get("type") == "LINE":
        return [[_to_float(ent.get("x")), _to_float(ent.get("y"))], [_to_float(ent.get("x2")), _to_float(ent.get("y2"))]]
    return [[_to_float(p[0]), _to_float(p[1])] for p in (ent.get("points") or [])]

def _polyline_length(points: list[list[float]]) -> float:
    total = 0.0
    for i in range(len(points) - 1):
        total += math.hypot(points[i + 1][0] - points[i][0], points[i + 1][1] - points[i][1])
    return total

def _normalize_layer_name(layer: str | None) -> str:
    return (layer or "").strip().upper().replace(" ", "")

def _entity_preview_row(ent: dict, idx: int) -> dict | None:
    points = _cad_entity_points(ent)
    if len(points) < 2:
        return None
    return {
        "id": f"E{idx}",
        "type": ent.get("type", "LINE"),
        "layer": ent.get("layer", ""),
        "points": points,
        "draw_length": _polyline_length(points),
    }

def _approx_arc_points(ent: dict, steps: int = 16) -> list[list[float]]:
    cx = _to_float(ent.get("x"))
    cy = _to_float(ent.get("y"))
    radius = abs(_to_float(ent.get("radius")))
    start = math.radians(_to_float(ent.get("start_angle")))
    end = math.radians(_to_float(ent.get("end_angle")))
    if radius <= 0:
        return []
    if end < start:
        end += math.tau
    return [[cx + math.cos(start + (end - start) * i / steps) * radius, cy + math.sin(start + (end - start) * i / steps) * radius] for i in range(steps + 1)]

def _bbox(points: list[dict]) -> dict | None:
    valid = [p for p in points if p.get("x") is not None and p.get("y") is not None]
    if not valid:
        return None
    xs = [float(p["x"]) for p in valid]
    ys = [float(p["y"]) for p in valid]
    return {"min_x": min(xs), "max_x": max(xs), "min_y": min(ys), "max_y": max(ys)}

def _norm_point(p: dict, box: dict) -> tuple[float, float]:
    w = max(float(box["max_x"]) - float(box["min_x"]), 1e-9)
    h = max(float(box["max_y"]) - float(box["min_y"]), 1e-9)
    return ((float(p["x"]) - float(box["min_x"])) / w, (float(p["y"]) - float(box["min_y"])) / h)

def _norm_xy(x: float, y: float, box: dict) -> tuple[float, float]:
    w = max(float(box["max_x"]) - float(box["min_x"]), 1e-9)
    h = max(float(box["max_y"]) - float(box["min_y"]), 1e-9)
    return ((float(x) - float(box["min_x"])) / w, (float(y) - float(box["min_y"])) / h)

def _segments_from_points(points: list[list[float]], box: dict, source_id: str, label: str | int | None = None) -> list[dict]:
    rows = []
    for i in range(len(points) - 1):
        x1, y1 = points[i]
        x2, y2 = points[i + 1]
        nx1, ny1 = _norm_xy(x1, y1, box)
        nx2, ny2 = _norm_xy(x2, y2, box)
        length = math.hypot(nx2 - nx1, ny2 - ny1)
        if length <= 1e-7:
            continue
        rows.append(
            {
                "source_id": source_id,
                "label": label,
                "mid": ((nx1 + nx2) / 2, (ny1 + ny2) / 2),
                "length": length,
                "angle": math.atan2(ny2 - ny1, nx2 - nx1),
            }
        )
    return rows

def _edge_points(edge: dict) -> list[dict]:
    pts = edge.get("points") or []
    clean = []
    for pt in pts:
        try:
            clean.append({"x": float(pt.get("x", 0.0)), "y": float(pt.get("y", 0.0))})
        except Exception:
            continue
    if len(clean) >= 2:
        return clean
    start = edge.get("start") or {}
    end = edge.get("end") or {}
    try:
        return [
            {"x": float(start.get("x", 0.0)), "y": float(start.get("y", 0.0))},
            {"x": float(end.get("x", 0.0)), "y": float(end.get("y", 0.0))},
        ]
    except Exception:
        return []

def _edge_length(edge: dict) -> float:
    pts = _edge_points(edge)
    if len(pts) < 2:
        return 0.0
    return sum(math.hypot(pts[i + 1]["x"] - pts[i]["x"], pts[i + 1]["y"] - pts[i]["y"]) for i in range(len(pts) - 1))

def _edge_angle(edge: dict) -> float:
    pts = _edge_points(edge)
    if len(pts) < 2:
        return 0.0
    return math.atan2(pts[-1]["y"] - pts[0]["y"], pts[-1]["x"] - pts[0]["x"])

def _angle_delta(a: float, b: float) -> float:
    diff = abs(a - b) % math.pi
    return math.pi - diff if diff > math.pi / 2 else diff

def _graph_bbox_from_edges(edges: list[dict]) -> tuple[float, float, float, float]:
    pts = [pt for edge in edges for pt in _edge_points(edge)]
    if not pts:
        return 0.0, 0.0, 100.0, 100.0
    min_x = min(pt["x"] for pt in pts)
    min_y = min(pt["y"] for pt in pts)
    max_x = max(pt["x"] for pt in pts)
    max_y = max(pt["y"] for pt in pts)
    if abs(max_x - min_x) < 1e-9:
        max_x += 100.0
    if abs(max_y - min_y) < 1e-9:
        max_y += 100.0
    return min_x, min_y, max_x, max_y

def _node_key(pt: dict, tolerance: float) -> str:
    return f"{round(float(pt.get('x', 0.0)) / tolerance)},{round(float(pt.get('y', 0.0)) / tolerance)}"
