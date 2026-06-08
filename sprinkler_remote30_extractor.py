# sprinkler_remote30_extractor.py
# 목적:
# 1) DXF에서 건축 배경 레이어를 제외하고 스프링클러 배관/헤드/관경문자만 추출
# 2) 알람밸브 기준 연결 배관망을 그래프로 구성
# 3) 알람밸브로부터 배관길이 기준 최원단 구역을 찾고, 상호 인접한 헤드 30개 선정
# 4) 실제 CAD 평면 형상을 유지한 Remote 30 Isometric PNG와 Excel Schedule 생성
#
# 사용 패키지: ezdxf, shapely, networkx, pandas, openpyxl, matplotlib, scipy
#
# 서버 통합용 함수형 API:
#   run_remote30_extraction(dxf_path, alarm_xy=None, out_dir=..., overrides=None) -> dict

from __future__ import annotations

import math
import re
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import ezdxf
import networkx as nx
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


DEFAULT_PIPE_LAYER_KEYWORDS = ["SP", "SPR", "SPRINKLER", "배관", "소방", "가지관", "후렉시블", "후렉", "FLEX"]
DEFAULT_HEAD_LAYER_KEYWORDS = ["HEAD", "헤드", "SP-H", "SPR_HEAD", "하향식", "하향", "상향식", "상향", "헤드반경"]
DEFAULT_TEXT_LAYER_KEYWORDS = ["TEXT", "문자", "관경", "SP", "TEX"]
DEFAULT_ARCH_LAYER_KEYWORDS = ["A-", "ARCH", "WALL", "DOOR", "WINDOW", "건축", "가구", "DIM", "SHEET", "AREA"]
DEFAULT_ALARM_VALVE_KEYWORDS = [
    "ALARM", "ALV", "AV", "알람", "알람밸브", "ALARM_VALVE", "ALARMVALVE",
    # 라이저 (입상관) 레이어 — 알람밸브 INSERT 가 라이저 레이어에 같이 있는 도면 흔함
    # (예: 다이소 세종허브센터 MF-125 의 "RISER" 레이어).
    "RISER", "라이저", "STAND-PIPE", "STANDPIPE", "STAND_PIPE",
]
# 소화기 / 옥내소화전 / 자동식 기기 등 — sprinkler 와 무관하므로 PIPE/HEAD 매칭에서 제외해야 함
DEFAULT_EXCLUDE_LAYER_KEYWORDS = ["소화기", "옥내소화전", "자동식", "CO2"]

DEFAULT_SNAP_TOL = 300.0
DEFAULT_HEAD_TO_PIPE_TOL = 800.0
DEFAULT_DIAMETER_TEXT_SEARCH_RADIUS = 1500.0
DEFAULT_REMOTE_HEAD_COUNT = 30
DEFAULT_CAD_UNIT_TO_M = 0.001  # 도면이 mm 기준이라는 가정 (인치면 0.0254 등)
DEFAULT_C_FACTOR = 120.0       # 일반 강관 기본 C값
DEFAULT_NOZZLE_LIBRARY_ITEM = "SP-HEAD"

# PIPENET 표준 명목지름 (m) — pipenet_converter.models.SUPPORTED_DIAMETERS_M 와 동일
PIPENET_SUPPORTED_DIAMETERS_M = {
    "25A": 0.025, "32A": 0.032, "40A": 0.040, "50A": 0.050,
    "65A": 0.065, "80A": 0.080, "100A": 0.100, "125A": 0.125,
    "150A": 0.150, "200A": 0.200,
}


def diameter_mm_to_label(dia_mm) -> str:
    if dia_mm is None:
        return ""
    target_m = round(float(dia_mm) / 1000.0, 6)
    for label, value in PIPENET_SUPPORTED_DIAMETERS_M.items():
        if round(value, 6) == target_m:
            return label
    return ""


@dataclass
class Remote30Settings:
    pipe_layer_keywords: list = field(default_factory=lambda: list(DEFAULT_PIPE_LAYER_KEYWORDS))
    head_layer_keywords: list = field(default_factory=lambda: list(DEFAULT_HEAD_LAYER_KEYWORDS))
    text_layer_keywords: list = field(default_factory=lambda: list(DEFAULT_TEXT_LAYER_KEYWORDS))
    arch_layer_keywords: list = field(default_factory=lambda: list(DEFAULT_ARCH_LAYER_KEYWORDS))
    alarm_valve_keywords: list = field(default_factory=lambda: list(DEFAULT_ALARM_VALVE_KEYWORDS))
    exclude_layer_keywords: list = field(default_factory=lambda: list(DEFAULT_EXCLUDE_LAYER_KEYWORDS))
    snap_tol: float = DEFAULT_SNAP_TOL
    head_to_pipe_tol: float = DEFAULT_HEAD_TO_PIPE_TOL
    diameter_text_search_radius: float = DEFAULT_DIAMETER_TEXT_SEARCH_RADIUS
    remote_head_count: int = DEFAULT_REMOTE_HEAD_COUNT
    cad_unit_to_m: float = DEFAULT_CAD_UNIT_TO_M
    c_factor: float = DEFAULT_C_FACTOR
    nozzle_library_item: str = DEFAULT_NOZZLE_LIBRARY_ITEM
    # === Hydraulic remote 계산용 ===
    remote_mode: str = "length"  # "length" | "hydraulic"
    elevation_alarm_m: float = 1.0     # 알람밸브 elevation (RV03_NEW 기준 1m)
    elevation_head_m: float = 2.7      # 헤드 elevation (천장, RV03_NEW 기준 2.7m 다수)
    # 자연낙차 / 옥상수원 가정 — 알람밸브-단위세대 사이 가상 수직 트렁크
    # 답안지 RV03_NEW: 92.8m + 40.6m + ... = 약 137m 수직 강하
    # DXF 에는 없는 도메인 정보. 0 이면 비활성 (평면 추출만)
    natural_fall_height_m: float = 0.0  # 옥상수원 → 단위세대 사이 자연낙차 높이 (m)
    natural_fall_trunk_bore_mm: float = 150.0  # 트렁크 직경
    natural_fall_n_stages: int = 2     # 트렁크를 몇 단(stage)으로 나눌지 (rise 분배)
    k_factor: float = 80.0             # K-factor (L/min / sqrt(bar))
    design_flow_per_head_lpm: float = 80.0  # 1방수당 유량 (LPM, NFTC 80LPM @ K=80, P=1bar)
    fallback_dia_mm: float = 50.0      # 관경 TEXT 누락 시 사용할 fallback 직경 (50A 강관 기본)
    branch_leaf_dia_mm: float = 25.0   # degree=1 끝가지(leaf) edge 의 추정 직경 (가지관 25A)
    emit_sdf: bool = False             # PIPENET SDF XML 파일 동시 출력
    emit_csv: bool = False             # PIPENET 4 tables 를 CSV 4개 파일로 동시 출력
    # === 누락 보완 휴리스틱 ===
    explode_inserts: bool = True       # INSERT 블록 분해해서 내부 line/arc 도 추출 (tee/elbow fitting)
    include_lines_near_diameter_text: bool = True   # 관경 TEXT 반경 안 layer-무관 LINE 자동 포함
    near_text_search_radius: float = 500.0          # TEXT 근접 LINE 포함 반경 (보수적으로)
    near_text_max_line_length: float = 8000.0       # 너무 긴 LINE(>8m)은 노이즈 가능성 → 제외
    graph_closure_tol: float = 1500.0  # degree=1 끝점 → 가까운 line 끝/중간 자동 snap 허용 거리 (0이면 비활성)
    # === Zone 영역 제한 (워크벤치) ===
    zone_bbox: tuple = None            # (x_min, y_min, x_max, y_max) — 이 영역 안 entity 만 사용 (None=전체)


# =========================
# 기본 유틸
# =========================
def layer_match(layer_name, keywords):
    layer_upper = str(layer_name).upper()
    return any(k.upper() in layer_upper for k in keywords)


def _in_bbox(xy, bbox):
    """xy 가 bbox 안에 있는지. bbox=None 이면 항상 True."""
    if bbox is None:
        return True
    x, y = xy
    return bbox[0] <= x <= bbox[2] and bbox[1] <= y <= bbox[3]


def _seg_in_bbox(p1, p2, bbox):
    """segment 의 어느 한쪽 끝이라도 bbox 안에 있으면 True."""
    if bbox is None:
        return True
    return _in_bbox(p1, bbox) or _in_bbox(p2, bbox)


def _get_line_entities(msp, settings: Remote30Settings):
    segments = []
    bbox = settings.zone_bbox
    for e in msp:
        layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
        if layer_match(layer, settings.arch_layer_keywords):
            continue
        if layer_match(layer, settings.exclude_layer_keywords):
            continue
        if not layer_match(layer, settings.pipe_layer_keywords):
            continue

        etype = e.dxftype()
        if etype == "LINE":
            p1 = (float(e.dxf.start.x), float(e.dxf.start.y))
            p2 = (float(e.dxf.end.x), float(e.dxf.end.y))
            if not _seg_in_bbox(p1, p2, bbox):
                continue
            segments.append({"layer": layer, "p1": p1, "p2": p2})
        elif etype == "LWPOLYLINE":
            pts = [(float(p[0]), float(p[1])) for p in e.get_points()]
            for a, b in zip(pts[:-1], pts[1:]):
                if _seg_in_bbox(a, b, bbox):
                    segments.append({"layer": layer, "p1": a, "p2": b})
        elif etype == "POLYLINE":
            pts = [(float(v.dxf.location.x), float(v.dxf.location.y)) for v in e.vertices]
            for a, b in zip(pts[:-1], pts[1:]):
                if _seg_in_bbox(a, b, bbox):
                    segments.append({"layer": layer, "p1": a, "p2": b})
        elif etype == "ARC":
            # 90도 elbow 등 곡선 배관/fitting — chord (start→end) 로 근사
            cx, cy = float(e.dxf.center.x), float(e.dxf.center.y)
            r = float(e.dxf.radius)
            sa = math.radians(float(e.dxf.start_angle))
            ea = math.radians(float(e.dxf.end_angle))
            p1 = (cx + r * math.cos(sa), cy + r * math.sin(sa))
            p2 = (cx + r * math.cos(ea), cy + r * math.sin(ea))
            if not _seg_in_bbox(p1, p2, bbox):
                continue
            segments.append({"layer": layer, "p1": p1, "p2": p2})
    return segments


def _explode_insert_to_segments(insert_entity, settings: Remote30Settings):
    """INSERT 블록을 virtual_entities() 로 분해해 내부 LINE/ARC/LWPOLYLINE 을 segment 로 변환.
    parent INSERT 의 layer 가 pipe-layer 매칭일 때만 호출하는 것을 권장.
    """
    out = []
    parent_layer = insert_entity.dxf.layer if hasattr(insert_entity.dxf, "layer") else ""
    try:
        for v in insert_entity.virtual_entities():
            etype = v.dxftype()
            vlayer = v.dxf.layer if hasattr(v.dxf, "layer") else ""
            effective_layer = vlayer if (vlayer and vlayer not in ("0",)) else parent_layer
            if layer_match(effective_layer, settings.exclude_layer_keywords):
                continue
            if layer_match(effective_layer, settings.arch_layer_keywords):
                continue
            if etype == "LINE":
                p1 = (float(v.dxf.start.x), float(v.dxf.start.y))
                p2 = (float(v.dxf.end.x), float(v.dxf.end.y))
                out.append({"layer": parent_layer, "p1": p1, "p2": p2, "source": "INSERT-exploded"})
            elif etype == "ARC":
                cx, cy = float(v.dxf.center.x), float(v.dxf.center.y)
                r = float(v.dxf.radius)
                sa = math.radians(float(v.dxf.start_angle))
                ea = math.radians(float(v.dxf.end_angle))
                p1 = (cx + r * math.cos(sa), cy + r * math.sin(sa))
                p2 = (cx + r * math.cos(ea), cy + r * math.sin(ea))
                out.append({"layer": parent_layer, "p1": p1, "p2": p2, "source": "INSERT-exploded"})
            elif etype == "LWPOLYLINE":
                pts = [(float(p[0]), float(p[1])) for p in v.get_points()]
                for a, b in zip(pts[:-1], pts[1:]):
                    out.append({"layer": parent_layer, "p1": a, "p2": b, "source": "INSERT-exploded"})
    except Exception:
        pass
    return out


def _collect_supplementary_segments(msp, base_segments, text_items, settings: Remote30Settings):
    """누락 보완 휴리스틱:
    (1) pipe-layer INSERT 분해 (tee/elbow fitting 내부 line)
    (2) 관경 TEXT 반경 안의 LINE 자동 포함 (layer 무관, arch/exclude 만 제외)
    returns: 추가 segment 리스트
    """
    extra = []
    counts = {"insert_exploded": 0, "near_text_line": 0}

    if settings.explode_inserts:
        for e in msp:
            if e.dxftype() != "INSERT":
                continue
            layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
            if layer_match(layer, settings.exclude_layer_keywords):
                continue
            if layer_match(layer, settings.arch_layer_keywords):
                continue
            if not layer_match(layer, settings.pipe_layer_keywords):
                continue
            new_segs = _explode_insert_to_segments(e, settings)
            extra.extend(new_segs)
            counts["insert_exploded"] += len(new_segs)

    if settings.include_lines_near_diameter_text and text_items:
        text_xys = [t["xy"] for t in text_items]
        radius = settings.near_text_search_radius
        radius_sq = radius * radius
        max_len = settings.near_text_max_line_length
        # 이미 잡힌 segment 중복 회피용 set
        existing_set = {(round(s["p1"][0], 2), round(s["p1"][1], 2),
                         round(s["p2"][0], 2), round(s["p2"][1], 2)) for s in base_segments}
        # 기존 pipe segment 끝점 + 중간점 — 추가 LINE이 기존 네트워크와 연결성 있어야 진짜 pipe
        existing_endpoints = []
        for s in base_segments:
            existing_endpoints.append(s["p1"])
            existing_endpoints.append(s["p2"])
        connectivity_tol = radius  # 기존 끝점 반경 안에 들어와야 함

        for e in msp:
            if e.dxftype() != "LINE":
                continue
            layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
            if layer_match(layer, settings.arch_layer_keywords):
                continue
            if layer_match(layer, settings.exclude_layer_keywords):
                continue
            if layer_match(layer, settings.pipe_layer_keywords):
                continue  # 이미 base 에 포함
            p1 = (float(e.dxf.start.x), float(e.dxf.start.y))
            p2 = (float(e.dxf.end.x), float(e.dxf.end.y))
            # 너무 긴 LINE 은 건축물/도면 외곽선일 가능성 — 제외
            if math.dist(p1, p2) > max_len:
                continue
            mid = ((p1[0] + p2[0]) / 2, (p1[1] + p2[1]) / 2)
            # 조건 1: 관경 TEXT 가까이
            near_text = False
            for tx, ty in text_xys:
                dx = mid[0] - tx
                dy = mid[1] - ty
                if dx * dx + dy * dy <= radius_sq:
                    near_text = True
                    break
            if not near_text:
                continue
            # 조건 2: LINE 끝점이 기존 pipe 네트워크와 가까워야 (연결성)
            connected = False
            for ep in existing_endpoints:
                if math.dist(p1, ep) <= connectivity_tol or math.dist(p2, ep) <= connectivity_tol:
                    connected = True
                    break
            if not connected:
                continue
            key = (round(p1[0], 2), round(p1[1], 2), round(p2[0], 2), round(p2[1], 2))
            if key in existing_set:
                continue
            existing_set.add(key)
            extra.append({"layer": layer, "p1": p1, "p2": p2, "source": "near-text"})
            counts["near_text_line"] += 1
    return extra, counts


def _close_graph_gaps(G, node_coords, settings: Remote30Settings):
    """그래프 closure: degree==1 끝점이 다른 노드/segment 끝과 N 거리 안이면 가짜 edge 로 연결.
    원본 segment 가 아니므로 dia=None, source='closure', weight=실제 거리.
    """
    if settings.graph_closure_tol <= 0:
        return 0
    closure_tol = settings.graph_closure_tol
    closure_tol_sq = closure_tol * closure_tol

    leaf_nodes = [n for n in G.nodes() if G.degree(n) == 1]
    added = 0
    for u in leaf_nodes:
        ux, uy = node_coords[u]
        best = None
        best_d_sq = closure_tol_sq + 1
        for v, (vx, vy) in enumerate(node_coords):
            if v == u:
                continue
            if G.has_edge(u, v):
                continue
            dx = ux - vx
            dy = uy - vy
            d_sq = dx * dx + dy * dy
            if d_sq < best_d_sq:
                best_d_sq = d_sq
                best = v
        if best is not None and best_d_sq <= closure_tol_sq:
            d = math.sqrt(best_d_sq)
            raw_p1 = (float(node_coords[u][0]), float(node_coords[u][1]))
            raw_p2 = (float(node_coords[best][0]), float(node_coords[best][1]))
            G.add_edge(u, best, weight=d, pipe_no=f"PC{added:04d}", length=d, dia=None,
                       layer="<closure>", raw_p1=raw_p1, raw_p2=raw_p2, is_closure=True)
            added += 1
    return added


def _refine_edge_diameters(G, settings: Remote30Settings):
    """edge dia 가 None 인 경우 그래프 구조로 추정:
       - degree=1 (leaf) 양쪽 끝 edge → branch_leaf_dia_mm (가지관 25A)
       - 짧은 edge (head 부근) → 가지관 25A
       - 그 외 → fallback_dia_mm (50A)
    """
    fdia = settings.fallback_dia_mm
    bdia = settings.branch_leaf_dia_mm
    short_len_threshold = 1.5  # 1.5m 보다 짧은 edge 는 가지말단으로 가정
    scale = settings.cad_unit_to_m
    refined = 0
    for u, v, data in G.edges(data=True):
        if data.get("dia"):
            continue  # TEXT 매칭 결과 보존
        length_m = data.get("length", 0) * scale
        deg_u = G.degree(u); deg_v = G.degree(v)
        if deg_u == 1 or deg_v == 1:
            # leaf edge - 가지관 말단
            data["dia"] = bdia
            refined += 1
        elif length_m < short_len_threshold:
            data["dia"] = bdia
            refined += 1
        else:
            data["dia"] = fdia
    return refined


def _bridge_components(G, node_coords, alarm_node):
    """알람밸브가 속한 main component 와 다른 모든 component 를 가장 가까운 노드쌍으로
    가짜 edge 로 연결한다. 결과: 모든 헤드가 알람밸브와 한 connected graph 에 속함.
    """
    if G.number_of_edges() == 0 or G.number_of_nodes() == 0:
        return 0
    components = list(nx.connected_components(G))
    if len(components) <= 1:
        return 0
    main = next((c for c in components if alarm_node in c), None)
    if main is None:
        # 알람밸브가 어떤 component 에도 없으면 가장 큰 component 를 main
        main = max(components, key=len)
    main_nodes = list(main)
    added = 0
    for comp in components:
        if comp is main:
            continue
        # comp 의 각 노드에서 main 의 가장 가까운 노드 (최소 거리 쌍)
        best_u, best_v, best_d = None, None, float("inf")
        for u in comp:
            ux, uy = node_coords[u]
            for v in main_nodes:
                vx, vy = node_coords[v]
                d = math.hypot(ux - vx, uy - vy)
                if d < best_d:
                    best_d = d; best_u = u; best_v = v
        if best_u is not None:
            raw_p1 = (float(node_coords[best_u][0]), float(node_coords[best_u][1]))
            raw_p2 = (float(node_coords[best_v][0]), float(node_coords[best_v][1]))
            G.add_edge(
                best_u, best_v,
                weight=best_d, pipe_no=f"PB{added:04d}",
                length=best_d, dia=None, layer="<auto-bridge>",
                raw_p1=raw_p1, raw_p2=raw_p2, is_closure=True,
            )
            added += 1
            main_nodes.extend(comp)  # 다음 comp는 이제 main+comp 둘 다와 비교
    return added


def _force_attach_orphan_heads(orphan_heads, G, node_coords, settings: Remote30Settings):
    """head_to_pipe_tol 초과로 부착 실패한 헤드들을 강제 부착.
    각 헤드 좌표를 새 노드로 추가하고, 가장 가까운 기존 노드와 가짜 edge 로 연결.
    """
    added_edges = 0
    new_attached = []
    if not orphan_heads or not node_coords:
        return new_attached, added_edges
    for i, h in enumerate(orphan_heads, 1):
        xy = h["xy"]
        # 가장 가까운 기존 노드
        nearest_idx = _nearest_node_id(xy, node_coords)
        d = math.dist(xy, node_coords[nearest_idx])
        # 새 노드 추가 (헤드 위치 그대로)
        new_node_idx = len(node_coords)
        node_coords.append((float(xy[0]), float(xy[1])))
        G.add_node(new_node_idx, xy=(float(xy[0]), float(xy[1])))
        # 가짜 edge 로 연결
        raw_p1 = (float(xy[0]), float(xy[1]))
        raw_p2 = (float(node_coords[nearest_idx][0]), float(node_coords[nearest_idx][1]))
        G.add_edge(
            new_node_idx, nearest_idx,
            weight=d, pipe_no=f"PH{i:04d}",
            length=d, dia=None, layer="<head-bridge>",
            raw_p1=raw_p1, raw_p2=raw_p2, is_closure=True,
        )
        added_edges += 1
        new_attached.append({
            "Head No": f"OH{i:03d}",
            "Node": new_node_idx,
            "Node Name": f"N{new_node_idx:04d}",
            "X": float(xy[0]), "Y": float(xy[1]),
            "Attach Distance": round(d, 2),
            "Layer": h.get("layer"), "Block": h.get("block"),
        })
    return new_attached, added_edges


def _get_text_entities(msp, settings: Remote30Settings):
    texts = []
    bbox = settings.zone_bbox
    for e in msp:
        etype = e.dxftype()
        if etype not in ["TEXT", "MTEXT"]:
            continue
        layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
        if layer_match(layer, settings.arch_layer_keywords):
            continue

        raw = e.dxf.text if etype == "TEXT" else e.text
        raw = str(raw).strip()

        m = re.search(r"\b(20|25|32|40|50|65|80|100|125|150|200)\b", raw)
        if not m:
            continue

        xy = (float(e.dxf.insert.x), float(e.dxf.insert.y))
        if not _in_bbox(xy, bbox):
            continue
        texts.append({"xy": xy, "text": raw, "dia": int(m.group(1)), "layer": layer})
    return texts


def _get_head_candidates(msp, settings: Remote30Settings):
    heads = []
    bbox = settings.zone_bbox
    for e in msp:
        layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
        if layer_match(layer, settings.exclude_layer_keywords):
            continue
        etype = e.dxftype()
        if etype == "INSERT":
            name = e.dxf.name
            if layer_match(layer, settings.head_layer_keywords) or layer_match(name, settings.head_layer_keywords):
                xy = (float(e.dxf.insert.x), float(e.dxf.insert.y))
                if not _in_bbox(xy, bbox):
                    continue
                heads.append({"xy": xy, "layer": layer, "block": name})
        elif etype == "CIRCLE":
            if layer_match(layer, settings.head_layer_keywords):
                xy = (float(e.dxf.center.x), float(e.dxf.center.y))
                if not _in_bbox(xy, bbox):
                    continue
                heads.append({"xy": xy, "layer": layer, "block": "CIRCLE_HEAD"})
    return heads


def detect_alarm_valve_xy(msp, settings: Remote30Settings) -> Optional[tuple]:
    """DXF에서 알람밸브 위치를 자동 검출. INSERT 블록명/레이어 우선, 그다음 TEXT."""
    for e in msp:
        if e.dxftype() != "INSERT":
            continue
        layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
        name = e.dxf.name
        if layer_match(layer, settings.alarm_valve_keywords) or layer_match(name, settings.alarm_valve_keywords):
            return (float(e.dxf.insert.x), float(e.dxf.insert.y))

    for e in msp:
        etype = e.dxftype()
        if etype not in ["TEXT", "MTEXT"]:
            continue
        layer = e.dxf.layer if hasattr(e.dxf, "layer") else ""
        raw = e.dxf.text if etype == "TEXT" else e.text
        raw = str(raw).strip()
        if layer_match(layer, settings.alarm_valve_keywords) or layer_match(raw, settings.alarm_valve_keywords):
            return (float(e.dxf.insert.x), float(e.dxf.insert.y))
    return None


def _snap_points(points, tol):
    snapped = []
    node_coords = []
    for p in points:
        found = None
        for i, q in enumerate(node_coords):
            if math.dist(p, q) <= tol:
                found = i
                break
        if found is None:
            node_coords.append(p)
            snapped.append(len(node_coords) - 1)
        else:
            snapped.append(found)
    return snapped, node_coords


def _nearest_node_id(xy, node_coords):
    dists = [math.dist(xy, n) for n in node_coords]
    return int(min(range(len(dists)), key=lambda i: dists[i]))


def _assign_pipe_diameter(pipe_mid, text_items, search_radius):
    if not text_items:
        return None
    best = None
    best_dist = 10**18
    for t in text_items:
        dist = math.dist(pipe_mid, t["xy"])
        if dist < best_dist and dist <= search_radius:
            best = t
            best_dist = dist
    return best["dia"] if best else None


def _build_pipe_graph(segments, text_items, settings: Remote30Settings):
    all_pts = []
    for s in segments:
        all_pts.append(s["p1"])
        all_pts.append(s["p2"])

    snapped_ids, node_coords = _snap_points(all_pts, settings.snap_tol)

    G = nx.Graph()
    for i, xy in enumerate(node_coords):
        G.add_node(i, xy=xy)

    pipe_records = []
    sid = 0
    for idx, s in enumerate(segments):
        n1 = snapped_ids[2 * idx]
        n2 = snapped_ids[2 * idx + 1]
        if n1 == n2:
            continue

        # raw 좌표 = 원본 segment 의 실제 끝점 (시각화 일탈 방지)
        raw_p1 = (float(s["p1"][0]), float(s["p1"][1]))
        raw_p2 = (float(s["p2"][0]), float(s["p2"][1]))
        # 노드 끝점이 n1,n2 순서대로 매칭되도록 raw 좌표도 정렬
        # n1 은 첫 끝점(2*idx) 에 대응 = s["p1"]
        # n2 는 두 번째 끝점(2*idx+1) 에 대응 = s["p2"]
        length = math.dist(raw_p1, raw_p2)
        mid = ((raw_p1[0] + raw_p2[0]) / 2, (raw_p1[1] + raw_p2[1]) / 2)
        dia = _assign_pipe_diameter(mid, text_items, settings.diameter_text_search_radius)

        sid += 1
        pipe_no = f"P{sid:04d}"

        # 이미 같은 (n1,n2) edge 가 있는 경우 — 첫 번째 raw 좌표 유지 (가장 가까운 segment 의 좌표)
        if G.has_edge(n1, n2):
            existing = G.get_edge_data(n1, n2)
            # 더 짧은 length 의 raw 를 유지 (snap 후 가장 정확한 원본)
            if existing.get("length", float("inf")) <= length:
                continue
        G.add_edge(
            n1, n2,
            weight=length, pipe_no=pipe_no, length=length, dia=dia, layer=s["layer"],
            raw_p1=raw_p1, raw_p2=raw_p2,
        )

        pipe_records.append({
            "Pipe No": pipe_no,
            "From Node": f"N{n1:04d}",
            "To Node": f"N{n2:04d}",
            "Dia(mm)": dia,
            "Length(CAD unit)": round(length, 2),
            "Layer": s["layer"],
        })
    return G, node_coords, pipe_records


def _attach_heads_to_graph(heads, G, node_coords, settings: Remote30Settings):
    attached = []
    for i, h in enumerate(heads, 1):
        xy = h["xy"]
        n = _nearest_node_id(xy, node_coords)
        dist = math.dist(xy, node_coords[n])
        if dist <= settings.head_to_pipe_tol:
            attached.append({
                "Head No": f"H{i:03d}",
                "Node": n,
                "Node Name": f"N{n:04d}",
                "X": xy[0],
                "Y": xy[1],
                "Attach Distance": round(dist, 2),
                "Layer": h.get("layer"),
                "Block": h.get("block"),
            })
    return attached


def hw_friction_loss_kgcm2(q_lpm: float, length_m: float, c_factor: float, bore_mm: float) -> float:
    """Hazen-Williams 마찰손실 (pipenet_validator.py:757 와 동일 공식).
    returns: kgf/cm² (= 0.1 MPa = 10 mH2O ≈ 0.9806 bar)
    """
    if bore_mm <= 0 or length_m <= 0 or q_lpm <= 0 or c_factor <= 0:
        return 0.0
    dp_mpa = 6.174e4 * (q_lpm ** 1.85) * length_m / ((c_factor ** 1.85) * (bore_mm ** 4.87))
    return dp_mpa / 0.1


def _compute_head_paths_and_branch_load(G, alarm_node, attached_heads):
    """각 헤드별 alarm→head 경로를 계산하고, 각 edge에 매달린 헤드 수를 카운트.
    트리(분기) 가정으로 동시방수 시 pipe별 유량 ≈ (그 pipe 아래 매달린 헤드 수) × flow_per_head.
    """
    from collections import Counter as _Counter
    lengths = nx.single_source_dijkstra_path_length(G, alarm_node, weight="weight")
    head_paths = {}
    for h in attached_heads:
        n = h["Node"]
        if n in lengths and n != alarm_node:
            try:
                head_paths[n] = nx.shortest_path(G, alarm_node, n, weight="weight")
            except nx.NetworkXNoPath:
                continue
    edge_head_count = _Counter()
    for n, path in head_paths.items():
        for a, b in zip(path[:-1], path[1:]):
            edge_head_count[frozenset((a, b))] += 1
    return head_paths, edge_head_count, lengths


def _compute_hydraulic_loss_per_head(G, head_paths, edge_head_count, settings: Remote30Settings) -> dict:
    """각 헤드별 alarm→head 경로의 friction loss + static head (kgf/cm²) 계산.
    returns: {node_id: {'friction': float, 'static': float, 'total': float, 'path_length_m': float}}
    """
    flow_per_head = settings.design_flow_per_head_lpm
    c = settings.c_factor
    scale = settings.cad_unit_to_m
    static_head_m = settings.elevation_head_m - settings.elevation_alarm_m
    static_head_kgcm2 = static_head_m * 0.0980665  # 1 mH2O ≈ 0.0980665 kgf/cm²

    out = {}
    for n, path in head_paths.items():
        friction = 0.0
        path_len = 0.0
        for a, b in zip(path[:-1], path[1:]):
            ed = G.get_edge_data(a, b)
            length_m = ed.get("length", 0.0) * scale
            dia_mm = ed.get("dia") or settings.fallback_dia_mm
            n_below = edge_head_count[frozenset((a, b))]
            q_lpm = n_below * flow_per_head
            friction += hw_friction_loss_kgcm2(q_lpm, length_m, c, dia_mm)
            path_len += length_m
        out[n] = {
            "friction": friction,
            "static": static_head_kgcm2,
            "total": friction + static_head_kgcm2,
            "path_length_m": path_len,
        }
    return out


def _select_remote_hydraulic(G, node_coords, attached_heads, alarm_xy, settings: Remote30Settings):
    """Hydraulic 모드: 선정 자체는 length 모드와 같이 배관망 경로 거리 기준 상위 N개.
    선정된 N개만 동시방수한다고 가정하고 hydraulic loss(마찰 + 정수두)를 계산해서
    정보(Friction/Static/Total Loss, Path Length(m))를 각 헤드에 첨부한다.

    설계 결정 — 변위(평면거리) 군집은 제거한다.
    사용자 요구: "배관망을 하나의 길로 삼아서 그 거리들 중 가장 먼 30개".
    """
    # === Pass 1: 배관망 경로 거리(dijkstra) 기준 상위 N개 선정 ===
    selected, alarm_node, path_edges = _select_remote_contiguous(
        G, node_coords, attached_heads, alarm_xy, settings
    )
    if not selected:
        return [], alarm_node, set()

    # === Pass 2: 선정된 N개만 동시방수한다고 가정하고 hydraulic loss 계산 ===
    head_paths, edge_head_count, _ = _compute_head_paths_and_branch_load(G, alarm_node, selected)
    hydra = _compute_hydraulic_loss_per_head(G, head_paths, edge_head_count, settings)

    for h in selected:
        n = h["Node"]
        if n in hydra:
            h["Friction Loss(kgcm2)"] = hydra[n]["friction"]
            h["Static Head(kgcm2)"] = hydra[n]["static"]
            h["Total Loss(kgcm2)"] = hydra[n]["total"]
            h["Path Length(m)"] = hydra[n]["path_length_m"]

    # 표시 순서: 배관 경로 거리 먼 순 (RH01 = 가장 먼 헤드)
    selected.sort(key=lambda x: -x["Distance from AV(CAD unit)"])
    for idx, h in enumerate(selected, 1):
        h["Remote Head No"] = f"RH{idx:02d}"

    # path_edges 재구성 (번호 변경 후)
    path_edges = set()
    for h in selected:
        try:
            path = nx.shortest_path(G, alarm_node, h["Node"], weight="weight")
            for a, b in zip(path[:-1], path[1:]):
                path_edges.add(tuple(sorted((a, b))))
        except nx.NetworkXNoPath:
            continue

    return selected, alarm_node, path_edges


def _select_remote_contiguous(G, node_coords, attached_heads, alarm_xy, settings: Remote30Settings):
    """배관망 경로 거리(dijkstra) 기준 상위 N개 헤드를 선정.
    평면거리(변위) 군집은 사용하지 않는다 — 사용자 요구: "배관망을 하나의 길로 삼아서 가장 먼 30개".
    """
    alarm_node = _nearest_node_id(alarm_xy, node_coords)
    lengths = nx.single_source_dijkstra_path_length(G, alarm_node, weight="weight")

    enriched = []
    for h in attached_heads:
        n = h["Node"]
        if n not in lengths or n == alarm_node:
            continue
        item = dict(h)
        item["Distance from AV(CAD unit)"] = lengths[n]
        # 'Plan distance to farthest' 는 더 이상 선정 기준이 아니지만 기존 export 와 호환을 위해 0 으로 채움
        item["Plan distance to farthest"] = 0.0
        enriched.append(item)

    if not enriched:
        return [], alarm_node, set()

    # 배관망 경로 거리(=dijkstra) 큰 순으로 상위 N 개
    selected = sorted(enriched, key=lambda x: -x["Distance from AV(CAD unit)"])[: settings.remote_head_count]
    for idx, h in enumerate(selected, 1):
        h["Remote Head No"] = f"RH{idx:02d}"

    path_edges = set()
    for h in selected:
        try:
            path = nx.shortest_path(G, alarm_node, h["Node"], weight="weight")
            for a, b in zip(path[:-1], path[1:]):
                path_edges.add(tuple(sorted((a, b))))
        except nx.NetworkXNoPath:
            continue

    return selected, alarm_node, path_edges


def _plot_extracted_isometric(G, node_coords, selected_heads, alarm_node, path_edges, out_png):
    plt.figure(figsize=(16, 11))
    ax = plt.gca()

    for a, b in path_edges:
        x1, y1 = node_coords[a]
        x2, y2 = node_coords[b]
        ed = G.get_edge_data(a, b)
        dia = ed.get("dia")
        pipe_no = ed.get("pipe_no")

        ax.plot([x1, x2], [y1, y2], linewidth=2.2, color="black")
        mx, my = (x1 + x2) / 2, (y1 + y2) / 2
        label = f"{pipe_no}"
        if dia:
            label += f" / {dia}A"
        ax.text(mx, my, label, fontsize=7, color="blue")
        ax.text(x1, y1, f"N{a:04d}", fontsize=6, color="green")
        ax.text(x2, y2, f"N{b:04d}", fontsize=6, color="green")

    avx, avy = node_coords[alarm_node]
    ax.scatter([avx], [avy], s=180, marker="s", color="red")
    ax.text(avx, avy, "ALARM VALVE", fontsize=10, color="red", fontweight="bold")

    for h in selected_heads:
        x, y = h["X"], h["Y"]
        ax.scatter([x], [y], s=80, facecolors="yellow", edgecolors="black")
        ax.text(x, y, h["Remote Head No"], fontsize=8, color="red", fontweight="bold")

    ax.set_aspect("equal", adjustable="box")
    ax.set_title("Remote 30 Sprinkler Heads - Actual CAD Shape Extraction")
    ax.axis("off")
    plt.tight_layout()
    plt.savefig(out_png, dpi=220)
    plt.close()


def build_pipenet_tables(G, node_coords, selected_heads, path_edges, alarm_node, settings: Remote30Settings):
    """PIPENET 호환 4-시트 dict (Excel + SDF + 워크벤치 미리보기 공통)."""
    scale = settings.cad_unit_to_m

    used_nodes = set()
    for a, b in path_edges:
        used_nodes.add(a)
        used_nodes.add(b)
    used_nodes.add(alarm_node)

    head_nodes_to_id = {h["Node"]: h["Remote Head No"] for h in selected_heads}

    pipenet_node_rows = []
    for n in sorted(used_nodes):
        x_cad, y_cad = node_coords[n]
        if n == alarm_node:
            ntype = "source"
            src = "alarm_valve"
            elev = settings.elevation_alarm_m
        elif n in head_nodes_to_id:
            ntype = "nozzle"
            src = head_nodes_to_id[n]
            elev = settings.elevation_head_m
        else:
            ntype = "junction"
            src = "DXF"
            elev = settings.elevation_head_m
        pipenet_node_rows.append({
            "node_id": f"N{n:04d}",
            "x": round(x_cad * scale, 6),
            "y": round(y_cad * scale, 6),
            "z": elev,
            "node_type": ntype,
            "source": src,
        })

    pipenet_pipe_rows = []
    for a, b in sorted(path_edges):
        ed = G.get_edge_data(a, b)
        dia_mm = ed.get("dia")
        dia_m = (float(dia_mm) / 1000.0) if dia_mm else (settings.fallback_dia_mm / 1000.0)
        pipenet_pipe_rows.append({
            "pipe_id": ed.get("pipe_no") or f"P{a:04d}_{b:04d}",
            "from_node": f"N{a:04d}",
            "to_node": f"N{b:04d}",
            "diameter_m": dia_m,
            "diameter_label": diameter_mm_to_label(dia_mm),
            "length_m": round(ed.get("length", 0.0) * scale, 6),
            "rise_m": 0.0,
            "c_factor": settings.c_factor,
            "material": "",
            "status": "normal",
        })

    pipenet_nozzle_rows = []
    for h in selected_heads:
        rh_no = h["Remote Head No"]
        pipenet_nozzle_rows.append({
            "nozzle_id": rh_no,
            "input_node": h["Node Name"],
            "output_node": f"OUT_{rh_no}",
            "flow_m3s": round(settings.design_flow_per_head_lpm / 60000.0, 8),
            "status": 1,
            "library_item": settings.nozzle_library_item,
        })

    pipenet_valve_rows = [{
        "valve_id": "AV01",
        "input_node": f"N{alarm_node:04d}",
        "output_node": f"N{alarm_node:04d}",
        "valve_type": "alarm",
        "target_value": None,
    }]

    return {
        "nodes": pipenet_node_rows,
        "pipes": pipenet_pipe_rows,
        "nozzles": pipenet_nozzle_rows,
        "valves": pipenet_valve_rows,
    }


def _export_excel(G, node_coords, pipe_records, selected_heads, path_edges, alarm_node, settings: Remote30Settings, out_xlsx, tables=None):
    if tables is None:
        tables = build_pipenet_tables(G, node_coords, selected_heads, path_edges, alarm_node, settings)
    scale = settings.cad_unit_to_m
    pipenet_node_rows = tables["nodes"]
    pipenet_pipe_rows = tables["pipes"]
    pipenet_nozzle_rows = tables["nozzles"]
    pipenet_valve_rows = tables["valves"]

    extracted_pipe_rows = []
    for a, b in sorted(path_edges):
        ed = G.get_edge_data(a, b)
        extracted_pipe_rows.append({
            "Pipe No": ed.get("pipe_no"),
            "From Node": f"N{a:04d}",
            "To Node": f"N{b:04d}",
            "Dia(mm)": ed.get("dia"),
            "Length(CAD unit)": round(ed.get("length"), 2),
            "Length(m)": round(ed.get("length", 0.0) * scale, 4),
            "Layer": ed.get("layer"),
        })

    head_rows = []
    for h in selected_heads:
        head_rows.append({
            "Remote Head No": h["Remote Head No"],
            "Original Head No": h["Head No"],
            "Node": h["Node Name"],
            "X(CAD)": h["X"],
            "Y(CAD)": h["Y"],
            "X(m)": round(h["X"] * scale, 4),
            "Y(m)": round(h["Y"] * scale, 4),
            "Distance from AV(CAD unit)": round(h["Distance from AV(CAD unit)"], 2),
            "Distance from AV(m)": round(h["Distance from AV(CAD unit)"] * scale, 4),
            "Plan distance to farthest(CAD unit)": round(h["Plan distance to farthest"], 2),
            "Layer": h["Layer"],
            "Block": h["Block"],
        })

    pipenet_node_cols = ["node_id", "x", "y", "z", "node_type", "source"]
    pipenet_pipe_cols = ["pipe_id", "from_node", "to_node", "diameter_m", "diameter_label",
                         "length_m", "rise_m", "c_factor", "material", "status"]
    pipenet_nozzle_cols = ["nozzle_id", "input_node", "output_node", "flow_m3s", "status", "library_item"]
    pipenet_valve_cols = ["valve_id", "input_node", "output_node", "valve_type", "target_value"]

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        pd.DataFrame(pipenet_node_rows, columns=pipenet_node_cols).to_excel(writer, sheet_name="network_3d_nodes", index=False)
        pd.DataFrame(pipenet_pipe_rows, columns=pipenet_pipe_cols).to_excel(writer, sheet_name="network_3d_pipes", index=False)
        pd.DataFrame(pipenet_nozzle_rows, columns=pipenet_nozzle_cols).to_excel(writer, sheet_name="network_3d_nozzles", index=False)
        pd.DataFrame(pipenet_valve_rows, columns=pipenet_valve_cols).to_excel(writer, sheet_name="network_3d_valves", index=False)

        pd.DataFrame(head_rows).to_excel(writer, sheet_name="Remote 30 Heads", index=False)
        pd.DataFrame(extracted_pipe_rows).to_excel(writer, sheet_name="Selected Pipes (CAD)", index=False)
        pd.DataFrame(pipe_records).to_excel(writer, sheet_name="All Extracted Pipes", index=False)

        assumptions = pd.DataFrame(
            [
                ["Remote 기준", "알람밸브로부터 배관길이 기준 최장거리"],
                ["헤드 선정", "최원단 헤드를 기준으로 동일세대/인접 헤드 30개 군집 선정"],
                ["동떨어진 헤드 제외", "최원단 헤드와 평면 좌표거리 기준 인접성 우선"],
                ["관경", "배관 인근 숫자 TEXT에서 25/32/40/50/65/80/100/125/150 등 추출"],
                ["CAD 단위→m 스케일", f"{scale} (mm=0.001, m=1.0, inch=0.0254 등)"],
                ["C-factor", f"{settings.c_factor} (강관 기본값)"],
                ["PIPENET 시트", "network_3d_nodes / pipes / nozzles / valves 4종 - pipenet_converter import 호환"],
                ["주의", "CAD 축척, layer/block 명칭, K-factor(flow_m3s)는 실제 도면/스펙에 맞게 후처리 필요"],
            ],
            columns=["항목", "내용"],
        )
        assumptions.to_excel(writer, sheet_name="Assumptions", index=False)


def build_sdf_xml(tables: dict, settings: Remote30Settings, *, title: str = "Remote 30 Auto-Extracted") -> str:
    """PIPENET tables dict → SDF XML 문자열.
    인라인 편집된 tables 로도 호출 가능 (워크벤치 SDF 재생성).
    """
    import xml.etree.ElementTree as ET
    from xml.dom import minidom

    project = ET.Element("Project", version="1.11.0  (3604)")
    network = ET.SubElement(project, "Network-spray")
    ET.SubElement(network, "Title").text = title
    ET.SubElement(network, "Network-description").text = " "

    attrs = ET.SubElement(network, "Attributes")
    units = ET.SubElement(attrs, "Units", type="user-defined")
    ET.SubElement(units, "Length-unit", display="2", precision="2", unit="metres")
    ET.SubElement(units, "Diameter-unit", display="3", precision="3", unit="millimetres")
    ET.SubElement(units, "Pressure-unit", display="0", precision="1", unit="kg-f-cm2-g")
    ET.SubElement(units, "Velocity-unit", display="2", precision="2", unit="m-s")
    ET.SubElement(units, "Volumetric-flow-unit", display="1", precision="1", unit="l-min")

    ET.SubElement(attrs, "Fluid-fixed-user", density="998.2", **{"vapour-pressure": "3.62070676e+205", "viscosity": "1e-09"})
    design = ET.SubElement(
        attrs, "Design-options",
        **{
            "analysis-phase": "1", "design-phase": "0", "design-rules": "NFPA",
            "pressure-equation": "hazen-williams", "specification-type": "user-defined",
        },
    )
    flow_m3s_default = settings.design_flow_per_head_lpm / 60000.0
    ET.SubElement(design, "Nozzle-specification", flowrate=f"{flow_m3s_default:.11f}", label="")
    ET.SubElement(attrs, "Defaults", elevation="0", friction="Unset", **{"nozzle-flowrate": f"{flow_m3s_default:.11f}"})

    # Nodes
    nodes_el = ET.SubElement(network, "Nodes")
    node_label_map = {}
    for idx, row in enumerate(tables["nodes"], 1):
        label = str(idx)
        node_label_map[row["node_id"]] = label
        node_el = ET.SubElement(
            nodes_el, "Node",
            elevation=f"{float(row.get('z', 0.0)):g}",
            **{"io-node": "No"},
            label=label,
        )
        # x/y 는 PIPENET 좌표(metres). build 시 cad_unit_to_m 이미 곱해진 상태.
        # SDF 의 Position 은 mm 단위로 저장 (RV03_NEW 참고).
        x_mm = float(row.get("x", 0.0)) * 1000.0
        y_mm = float(row.get("y", 0.0)) * 1000.0
        ET.SubElement(node_el, "Position", x=f"{x_mm:g}", y=f"{y_mm:g}")

    # Pipes
    pipes_el = ET.SubElement(network, "Pipes")
    for row in tables["pipes"]:
        from_label = node_label_map.get(row["from_node"], "0")
        to_label = node_label_map.get(row["to_node"], "0")
        ET.SubElement(
            pipes_el, "Pipe",
            bore=f"{float(row.get('diameter_m', 0.0)):g}",
            input=from_label,
            label=str(row.get("pipe_id", "P")),
            length=f"{float(row.get('length_m', 0.0)):.3f}",
            output=to_label,
            rise=f"{float(row.get('rise_m', 0.0)):g}",
            **{"roughness-or-c": f"{float(row.get('c_factor', settings.c_factor)):g}"},
            status=str(row.get("status", "normal")),
        )

    # Nozzles
    nozzles_el = ET.SubElement(network, "Nozzles")
    for row in tables["nozzles"]:
        in_label = node_label_map.get(row["input_node"], "0")
        nozzle_label_raw = str(row.get("nozzle_id", ""))
        nozzle_label = nozzle_label_raw.replace("RH", "").lstrip("0") or "1"
        ET.SubElement(
            nozzles_el, "Nozzle",
            input=in_label,
            label=nozzle_label,
            output=str(row.get("output_node", f"@/{nozzle_label}")),
            status=str(row.get("status", "1")),
        )

    # === Graphics 블록 — PIPENET isometric 표시용 ===
    # 답안지 RV03_NEW.sdf 구조 모방. 이 블록 없으면 PIPENET 에서 테이블만 보임.
    graphics = ET.SubElement(network, "Graphics")
    display_opts = ET.SubElement(graphics, "Display-options")
    ET.SubElement(display_opts, "Grid-options",
                  displayed="1", grid="isometric", snap="1", style="points")
    ET.SubElement(display_opts, "Label-display",
                  arrows="1", **{"dec-limit": "0", "dec-limitno": "1", "display-font": "20",
                                 "dry": "0", "fittings": "0", "force-labels": "0",
                                 "force-legend": "1", "forces": "0", "label-all": "0",
                                 "link-labels": "1", "node-labels": "0", "print-font": "48"})
    ET.SubElement(display_opts, "Results-display",
                  **{"flow-arrows": "1", "links": "1", "nodes": "0"})
    ET.SubElement(display_opts, "Background", colour="white")
    ET.SubElement(display_opts, "Line-thickness", line="1")
    ET.SubElement(display_opts, "Tool-tips", on="0")

    link_schemes = ET.SubElement(graphics, "Link-schemes", **{"select-name": "Pipe velocity"})
    for tag in ["Default-link-scheme", "Pipe-bore-scheme", "Pipe-length-scheme",
                "Pipe-mass-flow-scheme", "Pipe-pressure-difference-scheme",
                "Pipe-pressure-gradient-scheme", "Pipe-type-scheme", "Pipe-velocity-scheme",
                "Pipe-volumetric-flow-scheme"]:
        ET.SubElement(link_schemes, tag,
                      **{"auto-classify": "1", "reversed": "0", "use-modulus": "0"})

    node_schemes = ET.SubElement(graphics, "Node-schemes", **{"select-name": "None"})
    for tag in ["Default-node-scheme", "Node-elevation-scheme", "Node-pressure-scheme",
                "Nozzle-calculated-deviation-scheme", "Nozzle-calculated-flow-scheme",
                "Nozzle-calculated-pressure-scheme", "Nozzle-required-flow-scheme",
                "Nozzle-type-scheme"]:
        attrs = {"auto-classify": "1", "reversed": "0", "use-modulus": "0"}
        if tag == "Node-pressure-scheme":
            attrs["graduated"] = "#000000"
            attrs["reversed"] = "1"
        ET.SubElement(node_schemes, tag, **attrs)

    raw = ET.tostring(project, encoding="utf-8")
    pretty = minidom.parseString(raw).toprettyxml(indent="\t", encoding="UTF-8")
    text = pretty.decode("utf-8")
    text = text.replace(
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE Project SYSTEM "spray.dtd">',
        1,
    )
    return text


def _inject_natural_fall_trunk(tables: dict, settings: Remote30Settings) -> dict:
    """자연낙차 옥상수원 → 단위세대 알람밸브 사이의 가상 수직 트렁크 추가.
    답안지 RV03_NEW.sdf 처럼 92.8m + 40.6m 등 큰 rise 의 trunk pipe 를 자동 생성.
    PIPENET 도메인 가정 (DXF 평면도에는 없는 정보).
    """
    fall_h = float(settings.natural_fall_height_m)
    if fall_h <= 0:
        return tables
    n_stages = max(1, int(settings.natural_fall_n_stages))
    bore_mm = float(settings.natural_fall_trunk_bore_mm)
    bore_m = bore_mm / 1000.0
    c = settings.c_factor

    # 기존 알람밸브 node (node_type="source") 찾기
    src_node = next((n for n in tables["nodes"] if n["node_type"] == "source"), None)
    if src_node is None:
        return tables
    src_id = src_node["node_id"]
    src_x = float(src_node["x"]); src_y = float(src_node["y"]); src_z = float(src_node["z"])

    # n_stages 개의 가상 node 를 src 위 fall_h 높이부터 등간격으로 배치
    # 최상단을 새 source (옥상 수원) 로 변경하고, 기존 src 는 junction 으로 강등
    new_nodes = []
    stage_height = fall_h / n_stages
    new_node_ids = []
    for i in range(n_stages):
        nid = f"TRUNK{i+1:02d}"
        # i=0 가 가장 아래(기존 src 바로 위), i=n_stages-1 가 최상단 (수원)
        z = src_z + stage_height * (i + 1)
        node_type = "source" if i == n_stages - 1 else "junction"
        source_tag = "rooftop_reservoir" if i == n_stages - 1 else "trunk"
        new_nodes.append({
            "node_id": nid, "x": src_x, "y": src_y, "z": z,
            "node_type": node_type, "source": source_tag,
        })
        new_node_ids.append(nid)

    # 기존 src 를 junction 으로 강등 (수원 → 알람밸브 라인의 끝점)
    src_node["node_type"] = "junction"
    src_node["source"] = "alarm_valve"

    # 트렁크 pipe 추가: 수원 (TRUNK_top) → ... → 기존 src
    new_pipes = []
    # 최상단 → 중간 → ... → src 순서로 연결
    for i in range(n_stages):
        if i == 0:
            # TRUNK01 → 기존 알람밸브
            from_node = new_node_ids[0]; to_node = src_id
        else:
            from_node = new_node_ids[i]; to_node = new_node_ids[i - 1]
        new_pipes.append({
            "pipe_id": f"TPIPE{i+1:02d}",
            "from_node": from_node, "to_node": to_node,
            "diameter_m": bore_m, "diameter_label": diameter_mm_to_label(bore_mm),
            "length_m": stage_height, "rise_m": -stage_height,
            "c_factor": c, "material": "", "status": "normal",
        })

    # 최종: 새 nodes/pipes 를 기존 위에 추가
    tables["nodes"] = new_nodes + tables["nodes"]
    tables["pipes"] = new_pipes + tables["pipes"]
    return tables


def export_pipenet_csv(tables: dict, out_dir: Path, base_name: str = "network_3d") -> list:
    """PIPENET 4 tables 를 CSV 파일 4개로 저장 (pipenet_converter.export_tables 와 동일 스키마).
    Returns: 생성된 파일 경로 리스트
    """
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    schema = {
        "nodes":   ["node_id", "x", "y", "z", "node_type", "source"],
        "pipes":   ["pipe_id", "from_node", "to_node", "diameter_m", "diameter_label",
                    "length_m", "rise_m", "c_factor", "material", "status"],
        "nozzles": ["nozzle_id", "input_node", "output_node", "flow_m3s", "status", "library_item"],
        "valves":  ["valve_id", "input_node", "output_node", "valve_type", "target_value"],
    }
    for key, cols in schema.items():
        path = out_dir / f"{base_name}_{key}.csv"
        rows = tables.get(key, []) or []
        df = pd.DataFrame(rows, columns=cols)
        df.to_csv(path, index=False, encoding="utf-8-sig")
        written.append(path)
    return written


def _export_sdf(G, node_coords, selected_heads, path_edges, alarm_node, settings: Remote30Settings, out_sdf, tables=None):
    """RV03_NEW.sdf 와 동일한 PIPENET XML 스키마로 출력 (파일로 저장)."""
    if tables is None:
        tables = build_pipenet_tables(G, node_coords, selected_heads, path_edges, alarm_node, settings)
    text = build_sdf_xml(tables, settings)
    Path(out_sdf).write_text(text, encoding="utf-8")
    return  # 아래의 옛 구현은 사용하지 않으므로 즉시 return

    import xml.etree.ElementTree as ET
    from xml.dom import minidom

    scale = settings.cad_unit_to_m

    used_nodes = set()
    for a, b in path_edges:
        used_nodes.add(a)
        used_nodes.add(b)
    used_nodes.add(alarm_node)
    head_nodes_to_id = {h["Node"]: h["Remote Head No"] for h in selected_heads}

    # SDF label 매핑 — RV03_NEW 는 정수 label 사용
    sorted_nodes = sorted(used_nodes)
    node_label_map = {n: str(i + 1) for i, n in enumerate(sorted_nodes)}

    project = ET.Element("Project", version="1.11.0  (3604)")
    network = ET.SubElement(project, "Network-spray")
    ET.SubElement(network, "Title").text = "Remote 30 Auto-Extracted"
    ET.SubElement(network, "Network-description").text = " "

    attrs = ET.SubElement(network, "Attributes")
    units = ET.SubElement(attrs, "Units", type="user-defined")
    ET.SubElement(units, "Length-unit", display="2", precision="2", unit="metres")
    ET.SubElement(units, "Diameter-unit", display="3", precision="3", unit="millimetres")
    ET.SubElement(units, "Pressure-unit", display="0", precision="1", unit="kg-f-cm2-g")
    ET.SubElement(units, "Velocity-unit", display="2", precision="2", unit="m-s")
    ET.SubElement(units, "Volumetric-flow-unit", display="1", precision="1", unit="l-min")

    ET.SubElement(attrs, "Fluid-fixed-user", density="998.2", **{"vapour-pressure": "3.62070676e+205", "viscosity": "1e-09"})
    design = ET.SubElement(
        attrs, "Design-options",
        **{
            "analysis-phase": "1", "design-phase": "0", "design-rules": "NFPA",
            "pressure-equation": "hazen-williams", "specification-type": "user-defined",
        },
    )
    # NFTC 80 LPM @ K=80 = 1.333e-3 m^3/s
    flow_m3s = settings.design_flow_per_head_lpm / 60000.0
    ET.SubElement(design, "Nozzle-specification", flowrate=f"{flow_m3s:.11f}", label="")
    ET.SubElement(
        attrs, "Defaults",
        elevation="0", friction="Unset", **{"nozzle-flowrate": f"{flow_m3s:.11f}"},
    )

    nodes_el = ET.SubElement(network, "Nodes")
    for n in sorted_nodes:
        if n == alarm_node:
            elev = settings.elevation_alarm_m
        elif n in head_nodes_to_id:
            elev = settings.elevation_head_m
        else:
            elev = settings.elevation_head_m
        node_el = ET.SubElement(
            nodes_el, "Node",
            elevation=f"{elev:g}", **{"io-node": "No"}, label=node_label_map[n],
        )
        x_cad, y_cad = node_coords[n]
        ET.SubElement(node_el, "Position", x=f"{x_cad:g}", y=f"{y_cad:g}")

    pipes_el = ET.SubElement(network, "Pipes")
    pipe_label = 0
    for a, b in sorted(path_edges):
        ed = G.get_edge_data(a, b)
        dia_mm = ed.get("dia") or settings.fallback_dia_mm
        bore_m = float(dia_mm) / 1000.0
        length_m = ed.get("length", 0.0) * scale
        pipe_label += 1
        ET.SubElement(
            pipes_el, "Pipe",
            bore=f"{bore_m:g}",
            input=node_label_map[a],
            label=str(pipe_label),
            length=f"{length_m:.3f}",
            output=node_label_map[b],
            rise="0",
            **{"roughness-or-c": f"{settings.c_factor:g}"},
            status="normal",
        )

    nozzles_el = ET.SubElement(network, "Nozzles")
    for h in selected_heads:
        n = h["Node"]
        rh_no = h["Remote Head No"]
        nozzle_label = rh_no.replace("RH", "").lstrip("0") or "1"
        ET.SubElement(
            nozzles_el, "Nozzle",
            input=node_label_map[n],
            label=nozzle_label,
            output=f"@/{nozzle_label}",
            status="1",
        )

    # 정렬된 XML 출력
    raw = ET.tostring(project, encoding="utf-8")
    pretty = minidom.parseString(raw).toprettyxml(indent="\t", encoding="UTF-8")
    # PIPENET 의 DOCTYPE 라인 삽입
    pretty_text = pretty.decode("utf-8")
    pretty_text = pretty_text.replace(
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<?xml version="1.0" encoding="UTF-8"?>\n<!DOCTYPE Project SYSTEM "spray.dtd">',
        1,
    )
    Path(out_sdf).write_text(pretty_text, encoding="utf-8")


def run_remote30_extraction(
    dxf_path,
    alarm_xy: Optional[tuple] = None,
    out_dir: Optional[Path] = None,
    overrides: Optional[dict] = None,
    override_heads: Optional[list] = None,
    override_pipes: Optional[list] = None,
) -> dict:
    """
    서버용 메인 엔트리.
    Returns: { run_id, png_path, xlsx_path, alarm_xy, alarm_source,
               counts: {segments, texts, heads, nodes, edges, attached_heads, selected},
               summary, warnings }
    """
    dxf_path = Path(dxf_path)
    if out_dir is None:
        out_dir = Path("output_remote30")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    settings = Remote30Settings()
    if overrides:
        for k, v in overrides.items():
            if hasattr(settings, k) and v is not None:
                setattr(settings, k, v)

    warnings = []
    run_id = uuid.uuid4().hex[:12]

    doc = ezdxf.readfile(str(dxf_path))
    msp = doc.modelspace()

    segments = _get_line_entities(msp, settings)
    text_items = _get_text_entities(msp, settings)
    if override_heads is not None:
        # 사용자가 워크벤치에서 확정한 헤드 좌표 리스트로 대체 (Layer 매칭 우회)
        heads = []
        for h in override_heads:
            try:
                x = float(h["x"]); y = float(h["y"])
            except (KeyError, TypeError, ValueError):
                continue
            heads.append({
                "xy": (x, y),
                "layer": str(h.get("layer", "<user>")),
                "block": str(h.get("block", h.get("origin", "MANUAL"))),
            })
    else:
        heads = _get_head_candidates(msp, settings)

    # === 누락 보완: INSERT 분해 + 관경 TEXT 근접 LINE ===
    extra_segments, extra_counts = _collect_supplementary_segments(msp, segments, text_items, settings)
    if extra_segments:
        segments.extend(extra_segments)

    # === 사용자가 워크벤치에서 수동 추가한 배관 ===
    user_pipe_count = 0
    if override_pipes:
        for p in override_pipes:
            try:
                x1 = float(p["x1"]); y1 = float(p["y1"])
                x2 = float(p["x2"]); y2 = float(p["y2"])
            except (KeyError, TypeError, ValueError):
                continue
            segments.append({
                "layer": str(p.get("layer", "<user-pipe>")),
                "p1": (x1, y1), "p2": (x2, y2),
                "source": "manual",
            })
            user_pipe_count += 1
    extra_counts["user_pipes"] = user_pipe_count

    if not segments:
        warnings.append("배관 선분이 추출되지 않았습니다. PIPE_LAYER_KEYWORDS 확인 필요")
    if not heads:
        warnings.append("헤드 후보가 추출되지 않았습니다. HEAD_LAYER_KEYWORDS 확인 필요")

    alarm_source = "manual"
    if alarm_xy is None:
        detected = detect_alarm_valve_xy(msp, settings)
        if detected is not None:
            alarm_xy = detected
            alarm_source = "auto"
        else:
            alarm_xy = (0.0, 0.0)
            alarm_source = "fallback_origin"
            warnings.append("알람밸브 자동 검출 실패 — 좌표 (0,0) 사용. 수동 입력 권장")

    G, node_coords, pipe_records = _build_pipe_graph(segments, text_items, settings)
    # === 그래프 closure (degree==1 끝점 자동 snap) ===
    closure_added = _close_graph_gaps(G, node_coords, settings)
    attached_heads = _attach_heads_to_graph(heads, G, node_coords, settings)
    # 미부착 헤드가 많으면 head_to_pipe_tol 을 늘려 2차 시도
    if len(attached_heads) < len(heads) * 0.8:
        original_tol = settings.head_to_pipe_tol
        settings.head_to_pipe_tol = max(settings.head_to_pipe_tol * 2, settings.graph_closure_tol)
        attached_heads = _attach_heads_to_graph(heads, G, node_coords, settings)
        warnings.append(
            f"미부착 헤드 비율 높음 → head_to_pipe_tol 을 {original_tol:.0f} → {settings.head_to_pipe_tol:.0f} 로 자동 증가. "
            f"이후 부착 {len(attached_heads)}/{len(heads)}"
        )
    # === 그래도 부착 못 한 헤드 → 강제 부착 (새 노드 + 가짜 edge) ===
    attached_xys = {(h["X"], h["Y"]) for h in attached_heads}
    orphan_heads = [h for h in heads if h.get("xy") not in attached_xys]
    if orphan_heads:
        forced, fadd = _force_attach_orphan_heads(orphan_heads, G, node_coords, settings)
        if forced:
            attached_heads.extend(forced)
            warnings.append(f"부착 실패 헤드 {len(orphan_heads)}개 → 강제 부착 (가짜 edge {fadd}개 추가)")

    # 알람밸브 노드가 외딴 component에 있는지 진단 — 헤드와 같은 component에 있어야 hydraulic 가능
    if G.number_of_edges() > 0 and attached_heads:
        components = list(nx.connected_components(G))
        # 가장 많은 헤드가 부착된 component를 "main component" 로 정의
        head_node_set = {h["Node"] for h in attached_heads}
        comp_head_counts = [(len(comp & head_node_set), comp) for comp in components]
        comp_head_counts.sort(key=lambda x: -x[0])
        main_head_count, main_comp = comp_head_counts[0]
        av_idx = _nearest_node_id(alarm_xy, node_coords)
        if av_idx not in main_comp:
            warnings.append(
                f"알람밸브 노드가 외딴 component에 있습니다 (해당 component 헤드 0개). "
                f"가장 많은 헤드({main_head_count}개)가 모인 component의 centroid 노드로 알람밸브를 옮겨 재계산합니다."
            )
            xs = [node_coords[n][0] for n in main_comp]
            ys = [node_coords[n][1] for n in main_comp]
            centroid = (sum(xs) / len(xs), sum(ys) / len(ys))
            # main component 안에서 centroid에 가장 가까운 노드 선택
            best = min(main_comp, key=lambda n: math.dist(centroid, node_coords[n]))
            alarm_xy = node_coords[best]
            alarm_source = "fallback_main_component"

    # === 모든 component 를 알람밸브 component 에 강제 연결 (auto-bridge) ===
    # 결과: 모든 헤드 ↔ 알람밸브 path 보장
    bridge_added = 0
    if G.number_of_edges() > 0:
        av_idx_final = _nearest_node_id(alarm_xy, node_coords)
        bridge_added = _bridge_components(G, node_coords, av_idx_final)
        if bridge_added > 0:
            warnings.append(f"끊긴 component {bridge_added}개를 알람밸브에 강제 연결 (가짜 edge 추가)")

    # === edge dia 가 None 인 경우 그래프 구조로 추정 (leaf=25A, 짧음=25A, 그 외=50A) ===
    dia_refined = _refine_edge_diameters(G, settings)
    if dia_refined > 0:
        warnings.append(f"관경 TEXT 미매칭 edge {dia_refined}개에 구조 기반 직경 추정 적용")

    # 자동 zone 추천은 비활성 (사용자가 워크벤치에서 zone 박스 직접 그리는 것이 더 정확)
    # zone_bbox 가 지정돼 있으면 그 zone 안 attached_heads 만 사용 (path expansion 효과 ↑)
    user_zone_applied = False
    if settings.zone_bbox is not None and attached_heads:
        zb = settings.zone_bbox
        attached_in_zone = [h for h in attached_heads
                            if zb[0] <= h["X"] <= zb[2] and zb[1] <= h["Y"] <= zb[3]]
        if len(attached_in_zone) >= settings.remote_head_count:
            attached_heads = attached_in_zone
            user_zone_applied = True
            warnings.append(f"사용자 zone 적용 → zone 안 attached_heads {len(attached_heads)}개")

    if settings.remote_mode == "hydraulic":
        selected, alarm_node, path_edges = _select_remote_hydraulic(
            G, node_coords, attached_heads, alarm_xy, settings
        )
    else:
        selected, alarm_node, path_edges = _select_remote_contiguous(
            G, node_coords, attached_heads, alarm_xy, settings
        )

    # === Path 가지 확장 — 사용자 zone 박스 안일 때만 (zone heads union) ===
    # 답안지처럼 풍부한 가지 그래프를 만들려면 zone 안 모든 헤드의 path 가 필요.
    # zone 이 없으면 selected 30 union 만 (도면 전체 추출 시 trunk 위주가 됨).
    if user_zone_applied and selected and path_edges:
        expanded = set(path_edges)
        for h in attached_heads:
            n = h["Node"]
            if n == alarm_node:
                continue
            try:
                p = nx.shortest_path(G, alarm_node, n, weight="weight")
                for a, b in zip(p[:-1], p[1:]):
                    expanded.add(tuple(sorted((a, b))))
            except nx.NetworkXNoPath:
                continue
        new_added = len(expanded) - len(path_edges)
        if new_added > 0:
            warnings.append(f"Path 가지 확장: zone 안 {len(attached_heads)}개 헤드 union → +{new_added} edges")
        path_edges = expanded

    out_png = out_dir / f"remote30_{run_id}.png"
    out_xlsx = out_dir / f"remote30_{run_id}.xlsx"
    out_sdf = out_dir / f"remote30_{run_id}.sdf"
    out_csv_dir = out_dir / f"remote30_{run_id}_csv"

    sdf_tables = None
    csv_paths = []
    if selected:
        sdf_tables = build_pipenet_tables(G, node_coords, selected, path_edges, alarm_node, settings)

        # === 자연낙차 가상 트렁크 추가 ===
        # 답안지 RV03_NEW 처럼 옥상 수원 ↔ 단위세대 사이의 큰 수직 강하 배관을 시뮬레이션
        if settings.natural_fall_height_m > 0:
            sdf_tables = _inject_natural_fall_trunk(sdf_tables, settings)
            warnings.append(
                f"자연낙차 트렁크 추가: {settings.natural_fall_height_m}m 강하 "
                f"({settings.natural_fall_n_stages} 단, bore={settings.natural_fall_trunk_bore_mm}mm)"
            )
        _plot_extracted_isometric(G, node_coords, selected, alarm_node, path_edges, out_png)
        _export_excel(G, node_coords, pipe_records, selected, path_edges, alarm_node, settings, out_xlsx, tables=sdf_tables)
        if settings.emit_sdf:
            _export_sdf(G, node_coords, selected, path_edges, alarm_node, settings, out_sdf, tables=sdf_tables)
        if settings.emit_csv:
            csv_paths = export_pipenet_csv(sdf_tables, out_csv_dir, base_name=f"remote30_{run_id}")
    else:
        warnings.append("선정 가능한 Remote 30 헤드가 없습니다. 배관/헤드 추출 결과를 확인하세요")

    counts = {
        "segments": len(segments),
        "texts": len(text_items),
        "heads": len(heads),
        "nodes": G.number_of_nodes(),
        "edges": G.number_of_edges(),
        "attached_heads": len(attached_heads),
        "selected": len(selected),
        "supplementary_insert_exploded": extra_counts.get("insert_exploded", 0),
        "supplementary_near_text_line": extra_counts.get("near_text_line", 0),
        "graph_closure_edges": closure_added,
        "auto_bridge_edges": bridge_added,
        "orphan_heads_forced": len(orphan_heads) if orphan_heads else 0,
    }

    summary_rows = []
    for h in selected:
        row = {
            "remote_head_no": h["Remote Head No"],
            "node": h["Node Name"],
            "x": h["X"],
            "y": h["Y"],
            "distance_from_av": round(h["Distance from AV(CAD unit)"], 2),
        }
        if "Total Loss(kgcm2)" in h:
            row["friction_loss_kgcm2"] = round(h["Friction Loss(kgcm2)"], 4)
            row["static_head_kgcm2"] = round(h["Static Head(kgcm2)"], 4)
            row["total_loss_kgcm2"] = round(h["Total Loss(kgcm2)"], 4)
            row["path_length_m"] = round(h["Path Length(m)"], 3)
        summary_rows.append(row)

    # 워크벤치 캔버스 오버레이용 좌표 노출
    alarm_node_xy = list(node_coords[alarm_node]) if (selected and node_coords) else list(alarm_xy)
    selected_heads_xy = [
        {"rh_no": h["Remote Head No"], "x": h["X"], "y": h["Y"],
         "total_loss": h.get("Total Loss(kgcm2)"),
         "distance_av": h.get("Distance from AV(CAD unit)")}
        for h in selected
    ]
    # path_edges_xy: 원본 배관 segment 의 raw 좌표 사용 → 시각화 일탈 방지
    # closure(가짜) edge 는 별도 표시 가능하도록 5번째 값 origin 추가
    path_edges_xy = []
    for a, b in sorted(path_edges):
        ed = G.get_edge_data(a, b) or {}
        rp1 = ed.get("raw_p1")
        rp2 = ed.get("raw_p2")
        if rp1 is None or rp2 is None:
            # raw 가 없는 경우 (이전 데이터) → snap 노드 좌표 fallback
            x1, y1 = node_coords[a]
            x2, y2 = node_coords[b]
        else:
            x1, y1 = rp1
            x2, y2 = rp2
        is_closure = bool(ed.get("is_closure"))
        path_edges_xy.append([x1, y1, x2, y2, "closure" if is_closure else "pipe"])

    return {
        "run_id": run_id,
        "png_path": str(out_png) if selected else None,
        "xlsx_path": str(out_xlsx) if selected else None,
        "sdf_path": str(out_sdf) if (selected and settings.emit_sdf) else None,
        "csv_paths": [str(p) for p in csv_paths] if csv_paths else [],
        "csv_dir": str(out_csv_dir) if (selected and settings.emit_csv) else None,
        "alarm_xy": [alarm_xy[0], alarm_xy[1]],
        "alarm_node_xy": alarm_node_xy,
        "alarm_source": alarm_source,
        "remote_mode": settings.remote_mode,
        "counts": counts,
        "summary": summary_rows,
        "warnings": warnings,
        "selected_heads_xy": selected_heads_xy,
        "path_edges_xy": path_edges_xy,
        "sdf_tables": sdf_tables,
    }
