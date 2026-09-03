# -*- coding: utf-8 -*-
"""KFP 변환 엔진 — 편집 그래프(preflight 후) → Z 있는 실제 그래프 .kfp.

수직 규칙 (2026-08-13 오너):
  메인 = 급수에서 따라감. 호의 열린 곳으로 안 감 (1곳=엘보, 2곳=크로스).
  호가 없으면 갈래를 다 메인. 가지 상승 없음.
  가지면 = 메인 Z + branch_rise (기본 0.3). 호로 열린 갈래만.
  상향식 = 가지면 + ① (기본 0.3)
  하향식 = ①↑(0 이면 안 만듦) · 중간=도면 XY(없거나 ≤0.15m → 0.15m)
            · ②↓ (후렉시블 C·거칠기는 ②만, 체크 시에만)
            ① 위치 = 가지에서 팔이 갈라지는 티. 꺾인 팔 엘보가 아님.
  상하향식 = ①↑0.2 · ②위헤드 0.3(①에 붙음) · ③+X 0.3 · ④↓0.5(후렉시블)
            평면 원 하나 → 헤드 둘. 종류는 편집 head_kinds. 펌프 Z 그대로.
  알람밸브 = 찍은 점에서 아래 ①(기본 2.5) → 알람밸브 → ②(기본 0.5).
            ①=0 이면 안 그림. 찍기 없으면 안 그림.
원본 평면 그래프는 수정하지 않는다. 새 .kfp 만 쓴다.

복사하지 않는 것: DN>=80 메인 휴리스틱, 전부 상향 가정, B1F 경로, 구경분류.
메인/가지 = 호 따라가기만.

사용: python -X utf8 _tmp_kfp_convert.py [--smoke]
      from services.cad_import.convert.engine import convert_to_kfp
실도면은 캐시+유저손질로 평면 그래프를 메모리에 만든 뒤 Desktop/{key}_변환.kfp 에 쓴다.
"""
from __future__ import annotations

import copy
import json
import math
import os
from collections import defaultdict, deque

from services.cad_import.dto import (
    BRANCH_DEFAULT_M as BRANCH_RISE_M,
    COMBO_1_DEFAULT_M as COMBO_1_M,
    COMBO_2_DEFAULT_M as COMBO_2_M,
    COMBO_3_DEFAULT_M as COMBO_3_M,
    COMBO_UP_DEFAULT_M as COMBO_UP_M,
    PENDANT_2_DEFAULT_M as PENDANT_2_M,
    PENDANT_DEFAULT_M as PENDANT_1_M,
    UPRIGHT_DEFAULT_M as UPRIGHT_1_M,
    VALVE_1_DEFAULT_M as VALVE_1_M,
    VALVE_2_DEFAULT_M as VALVE_2_M,
)
from services.cad_import.kinds import normalize_head_kind, require_head_kinds
from services.cad_import.convert.main_walk import (
    ho_to_kfp_units, sit_arcs, snap_seed, walk_main, xf_mm_to_m)
from services.cad_import.convert.preflight import preflight_kfp_convert
from services.cad_import.pipeline.user_net import apply_kind_overrides

PENDANT_ARM_MIN_M = 0.15
VALVE_SPLIT_M = 0.25
# 가지면 대비 Δz. 상하향식은 ①②③④ 뼈대라 단순 Δz 없음.
HEAD_DZ_M = {"상향식": UPRIGHT_1_M, "하향식": -PENDANT_2_M, "상하향식": None}
_CONFIRMED = ("상향식", "하향식", "상하향식")
BLOCKER_PLANAR_MISSING = "planar_kfp_missing"
BLOCKER_UNMAPPED = "unmapped_head_kinds"
BLOCKER_SOURCE_OVERWRITE = "source_overwrite"
BLOCKER_SOURCE_SELECTION = "source_selection_required"


def _resolved_kinds(payload):
    head_kinds = list(payload.get("head_kinds") or [])
    ovs = payload.get("kind_overrides") or []
    if ovs:
        head_kinds = apply_kind_overrides(head_kinds, ovs)
    hcov = payload.get("hcov")
    if hcov is None:
        hcov = payload.get("disks")
    if hcov is not None:
        head_kinds = require_head_kinds(hcov, head_kinds)
    return head_kinds


def _load_kfp(payload):
    if payload.get("kfp") is not None:
        return copy.deepcopy(payload["kfp"]), None
    path = payload.get("kfp_path")
    if path:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f), path
    return None, None


def ensure_planar(payload):
    """kfp/kfp_path 가 없으면 기존 export 로 평면 그래프를 붙인다.

    원본 유저정리 .kfp 는 쓰지 않는다. 그래프가 이미 있으면 그대로.
    """
    payload = dict(payload or {})
    if payload.get("kfp") is not None or payload.get("kfp_path"):
        return payload
    has_graph = payload.get("pts") is not None and payload.get("edges") is not None
    key = payload.get("key")
    if not has_graph and not key:
        return payload
    from services.cad_import.convert.planar import build_planar_graph
    kwargs = {}
    if has_graph:
        kwargs.update(
            pts=payload.get("pts"),
            edges=payload.get("edges"),
            hcov=payload.get("hcov") or payload.get("disks"),
            ups=payload.get("ups"),
            head_kinds=_resolved_kinds(payload),
            user_sources=payload.get("sources"),
            ho=payload.get("ho"),
        )
    built = build_planar_graph(
        key or "out", write=False,
        selected_source=payload.get("selected_source"), **kwargs)
    if not built.get("ok") or built.get("kfp") is None:
        payload["_planar_error"] = (
            built.get("error") or "평면 그래프 .kfp 가 없습니다.")
        payload["_planar_code"] = built.get("code")
        return payload
    payload["kfp"] = built["kfp"]
    if "sources" in built:
        payload["sources"] = list(built.get("sources") or [])
    if not payload.get("node_head_kinds"):
        payload["node_head_kinds"] = built.get("node_head_kinds") or {}
    if not payload.get("head_kinds") and built.get("head_kinds"):
        payload["head_kinds"] = built["head_kinds"]
    if payload.get("hcov") is None and built.get("hcov") is not None:
        payload["hcov"] = built["hcov"]
    if payload.get("origin_mm") is None and built.get("origin_mm"):
        payload["origin_mm"] = built["origin_mm"]
    got_ho = payload.get("ho") or ()
    built_ho = built.get("ho") or ()
    if built_ho and not any(h.get("sa") is not None
                            and h.get("sweep") is not None for h in got_ho):
        payload["ho"] = built["ho"]
    return payload


def _head_nids(nodes):
    return [nid for nid, n in nodes.items()
            if str(n.get("type_id") or "") == "head"]


def _kind_by_head_nid(nodes, payload, head_kinds):
    """헤드 노드 → 종류. 좌표가 같거나, 레코드 수=헤드 수이고 종류가 하나일 때만.

    스케일 변환·전부 상향 가정은 하지 않는다. 못 맞추면 None.
    """
    heads = _head_nids(nodes)
    explicit = payload.get("node_head_kinds") or {}
    if explicit:
        mapped = {str(k): normalize_head_kind(v) for k, v in explicit.items()}
        if heads and all(mapped.get(h) in _CONFIRMED for h in heads):
            return mapped
        return None

    recs = [r for r in (head_kinds or ()) if isinstance(r, dict)]
    kinds = [normalize_head_kind(r.get("kind")) for r in recs]
    uniq = set(kinds)
    if (heads and recs and len(recs) == len(heads)
            and len(uniq) == 1 and next(iter(uniq)) in _CONFIRMED):
        k = next(iter(uniq))
        return {h: k for h in heads}

    if recs and all("c" in r and r["c"] for r in recs):
        mapped, used = {}, set()
        for rec in recs:
            cx, cy = float(rec["c"][0]), float(rec["c"][1])
            hits = []
            for nid in heads:
                c = nodes[nid].get("coords") or [0, 0, 0]
                if abs(float(c[0]) - cx) <= 1e-6 and abs(float(c[1]) - cy) <= 1e-6:
                    hits.append(nid)
            if len(hits) != 1 or hits[0] in used:
                return None
            used.add(hits[0])
            mapped[hits[0]] = normalize_head_kind(rec.get("kind"))
        if heads and all(mapped.get(h) in _CONFIRMED for h in heads):
            return mapped
    return None


def _source_xy_kfp(payload, origin_mm):
    srcs = payload.get("sources") or ()
    if not srcs:
        return None
    rec = srcs[0]
    xy = rec.get("xy") if isinstance(rec, dict) else rec
    if not xy or len(xy) < 2:
        return None
    if origin_mm is None:
        return (float(xy[0]), float(xy[1]))
    return xf_mm_to_m(float(xy[0]), float(xy[1]), origin_mm[0], origin_mm[1])


def _prepare_ho(payload, kfp):
    """호·급수를 kfp 좌표로. payload['ho'](mm)+origin_mm. DXF 안 연다.

    ho 없으면 전부 메인. origin 없으면 ho 는 이미 kfp 단위(스모크).
    """
    ho = payload.get("ho")
    if not ho:
        return None, None
    origin = payload.get("origin_mm")
    src = _source_xy_kfp(payload, origin)
    if origin is None:
        return list(ho), src
    return ho_to_kfp_units(ho, origin[0], origin[1]), src


def _valve_xy_kfp(payload, origin_mm):
    """밸브 찍기 xy → kfp 단위. origin 없으면 이미 kfp 단위(스모크)."""
    out = []
    for rec in payload.get("valve_picks") or ():
        xy = rec.get("xy") if isinstance(rec, dict) else rec
        if not xy or len(xy) < 2:
            continue
        if origin_mm is None:
            out.append((float(xy[0]), float(xy[1])))
        else:
            out.append(xf_mm_to_m(
                float(xy[0]), float(xy[1]), origin_mm[0], origin_mm[1]))
    return out


def _valve_library():
    from editor_core import PipeEditor
    from services.pipenet_import import _ensure_default_libraries
    editor = PipeEditor()
    _ensure_default_libraries(editor)
    return editor


def _stamp_alarm_valve(node, library):
    from domain.node_meta_factory import build_attribute_apply_meta
    from services.library_service import canonicalize_category_id
    coords = list(node.get("coords") or [0.0, 0.0, 0.0])
    nid = node.get("id")
    elev = float(node.get("elevation_m") if node.get("elevation_m") is not None
                 else (coords[2] if len(coords) > 2 else 0.0))
    result = build_attribute_apply_meta(
        copy.deepcopy(node), "Alarm Valve",
        library=library,
        canonicalize_category_id=canonicalize_category_id,
    )
    node.update(result.meta)
    node["id"] = nid
    node["coords"] = coords
    node["elevation_m"] = elev
    if not node.get("fitting_id"):
        fit = library.get_fitting_data_v3("VALVE_ALARM")
        if fit is not None:
            node["fitting_id"] = fit.id


def _xy_on_seg_2d(px, py, ax, ay, bx, by):
    dx, dy = bx - ax, by - ay
    den = dx * dx + dy * dy
    t = 0.0 if den == 0 else max(0.0, min(1.0,
        ((px - ax) * dx + (py - ay) * dy) / den))
    qx, qy = ax + t * dx, ay + t * dy
    d2 = (px - qx) ** 2 + (py - qy) ** 2
    return qx, qy, t, d2


def _walk_main_kfp(kfp, ho, source_xy):
    nodes = kfp["nodes_meta_runtime"]
    pipes = kfp["pipe_data"]
    xy = {nid: (float(n["coords"][0]), float(n["coords"][1]))
          for nid, n in nodes.items()}
    adj = defaultdict(set)
    for p in pipes.values():
        adj[p["start"]].add(p["end"])
        adj[p["end"]].add(p["start"])
    sit_r = max((float(s.get("r") or 0.0) for s in ho), default=0.0)
    node_arcs = sit_arcs(xy, ho, sit_r)
    if not node_arcs:
        return None, "no_arc_seated"
    seed = snap_seed(xy, adj, source_xy, snap=2.5)
    if seed is None:
        return None, "seed_snap_fail"
    return walk_main(xy, adj, node_arcs, seed), None


def _pipe_between(pipes, a, b):
    want = {a, b}
    for pid, p in pipes.items():
        if {p["start"], p["end"]} == want:
            return pid
    return None


def _collinear_xy(a, b, c, ang_tol=5.0):
    """a-b-c 가 일직선(축 각도 여유 5° · 편집 USER_AXIS_ANG 과 같음)."""
    v1x, v1y = b[0] - a[0], b[1] - a[1]
    v2x, v2y = c[0] - b[0], c[1] - b[1]
    n1 = math.hypot(v1x, v1y)
    n2 = math.hypot(v2x, v2y)
    if n1 < 1e-9 or n2 < 1e-9:
        return True
    return (v1x * v2x + v1y * v2y) / (n1 * n2) >= math.cos(math.radians(ang_tol))


def _takeoff_dir(line_id, h_pipes, pipes, nodes):
    """상하향식 하향 팔이 뻗을 «수평» 방향 — 붙은 가지관을 따라간다.

    실제 배관에서 그 팔은 같은 가지에서 딴 니플이므로 가지 축 위에 놓인다.
    붙은 관이 수직이거나(수평 성분 0) 좌표를 못 읽으면 **종전 그대로 +X** 를
    돌려준다 — 방향을 못 정할 때 그림이 갑자기 달라지지 않게.
    """
    cl = (nodes.get(line_id) or {}).get("coords")
    if not cl:
        return (1.0, 0.0)
    for pid in (h_pipes or ()):
        pipe = pipes.get(pid)
        if pipe is None:
            continue
        other = pipe["end"] if pipe["start"] == line_id else pipe["start"]
        co = (nodes.get(other) or {}).get("coords")
        if not co:
            continue
        dx = float(co[0]) - float(cl[0])
        dy = float(co[1]) - float(cl[1])
        h = math.hypot(dx, dy)
        if h > 1e-9:
            return (dx / h, dy / h)
    return (1.0, 0.0)


def _pendant_arm_to_tee(line_id, first, pipes, adj, nodes):
    """헤드에서 꺾인 팔만 따라가 가지 접속점.

    deg>=3 이더라도 들어온 선이 맞은편 관과 일직선이면 가지 몸통이다.
    그때는 직전 노드가 가지 끝 헤드의 접속점이다. 일직선 진행이 없을 때만
    진짜 팔 티로 본다. 이미 만든 Z수직은 평면 이웃에서 제외한다.
    꺾인 팔(비일직선)만 넘어간다. 막다른 기호 획은 None.
    반환: (tee, arm_nodes) arm 은 티에 안 드는 팔 노드.
    """
    def xy(nid):
        c = nodes[nid]["coords"]
        return (float(c[0]), float(c[1]))

    def others(nid, prev):
        out = []
        for pid in adj.get(nid) or ():
            p = pipes[pid]
            o = p["end"] if p["start"] == nid else p["start"]
            if o == prev:
                continue
            # 앞서 처리한 헤드가 만든 Z수직은 같은 XY다. 이것을 갈래로 세면
            # 끝단 헤드 판정이 처리 순서에 따라 달라진다.
            if math.hypot(xy(o)[0] - xy(nid)[0],
                          xy(o)[1] - xy(nid)[1]) <= 1e-9:
                continue
            out.append(o)
        return out

    prev = line_id
    cur = first
    arm = []
    seen = {line_id}
    while cur not in seen:
        seen.add(cur)
        if str((nodes.get(cur) or {}).get("type_id") or "") == "head":
            return None, arm
        nxts = others(cur, prev)
        if len(nxts) >= 2:
            # 끝에서 두 번째 헤드 티를 지나 마지막 헤드로 이어지는 관은
            # incoming과 맞은편 관이 일직선이다. 그 구간을 팔로 올리지 않는다.
            if any(_collinear_xy(xy(prev), xy(cur), xy(nxt))
                   for nxt in nxts):
                if arm and arm[-1] == prev:
                    return prev, arm[:-1]
                return None, arm
            return cur, arm
        if len(nxts) != 1:
            return None, arm
        nxt = nxts[0]
        beyond = others(nxt, cur)
        if (len(beyond) == 1
                and _collinear_xy(xy(cur), xy(nxt), xy(beyond[0]))):
            return cur, arm
        arm.append(cur)
        prev, cur = cur, nxt
    return None, arm


def _apply_vertical(kfp, kind_by_nid, branch_rise, upright_m, pendant_1_m,
                    pendant_2_m, combo_1, combo_2, combo_3, combo_up,
                    flex_c, flex_roughness_mm, head_k, required_pressure_bar,
                    ho=None, source_xy=None,
                    valve_xy=None, valve_1_m=VALVE_1_M, valve_2_m=VALVE_2_M,
                    valve_lib=None, head_active=False, head_spec_name=None):
    """메인→가지 수직 + 헤드 종류별 ①②③④ + 알람밸브 아래 ①②."""
    nodes = kfp["nodes_meta_runtime"]
    pipes = kfp["pipe_data"]
    kfp.setdefault("node_counter", {"N": 0})
    if "pipe_id_counter" not in kfp:
        kfp["pipe_id_counter"] = 0

    def next_node_id():
        n = int(kfp["node_counter"].get("N", 0)) + 1
        while f"N{n}" in nodes:
            n += 1
        kfp["node_counter"]["N"] = n
        return f"N{n}"

    def next_pipe_id():
        n = int(kfp["pipe_id_counter"]) + 1
        while f"P{n}" in pipes:
            n += 1
        kfp["pipe_id_counter"] = n
        return f"P{n}"

    def clone_base(template_id, new_id, coords):
        meta = copy.deepcopy(nodes[template_id])
        meta["id"] = new_id
        meta["coords"] = [float(coords[0]), float(coords[1]), float(coords[2])]
        meta["elevation_m"] = float(coords[2])
        meta["type"] = "기본"
        meta["type_id"] = "base"
        meta["category_id"] = ""
        meta["k_factor_si"] = None
        meta["head_spec_name"] = None
        meta["required_pressure_bar"] = 0.0
        return meta

    def make_vert_pipe(start, end, template_pipe, length):
        return {
            "start": start, "end": end,
            "type": template_pipe.get("type"),
            "diameter": template_pipe.get("diameter"),
            "nominal_mm": template_pipe.get("nominal_mm"),
            "length_m": float(length),
            "equivalent_length": 0.0,
            "C": template_pipe.get("C"),
            "roughness_mm": template_pipe.get("roughness_mm"),
            "fittings": [], "flow_lpm": 0.0, "velocity_mps": 0.0,
            "headloss_m": 0.0,
        }

    def clone_head(template_id, new_id, coords):
        meta = copy.deepcopy(nodes[template_id])
        meta["id"] = new_id
        meta["coords"] = [float(coords[0]), float(coords[1]), float(coords[2])]
        meta["elevation_m"] = float(coords[2])
        meta["type"] = "Head"
        meta["type_id"] = "head"
        if not meta.get("category_id"):
            meta["category_id"] = "head"
        return meta

    def apply_flex(pipe):
        if flex_c is not None:
            pipe["C"] = float(flex_c)
        if flex_roughness_mm is not None:
            pipe["roughness_mm"] = float(flex_roughness_mm)

    def stamp_head(hid):
        if head_k is not None:
            nodes[hid]["k_factor_si"] = float(head_k)
            nodes[hid]["k_factor"] = float(head_k)
        if head_spec_name:
            nodes[hid]["head_spec_name"] = str(head_spec_name)
        if required_pressure_bar is not None:
            nodes[hid]["required_pressure_bar"] = float(required_pressure_bar)
        nodes[hid]["is_active"] = bool(head_active)

    def attach_head_to_line(hid):
        hc = list(nodes[hid]["coords"])
        h_pipes = [pid for pid, pipe in pipes.items()
                   if pipe["start"] == hid or pipe["end"] == hid]
        if not h_pipes or base_tmpl is None:
            return None, None, None
        line_id = next_node_id()
        nodes[line_id] = clone_base(base_tmpl, line_id, hc)
        for pid in h_pipes:
            if pipes[pid]["start"] == hid:
                pipes[pid]["start"] = line_id
            if pipes[pid]["end"] == hid:
                pipes[pid]["end"] = line_id
        return line_id, hc, h_pipes

    def peel_head(hid, dz, flex=False):
        line_id, hc, h_pipes = attach_head_to_line(hid)
        if line_id is None:
            return False
        new_z = float(hc[2]) + float(dz)
        nodes[hid]["coords"] = [hc[0], hc[1], new_z]
        nodes[hid]["elevation_m"] = new_z
        hp0 = pipes[h_pipes[0]]
        dp = next_pipe_id()
        pipes[dp] = make_vert_pipe(line_id, hid, hp0, abs(float(dz)))
        if flex:
            apply_flex(pipes[dp])
        stamp_head(hid)
        return True

    def _floor_arm_length(pipe):
        a = nodes[pipe["start"]]["coords"]
        b = nodes[pipe["end"]]["coords"]
        geom = math.hypot(float(a[0]) - float(b[0]), float(a[1]) - float(b[1]))
        stored = float(pipe.get("length_m") or 0)
        if max(geom, stored) <= PENDANT_ARM_MIN_M + 1e-12:
            pipe["length_m"] = PENDANT_ARM_MIN_M

    def build_pendant(hid, rise_m, drop_m):
        """①=가지 접속점(꺾인 팔이면 티까지). ①=0 이면 가지에 직접. ②만 후렉시블."""
        line_id, hc, h_pipes = attach_head_to_line(hid)
        if line_id is None:
            return False
        for pid in h_pipes:
            _floor_arm_length(pipes[pid])
        rise_m = float(rise_m)
        drop_m = float(drop_m)
        hp0 = pipes[h_pipes[0]]
        if abs(rise_m) > 1e-9:
            adj_now = build_adj()
            raise_nodes = set()
            tee_jobs = []
            for pid in h_pipes:
                pipe = pipes[pid]
                first = (pipe["end"] if pipe["start"] == line_id
                         else pipe["start"])
                tee, arm = _pendant_arm_to_tee(
                    line_id, first, pipes, adj_now, nodes)
                if tee is None:
                    continue
                raise_nodes.update(arm)
                raise_nodes.add(line_id)
                arm_first = arm[-1] if arm else line_id  # 티에 닿은 쪽 마디
                tee_jobs.append((tee, arm_first))
            far_to_rise = {}
            for tee, arm_first in tee_jobs:
                if tee in far_to_rise:
                    continue
                pid = _pipe_between(pipes, tee, arm_first)
                if pid is None:
                    continue
                fc = nodes[tee]["coords"]
                rid = next_node_id()
                nodes[rid] = clone_base(
                    base_tmpl, rid,
                    [fc[0], fc[1], float(fc[2]) + rise_m])
                far_to_rise[tee] = rid
                pipes[next_pipe_id()] = make_vert_pipe(
                    tee, rid, hp0, rise_m)
                if pipes[pid]["start"] == tee:
                    pipes[pid]["start"] = rid
                elif pipes[pid]["end"] == tee:
                    pipes[pid]["end"] = rid
            for nid in raise_nodes:
                c = list(nodes[nid]["coords"])
                c[2] = float(c[2]) + rise_m
                nodes[nid]["coords"] = c
                nodes[nid]["elevation_m"] = c[2]
        lc = nodes[line_id]["coords"]
        new_z = float(lc[2]) - drop_m
        nodes[hid]["coords"] = [lc[0], lc[1], new_z]
        nodes[hid]["elevation_m"] = new_z
        if abs(drop_m) > 1e-9:
            dp = next_pipe_id()
            pipes[dp] = make_vert_pipe(line_id, hid, hp0, abs(drop_m))
            apply_flex(pipes[dp])
        stamp_head(hid)
        return True

    def build_combo(hid):
        """가지면 접속 C. ①↑ ②위헤드(①에 붙음) ③가지를 따라 ④↓헤드.

        ★③ 은 **가지관을 따라** 뻗는다. 종전에는 언제나 +X 였다 — 가지가
          어느 쪽으로 놓였든 하향 팔만 동쪽으로 삐져나와, 화면에서 그 헤드가
          «엉뚱한 데 박힌» 것으로 보였다(실측: 찍은 자리에서 300mm 치우침).
          하향식(`build_pendant`)은 이미 `_pendant_arm_to_tee` 로 실제 배관을
          따라간다 — 상하향식만 이 규칙 밖에 있었다.

        ★수리계산은 종전과 같다. 세로·팔 배관의 길이는 좌표로 재지 않고 입력값
          (①②③④)을 그대로 쓴다(`make_vert_pipe`). 바뀌는 것은 **그림** 이다.
        """
        line_id, hc, h_pipes = attach_head_to_line(hid)
        if line_id is None:
            return False
        x, y, z = float(hc[0]), float(hc[1]), float(hc[2])
        ux, uy = _takeoff_dir(line_id, h_pipes, pipes, nodes)
        kx, ky = x + combo_2 * ux, y + combo_2 * uy
        j_xyz = [x, y, z + combo_1]
        k_xyz = [kx, ky, z + combo_1]
        down_xyz = [kx, ky, z + combo_1 - combo_3]
        up_xyz = [x, y, z + combo_1 + combo_up]
        j_id = next_node_id()
        k_id = next_node_id()
        up_id = next_node_id()
        nodes[j_id] = clone_base(base_tmpl, j_id, j_xyz)
        nodes[k_id] = clone_base(base_tmpl, k_id, k_xyz)
        nodes[up_id] = clone_head(hid, up_id, up_xyz)
        nodes[hid]["coords"] = down_xyz
        nodes[hid]["elevation_m"] = down_xyz[2]
        hp0 = pipes[h_pipes[0]]
        pipes[next_pipe_id()] = make_vert_pipe(line_id, j_id, hp0, combo_1)
        pipes[next_pipe_id()] = make_vert_pipe(j_id, k_id, hp0, combo_2)
        p3 = next_pipe_id()
        pipes[p3] = make_vert_pipe(k_id, hid, hp0, combo_3)
        apply_flex(pipes[p3])
        pipes[next_pipe_id()] = make_vert_pipe(j_id, up_id, hp0, combo_up)
        stamp_head(hid)
        stamp_head(up_id)
        return True

    def build_adj():
        adj = defaultdict(list)
        for pid, pipe in pipes.items():
            adj[pipe["start"]].append(pid)
            adj[pipe["end"]].append(pid)
        return adj

    adj = build_adj()
    nonzero = {
        nid for nid, n in nodes.items()
        if abs(float((n.get("coords") or [0, 0, 0])[2])) > 1e-9
    }
    pump_valve = {
        nid for nid, n in nodes.items()
        if n.get("type_id") in ("pump", "valve")
    }
    protected = set(nonzero) | set(pump_valve)
    changed = True
    while changed:
        changed = False
        for nid in list(protected):
            for pid in adj[nid]:
                pipe = pipes[pid]
                o = pipe["end"] if pipe["start"] == nid else pipe["start"]
                z0 = float(nodes[nid]["coords"][2])
                z1 = float(nodes[o]["coords"][2])
                if (abs(z0 - z1) > 1e-9 or abs(z1) > 1e-9
                        or nodes[o]["type_id"] in ("pump", "valve")):
                    if o not in protected:
                        protected.add(o)
                        changed = True

    walked = None
    walk_reason = None
    if ho and source_xy is not None:
        walked, walk_reason = _walk_main_kfp(kfp, ho, source_xy)
    main_nodes = set()
    junctions = []
    if walked is not None:
        main_nodes = set(walked["main_nodes"]) - protected
        edge_pid = {}
        for pid, pipe in pipes.items():
            a, b = pipe["start"], pipe["end"]
            edge_pid[(a, b) if a < b else (b, a)] = pid
        for m, o in walked["branch_first"]:
            if m in protected or o in protected:
                continue
            pid = edge_pid.get((m, o) if m < o else (o, m))
            if pid is None:
                continue
            junctions.append((m, pid, o))

    raised_nodes = set()
    branch_comps = []
    for m, pid, o in junctions:
        if o in raised_nodes:
            continue
        q = deque([o])
        bn = set()
        while q:
            u = q.popleft()
            if u in bn or u in main_nodes or u in protected:
                continue
            bn.add(u)
            for pid2 in adj[u]:
                pipe2 = pipes[pid2]
                v = pipe2["end"] if pipe2["start"] == u else pipe2["start"]
                if v not in main_nodes and v not in protected and v not in bn:
                    q.append(v)
        if bn & raised_nodes:
            continue
        raised_nodes |= bn
        branch_comps.append((m, pid, o, bn))

    n_vert_branch = 0
    rise_at = {}
    for m, pid, o, bn in branch_comps:
        mc = nodes[m]["coords"]
        mz = float(mc[2])
        rise_z = mz + branch_rise
        rise_id = rise_at.get(m)
        if rise_id is None:
            rise_id = next_node_id()
            tmpl = o if nodes[o]["type_id"] == "base" else next(iter(bn))
            if nodes[tmpl]["type_id"] != "base":
                tmpl = next(
                    (nid for nid, n in nodes.items()
                     if n.get("type_id") == "base"),
                    tmpl)
            nodes[rise_id] = clone_base(tmpl, rise_id, [mc[0], mc[1], rise_z])
            rise_at[m] = rise_id
            pipes[next_pipe_id()] = make_vert_pipe(
                m, rise_id, pipes[pid], branch_rise)
            n_vert_branch += 1
        for nid in bn:
            c = list(nodes[nid]["coords"])
            c[2] = float(c[2]) + branch_rise
            nodes[nid]["coords"] = c
            nodes[nid]["elevation_m"] = float(c[2])
        bp = pipes[pid]
        if bp["start"] == m:
            bp["start"] = rise_id
        elif bp["end"] == m:
            bp["end"] = rise_id
        else:
            raise RuntimeError(f"junction pipe {pid} not attached to main {m}")

    adj = build_adj()
    heads = _head_nids(nodes)
    n_heads_peeled = 0
    n_vert_head = 0
    n_combo = 0
    base_tmpl = next(
        (nid for nid, n in nodes.items() if n.get("type_id") == "base"), None)

    for hid in heads:
        if hid in protected:
            continue
        kind = kind_by_nid.get(hid)
        if kind == "상향식":
            if peel_head(hid, float(upright_m), flex=False):
                n_heads_peeled += 1
                n_vert_head += 1
        elif kind == "하향식":
            if build_pendant(hid, pendant_1_m, pendant_2_m):
                n_heads_peeled += 1
                n_vert_head += 1
        elif kind == "상하향식":
            if build_combo(hid):
                n_combo += 1
                n_vert_head += 1

    def split_at_xy(px, py):
        best = None
        for pid, pipe in pipes.items():
            a = nodes[pipe["start"]]["coords"]
            b = nodes[pipe["end"]]["coords"]
            qx, qy, t, d2 = _xy_on_seg_2d(
                px, py, float(a[0]), float(a[1]), float(b[0]), float(b[1]))
            if best is None or d2 < best[0]:
                best = (d2, pid, qx, qy, t)
        if best is None or best[0] > VALVE_SPLIT_M * VALVE_SPLIT_M:
            return None
        _d2, pid, qx, qy, t = best
        pipe = pipes[pid]
        a = nodes[pipe["start"]]["coords"]
        b = nodes[pipe["end"]]["coords"]
        qz = float(a[2]) + t * (float(b[2]) - float(a[2]))
        if math.hypot(float(a[0]) - qx, float(a[1]) - qy) <= 1e-4:
            return pipe["start"]
        if math.hypot(float(b[0]) - qx, float(b[1]) - qy) <= 1e-4:
            return pipe["end"]
        tmpl = pipe["start"]
        if nodes[tmpl].get("type_id") != "base":
            if nodes[pipe["end"]].get("type_id") == "base":
                tmpl = pipe["end"]
            elif base_tmpl is not None:
                tmpl = base_tmpl
        nid = next_node_id()
        nodes[nid] = clone_base(tmpl, nid, [qx, qy, qz])
        old_end = pipe["end"]
        old_len = float(pipe.get("length_m") or 0)
        len1 = math.sqrt((float(a[0]) - qx) ** 2 + (float(a[1]) - qy) ** 2
                         + (float(a[2]) - qz) ** 2)
        len2 = math.sqrt((float(b[0]) - qx) ** 2 + (float(b[1]) - qy) ** 2
                         + (float(b[2]) - qz) ** 2)
        tot = len1 + len2
        if old_len > 1e-12 and tot > 1e-12:
            pipe["length_m"] = old_len * len1 / tot
            len2s = old_len * len2 / tot
        else:
            pipe["length_m"] = len1
            len2s = len2
        pipe["end"] = nid
        pipes[next_pipe_id()] = make_vert_pipe(nid, old_end, pipe, len2s)
        return nid

    n_valve = 0
    v1 = float(valve_1_m)
    v2 = float(valve_2_m)
    if abs(v1) > 1e-9:
        for xy in valve_xy or ():
            nid = split_at_xy(float(xy[0]), float(xy[1]))
            if nid is None or nodes[nid].get("type_id") == "head":
                continue
            inc = [pid for pid, p in pipes.items()
                   if p["start"] == nid or p["end"] == nid]
            if not inc:
                continue
            tmpl_p = pipes[inc[0]]
            pc = nodes[nid]["coords"]
            x, y, z = float(pc[0]), float(pc[1]), float(pc[2])
            vid = next_node_id()
            tmpl = base_tmpl if base_tmpl is not None else nid
            nodes[vid] = clone_base(tmpl, vid, [x, y, z - v1])
            if valve_lib is not None:
                _stamp_alarm_valve(nodes[vid], valve_lib)
            else:
                nodes[vid]["type"] = "Alarm Valve"
                nodes[vid]["type_id"] = "valve"
                nodes[vid]["category_id"] = "alarm_valve"
                nodes[vid]["fitting_id"] = "VALVE_ALARM"
                nodes[vid]["is_active"] = True
            pipes[next_pipe_id()] = make_vert_pipe(nid, vid, tmpl_p, abs(v1))
            if abs(v2) > 1e-9:
                bid = next_node_id()
                nodes[bid] = clone_base(tmpl, bid, [x, y, z - v1 - v2])
                pipes[next_pipe_id()] = make_vert_pipe(vid, bid, tmpl_p, abs(v2))
            n_valve += 1

    for hid in _head_nids(nodes):
        stamp_head(hid)

    present = {kind_by_nid.get(h) for h in heads} - {None}
    missing_offsets = sorted(
        k for k in present if k not in ("상향식", "하향식", "상하향식"))
    return {
        "n_tees": len({m for m, _, _, _ in branch_comps}),
        "n_branch_comps": len(branch_comps),
        "n_raised": len(raised_nodes),
        "n_heads_peeled": n_heads_peeled,
        "n_combo": n_combo,
        "n_vert_branch": n_vert_branch,
        "n_vert_head": n_vert_head,
        "n_valve": n_valve,
        "n_heads": len(_head_nids(nodes)),
        "n_heads_in": len(heads),
        "protected": sorted(protected),
        "missing_offsets": missing_offsets,
        "main_walk": walked is not None,
        "main_walk_reason": walk_reason,
    }


def convert_to_kfp(payload, out_path=None, *,
                   branch_rise_m=BRANCH_RISE_M, head_dz_m=None,
                   upright_1_m=None, pendant_1_m=None, pendant_2_m=None,
                   combo_1_m=None, combo_2_m=None, combo_3_m=None,
                   combo_takeoff_m=None, combo_up_m=None,
                   flex_c=None, flex_roughness_mm=None,
                   head_k=None, required_pressure_bar=None,
                   min_pressure_bar=None,
                   valve_1_m=None, valve_2_m=None,
                   head_active=False, head_spec_name=None):
    """편집 그래프 payload → Z 있는 .kfp. preflight 실패면 변환하지 않는다.

    payload: head_kinds, kind_overrides, hcov|disks,
             kfp 또는 kfp_path (평면 그래프). node_head_kinds 선택.
    길이 칸이 비면 기본값. 밸브 ①②는 찍힌 점이 있을 때만 그린다.
    하향 ①=0 이면 중간 배관이 가지에 직접 붙는다.
    후렉시블 C·거칠기는 하향 ②·상하향 ④ 만. None 이면 안 넣는다.
    """
    _ = combo_takeoff_m
    if required_pressure_bar is None:
        required_pressure_bar = min_pressure_bar
    if upright_1_m is not None:
        upright = float(upright_1_m)
    elif head_dz_m is not None and head_dz_m.get("상향식") is not None:
        upright = float(head_dz_m["상향식"])
    else:
        upright = UPRIGHT_1_M
    p1 = float(pendant_1_m if pendant_1_m is not None else PENDANT_1_M)
    p2 = float(pendant_2_m if pendant_2_m is not None else PENDANT_2_M)
    c1 = float(combo_1_m if combo_1_m is not None else COMBO_1_M)
    c2 = float(combo_2_m if combo_2_m is not None else COMBO_2_M)
    c3 = float(combo_3_m if combo_3_m is not None else COMBO_3_M)
    c_up = float(combo_up_m if combo_up_m is not None else COMBO_UP_M)
    v1 = float(valve_1_m if valve_1_m is not None else VALVE_1_M)
    v2 = float(valve_2_m if valve_2_m is not None else VALVE_2_M)
    payload = payload or {}
    pf = preflight_kfp_convert(payload)
    empty = {"ok": False, "path": None, "kfp": None, "preflight": pf,
             "blockers": list(pf["blockers"]),
             "diagnostics": list(pf.get("diagnostics") or []),
             "stats": None}
    if not pf["ok"]:
        return empty
    srcs = payload.get("sources") or ()
    if len(srcs) > 1:
        empty["blockers"] = [{
            "code": BLOCKER_SOURCE_SELECTION,
            "message": "급수원이 여러 개입니다. 변환할 급수원 하나를 지정하세요.",
        }]
        return empty

    kfp, src_path = _load_kfp(payload)
    if kfp is None:
        empty["blockers"] = [{
            "code": payload.get("_planar_code") or BLOCKER_PLANAR_MISSING,
            "message": payload.get("_planar_error") or "평면 그래프 .kfp 가 없습니다.",
        }]
        return empty
    if (out_path and src_path
            and os.path.normcase(os.path.abspath(out_path))
            == os.path.normcase(os.path.abspath(src_path))):
        empty["blockers"] = [{
            "code": BLOCKER_SOURCE_OVERWRITE,
            "message": "원본 .kfp 는 수정하지 않습니다.",
        }]
        return empty

    head_kinds = _resolved_kinds(payload)
    kind_by_nid = _kind_by_head_nid(kfp["nodes_meta_runtime"], payload,
                                    head_kinds)
    if kind_by_nid is None:
        empty["blockers"] = [{
            "code": BLOCKER_UNMAPPED,
            "message": "헤드 종류를 그래프 노드에 맞추지 못했습니다.",
        }]
        return empty

    ho, src_xy = _prepare_ho(payload, kfp)
    valve_xy = _valve_xy_kfp(payload, payload.get("origin_mm"))
    valve_lib = _valve_library() if valve_xy else None
    stats = _apply_vertical(
        kfp, kind_by_nid, float(branch_rise_m),
        upright, p1, p2, c1, c2, c3, c_up,
        flex_c, flex_roughness_mm, head_k, required_pressure_bar,
        ho=ho, source_xy=src_xy,
        valve_xy=valve_xy, valve_1_m=v1, valve_2_m=v2, valve_lib=valve_lib,
        head_active=head_active, head_spec_name=head_spec_name)
    if out_path:
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(kfp, f, ensure_ascii=False, indent=2)
    return {
        "ok": True, "path": out_path, "kfp": kfp, "preflight": pf,
        "blockers": [], "diagnostics": list(pf.get("diagnostics") or []),
        "stats": stats,
    }


def convert_drawing(key, out_path=None, *, selected_source=None, **kwargs):
    """저장본(캐시+유저손질) → 평면 그래프(메모리) → 변환 .kfp.

    유저정리(평면) 파일은 쓰지 않는다. 기본 출력은 Desktop/{key}_변환.kfp.
    급수원이 2개 이상이면 selected_source(태그 또는 번호)가 필요하다.
    """
    from services.cad_import.convert.planar import (
        build_planar_graph, default_out as planar_out)
    from services.cad_import.pipeline.handoff import default_edits_dir
    from services.cad_import.pipeline.user_net import user_edit_valve_picks
    if out_path is None:
        out_path = os.path.join(os.path.expanduser("~"), "Desktop",
                                f"{key}_변환.kfp")
    empty = {"ok": False, "path": None, "kfp": None, "preflight": None,
             "blockers": [], "diagnostics": [], "stats": None}
    if (os.path.normcase(os.path.abspath(out_path))
            == os.path.normcase(os.path.abspath(planar_out(key)))):
        empty["blockers"] = [{
            "code": BLOCKER_SOURCE_OVERWRITE,
            "message": "원본 .kfp 는 수정하지 않습니다.",
        }]
        return empty
    built = build_planar_graph(
        key, write=False, selected_source=selected_source)
    if not built.get("ok") or built.get("kfp") is None:
        empty["blockers"] = [{
            "code": built.get("code") or BLOCKER_PLANAR_MISSING,
            "message": built.get("error") or "평면 그래프 .kfp 가 없습니다.",
        }]
        return empty
    payload = {
        "key": key,
        "kfp": built["kfp"],
        "node_head_kinds": built.get("node_head_kinds") or {},
        "head_kinds": built.get("head_kinds") or [],
        "hcov": built.get("hcov"),
        "origin_mm": built.get("origin_mm"),
        "ho": list(built.get("ho") or ()),
        "sources": list(built.get("sources") or ()),
        "valve_picks": user_edit_valve_picks(key, default_edits_dir()),
    }
    return convert_to_kfp(payload, out_path, **kwargs)


def _smoke_planar():
    """복도+가지줄 둘. 호 없음. 전부 Z=0 평면."""
    def node(i, x, y, tid="base"):
        z = 0.0
        rec = {"id": i, "coords": [float(x), float(y), z], "elevation_m": z,
               "type": "기본", "type_id": tid, "category_id": "",
               "k_factor_si": None, "head_spec_name": None,
               "required_pressure_bar": 0.0}
        if tid == "head":
            rec["type"] = "Head"
            rec["k_factor_si"] = 80.0
        elif tid == "pump":
            rec["type"] = "펌프"
        elif tid == "valve":
            rec["type"] = "밸브"
        return rec

    def pipe(a, b, length):
        return {"start": a, "end": b, "type": "KSD3507", "diameter": 27.6,
                "nominal_mm": 25, "length_m": float(length),
                "equivalent_length": 0.0, "C": 120.0, "roughness_mm": 0.1,
                "fittings": [], "flow_lpm": 0.0, "velocity_mps": 0.0,
                "headloss_m": 0.0}

    nodes = dict((
        ("PUMP", node("PUMP", 0.0, 0.0, "pump")),
        ("VALVE", node("VALVE", 0.0, 1.0, "valve")),
        ("A", node("A", 1.0, 0.0)),
        ("B", node("B", 2.0, 0.0)),
        ("C", node("C", 3.0, 0.0)),
        ("HB", node("HB", 2.0, 1.0)),
        ("B1", node("B1", 1.5, 2.0)),
        ("B2", node("B2", 1.0, 3.0)),
        ("B3", node("B3", 2.5, 2.0)),
        ("B4", node("B4", 3.0, 3.0)),
        ("H1", node("H1", 1.5, 2.0, "head")),
        ("H2", node("H2", 1.0, 3.0, "head")),
        ("H3", node("H3", 2.5, 2.0, "head")),
        ("H4", node("H4", 3.0, 3.0, "head")),
        ("HC", node("HC", 3.0, 1.0)),
        ("C1", node("C1", 2.5, -1.0)),
        ("C2", node("C2", 2.0, -2.0)),
        ("C3", node("C3", 3.5, -1.0)),
        ("C4", node("C4", 4.0, -2.0)),
        ("H5", node("H5", 2.5, -1.0, "head")),
        ("H6", node("H6", 2.0, -2.0, "head")),
        ("H7", node("H7", 3.5, -1.0, "head")),
        ("H8", node("H8", 4.0, -2.0, "head")),
    ))
    pipes = {
        "PV": pipe("PUMP", "VALVE", 1.0),
        "SUP": pipe("PUMP", "A", 1.0),
        "M1": pipe("A", "B", 3.0),
        "M2": pipe("B", "C", 3.0),
        "STEM_B": pipe("B", "HB", 3.0),
        "BL1": pipe("HB", "B1", 3.0),
        "BL1B": pipe("B1", "B2", 3.0),
        "BL2": pipe("HB", "B3", 3.0),
        "BL2B": pipe("B3", "B4", 3.0),
        "D1": pipe("B1", "H1", 0.5),
        "D2": pipe("B2", "H2", 0.5),
        "D3": pipe("B3", "H3", 0.5),
        "D4": pipe("B4", "H4", 0.5),
        "STEM_C": pipe("C", "HC", 3.0),
        "CL1": pipe("HC", "C1", 3.0),
        "CL1B": pipe("C1", "C2", 3.0),
        "CL2": pipe("HC", "C3", 3.0),
        "CL2B": pipe("C3", "C4", 3.0),
        "D5": pipe("C1", "H5", 0.5),
        "D6": pipe("C2", "H6", 0.5),
        "D7": pipe("C3", "H7", 0.5),
        "D8": pipe("C4", "H8", 0.5),
    }
    return {"nodes_meta_runtime": nodes, "pipe_data": pipes,
            "node_counter": {"N": 0}, "pipe_id_counter": 0}


def smoke():
    planar = _smoke_planar()
    src_z = {
        nid: list(n["coords"]) for nid, n in planar["nodes_meta_runtime"].items()
    }
    heads_xy = [
        {"c": [1.5, 2.0], "head_r": 0.15, "kind": "상향식"},
        {"c": [1.0, 3.0], "head_r": 0.15, "kind": "상향식"},
        {"c": [2.5, 2.0], "head_r": 0.15, "kind": "상향식"},
        {"c": [3.0, 3.0], "head_r": 0.15, "kind": "상향식"},
        {"c": [2.5, -1.0], "head_r": 0.15, "kind": "상향식"},
        {"c": [2.0, -2.0], "head_r": 0.15, "kind": "상향식"},
        {"c": [3.5, -1.0], "head_r": 0.15, "kind": "상향식"},
        {"c": [4.0, -2.0], "head_r": 0.15, "kind": "하향식"},
    ]
    blocked = convert_to_kfp({
        "kfp": planar,
        "head_kinds": [{"c": [1.5, 2.0], "head_r": 0.15, "kind": "미지정"}],
        "hcov": [(1.5, 2.0, 0.15)],
    })
    assert blocked["ok"] is False
    assert blocked["blockers"][0]["code"] == "unconfirmed_heads"
    assert blocked["kfp"] is None
    assert planar["nodes_meta_runtime"]["H1"]["coords"][2] == 0.0

    missing = convert_to_kfp({"head_kinds": heads_xy})
    assert missing["ok"] is False
    assert missing["blockers"][0]["code"] == BLOCKER_PLANAR_MISSING
    two_src = convert_to_kfp({
        "kfp": planar, "head_kinds": heads_xy,
        "sources": [{"tag": "Z1", "xy": [0.0, 0.0]},
                    {"tag": "Z2", "xy": [10.0, 0.0]}],
    })
    assert two_src["ok"] is False
    assert two_src["blockers"][0]["code"] == BLOCKER_SOURCE_SELECTION
    kept = {"kfp": planar, "head_kinds": heads_xy}
    assert ensure_planar(kept)["kfp"] is planar

    out = convert_to_kfp({"kfp": planar, "head_kinds": heads_xy},
                         flex_c=140.0)
    assert out["ok"] is True
    n2 = out["kfp"]["nodes_meta_runtime"]
    assert planar["nodes_meta_runtime"]["H1"]["coords"][2] == 0.0
    assert n2["PUMP"]["coords"] == src_z["PUMP"]
    assert n2["VALVE"]["coords"] == src_z["VALVE"]
    assert abs(float(n2["B"]["coords"][2])) < 1e-9
    assert abs(float(n2["C"]["coords"][2])) < 1e-9
    assert abs(float(n2["HB"]["coords"][2])) < 1e-9
    assert abs(float(n2["H1"]["coords"][2]) - 0.3) < 1e-9
    assert abs(float(n2["H7"]["coords"][2]) - 0.3) < 1e-9
    h8 = n2["H8"]["coords"]
    assert abs(float(h8[0]) - 4.0) < 1e-9
    assert abs(float(h8[2])) < 1e-9
    assert out["stats"]["missing_offsets"] == []
    assert out["stats"]["n_heads_peeled"] == 8
    assert out["stats"]["n_combo"] == 0
    assert out["stats"]["n_vert_branch"] == 0
    assert out["stats"]["n_valve"] == 0
    assert out["stats"]["main_walk"] is False
    flex_drop = [p for p in out["kfp"]["pipe_data"].values()
                 if p.get("C") == 140.0]
    assert len(flex_drop) == 1
    assert abs(float(flex_drop[0]["length_m"]) - 0.3) < 1e-9

    p0 = convert_to_kfp(
        {"kfp": planar, "head_kinds": heads_xy}, pendant_1_m=0.0)
    assert abs(float(p0["kfp"]["nodes_meta_runtime"]["H8"]["coords"][2])
               + 0.3) < 1e-9

    def _n(i, x, y, tid="base"):
        rec = {"id": i, "coords": [float(x), float(y), 0.0], "elevation_m": 0.0,
               "type": "기본", "type_id": tid, "category_id": "",
               "k_factor_si": None, "head_spec_name": None,
               "required_pressure_bar": 0.0}
        if tid == "head":
            rec["type"] = "Head"
            rec["k_factor_si"] = 80.0
        return rec

    def _p(a, b, length):
        return {"start": a, "end": b, "type": "KSD3507", "diameter": 27.6,
                "nominal_mm": 25, "length_m": float(length),
                "equivalent_length": 0.0, "C": 120.0, "roughness_mm": 0.1,
                "fittings": [], "flow_lpm": 0.0, "velocity_mps": 0.0,
                "headloss_m": 0.0}

    short_kfp = {
        "nodes_meta_runtime": {
            "T": _n("T", 0.0, 0.0),
            "H": _n("H", 0.10, 0.0, "head"),
        },
        "pipe_data": {"ARM": _p("T", "H", 0.10)},
        "node_counter": {"N": 0}, "pipe_id_counter": 0,
    }
    short = convert_to_kfp({
        "kfp": short_kfp,
        "node_head_kinds": {"H": "하향식"},
        "head_kinds": [{"c": [0.10, 0.0], "kind": "하향식"}],
    })
    assert short["ok"] is True
    assert abs(float(short["kfp"]["pipe_data"]["ARM"]["length_m"])
               - 0.15) < 1e-9

    bent_kfp = {
        "nodes_meta_runtime": {
            "A": _n("A", 0.0, 0.0),
            "T": _n("T", 1.0, 0.0),
            "B": _n("B", 2.0, 0.0),
            "E": _n("E", 1.0, 1.0),
            "H": _n("H", 2.0, 1.0, "head"),
        },
        "pipe_data": {
            "AT": _p("A", "T", 1.0),
            "TB": _p("T", "B", 1.0),
            "TE": _p("T", "E", 1.0),
            "EH": _p("E", "H", 1.0),
        },
        "node_counter": {"N": 0}, "pipe_id_counter": 0,
    }
    bent = convert_to_kfp({
        "kfp": bent_kfp,
        "node_head_kinds": {"H": "하향식"},
        "head_kinds": [{"c": [2.0, 1.0], "kind": "하향식"}],
    })
    assert bent["ok"] is True
    bn = bent["kfp"]["nodes_meta_runtime"]
    assert abs(float(bn["T"]["coords"][2])) < 1e-9
    assert abs(float(bn["A"]["coords"][2])) < 1e-9
    assert abs(float(bn["E"]["coords"][2]) - 0.3) < 1e-9
    verts_at_e = [p for p in bent["kfp"]["pipe_data"].values()
                  if {p["start"], p["end"]} & {"E"}
                  and abs(float(p.get("length_m") or 0) - 0.3) < 1e-9
                  and abs(float(bn[p["start"]]["coords"][0])
                          - float(bn[p["end"]]["coords"][0])) < 1e-9
                  and abs(float(bn[p["start"]]["coords"][1])
                          - float(bn[p["end"]]["coords"][1])) < 1e-9]
    assert verts_at_e == []
    t_xy_verts = [p for p in bent["kfp"]["pipe_data"].values()
                  if abs(float(p.get("length_m") or 0) - 0.3) < 1e-9
                  and (p["start"] == "T" or p["end"] == "T")]
    assert len(t_xy_verts) == 1

    # 가지 끝 2헤드: T의 옆 헤드를 먼저 처리해 Z수직이 생겨도, 마지막
    # 헤드로 가는 T-ET 관통 가지는 팔로 오인해 올리지 않는다.
    end2_kfp = {
        "nodes_meta_runtime": {
            "A": _n("A", -1.0, 0.0),
            "T": _n("T", 0.0, 0.0),
            "ET": _n("ET", 1.0, 0.0),
            "ES": _n("ES", 0.0, 1.0),
            "HS": _n("HS", 0.0, 1.2, "head"),  # 먼저 처리되는 옆 헤드
            "HT": _n("HT", 1.2, 0.0, "head"),  # 마지막 헤드
        },
        "pipe_data": {
            "AT": _p("A", "T", 1.0),
            "TET": _p("T", "ET", 1.0),
            "ETHT": _p("ET", "HT", 0.2),
            "TES": _p("T", "ES", 1.0),
            "ESHS": _p("ES", "HS", 0.2),
        },
        "node_counter": {"N": 0}, "pipe_id_counter": 0,
    }
    end2 = convert_to_kfp({
        "kfp": end2_kfp,
        "node_head_kinds": {"HS": "하향식", "HT": "하향식"},
        "head_kinds": [
            {"c": [0.0, 1.2], "kind": "하향식"},
            {"c": [1.2, 0.0], "kind": "하향식"},
        ],
    })
    assert end2["ok"] is True
    en = end2["kfp"]["nodes_meta_runtime"]
    ep = end2["kfp"]["pipe_data"]
    assert abs(float(en["T"]["coords"][2])) < 1e-9
    assert abs(float(en["ET"]["coords"][2])) < 1e-9
    through = next(p for p in ep.values()
                   if {p["start"], p["end"]} == {"T", "ET"})
    assert abs(float(en[through["start"]]["coords"][2])
               - float(en[through["end"]]["coords"][2])) < 1e-9
    assert abs(float(en["ES"]["coords"][2]) - 0.3) < 1e-9

    combo_kinds = list(heads_xy)
    combo_kinds[-1] = {"c": [4.0, -2.0], "head_r": 0.15, "kind": "상하향식"}
    combo_out = convert_to_kfp(
        {"kfp": planar, "head_kinds": combo_kinds},
        flex_c=150.0, flex_roughness_mm=0.2)
    assert combo_out["ok"] is True
    assert combo_out["stats"]["n_combo"] == 1
    assert combo_out["stats"]["n_heads_in"] == 8
    assert combo_out["stats"]["n_heads"] == 9
    cn = combo_out["kfp"]["nodes_meta_runtime"]
    h8c = cn["H8"]["coords"]
    assert abs(float(h8c[0]) - 4.3) < 1e-9
    assert abs(float(h8c[1]) + 2.0) < 1e-9
    assert abs(float(h8c[2]) + 0.3) < 1e-9
    ups = [n for n in cn.values()
           if n.get("type_id") == "head"
           and abs(float(n["coords"][2]) - 0.5) < 1e-9]
    assert len(ups) == 1
    uc = ups[0]["coords"]
    assert abs(float(uc[0]) - 4.0) < 1e-9
    assert abs(float(uc[1]) + 2.0) < 1e-9
    flex3 = [p for p in combo_out["kfp"]["pipe_data"].values()
             if p.get("C") == 150.0
             and abs(float(p.get("roughness_mm") or 0) - 0.2) < 1e-9]
    assert len(flex3) == 1
    assert abs(float(flex3[0]["length_m"]) - 0.5) < 1e-9
    # 호 따라가기: B에서 위(+Y)만 열림 → 가지만 상승. 구경 분류 안 탐.
    ho_b = [{"cx": 1.0, "cy": 0.0, "r": 0.2, "sa": 135.0, "sweep": 270.0}]
    walk_kfp = {
        "nodes_meta_runtime": {
            "A": _n("A", 0.0, 0.0),
            "B": _n("B", 1.0, 0.0),
            "C": _n("C", 2.0, 0.0),
            "D": _n("D", 1.0, 1.0),
            "H": _n("H", 1.0, 1.0, "head"),
        },
        "pipe_data": {
            "AB": _p("A", "B", 1.0),
            "BC": _p("B", "C", 1.0),
            "BD": _p("B", "D", 1.0),
            "DH": _p("D", "H", 0.2),
        },
        "node_counter": {"N": 0}, "pipe_id_counter": 0,
    }
    walked = convert_to_kfp({
        "kfp": walk_kfp,
        "ho": ho_b,
        "sources": [{"xy": [0.0, 0.0]}],
        "node_head_kinds": {"H": "상향식"},
        "head_kinds": [{"c": [1.0, 1.0], "kind": "상향식"}],
    })
    assert walked["ok"] is True
    assert walked["stats"]["main_walk"] is True
    wn = walked["kfp"]["nodes_meta_runtime"]
    assert abs(float(wn["A"]["coords"][2])) < 1e-9
    assert abs(float(wn["C"]["coords"][2])) < 1e-9
    assert abs(float(wn["D"]["coords"][2]) - 0.3) < 1e-9

    ho_cross = [
        {"cx": 1.0, "cy": 0.0, "r": 0.2, "sa": 315.0, "sweep": 90.0},
        {"cx": 1.0, "cy": 0.0, "r": 0.2, "sa": 135.0, "sweep": 90.0},
    ]
    cross_kfp = {
        "nodes_meta_runtime": {
            "A": _n("A", 0.0, 0.0),
            "B": _n("B", 1.0, 0.0),
            "C": _n("C", 2.0, 0.0),
            "D": _n("D", 1.0, 1.0),
            "E": _n("E", 1.0, -1.0),
            "H1": _n("H1", 1.0, 1.0, "head"),
            "H2": _n("H2", 1.0, -1.0, "head"),
        },
        "pipe_data": {
            "AB": _p("A", "B", 1.0),
            "BC": _p("B", "C", 1.0),
            "BD": _p("B", "D", 1.0),
            "BE": _p("B", "E", 1.0),
            "DH": _p("D", "H1", 0.2),
            "EH": _p("E", "H2", 0.2),
        },
        "node_counter": {"N": 0}, "pipe_id_counter": 0,
    }
    crossed = convert_to_kfp({
        "kfp": cross_kfp,
        "ho": ho_cross,
        "sources": [{"xy": [0.0, 0.0]}],
        "node_head_kinds": {"H1": "상향식", "H2": "상향식"},
        "head_kinds": [
            {"c": [1.0, 1.0], "kind": "상향식"},
            {"c": [1.0, -1.0], "kind": "상향식"},
        ],
    })
    assert crossed["ok"] is True
    assert crossed["stats"]["main_walk"] is True
    assert crossed["stats"]["n_vert_branch"] == 1
    assert crossed["stats"]["n_tees"] == 1
    xn = crossed["kfp"]["nodes_meta_runtime"]
    xp = crossed["kfp"]["pipe_data"]
    assert abs(float(xn["B"]["coords"][2])) < 1e-9
    assert abs(float(xn["D"]["coords"][2]) - 0.3) < 1e-9
    assert abs(float(xn["E"]["coords"][2]) - 0.3) < 1e-9
    b_nbrs = set()
    for p in xp.values():
        if p["start"] == "B":
            b_nbrs.add(p["end"])
        elif p["end"] == "B":
            b_nbrs.add(p["start"])
    assert "A" in b_nbrs and "C" in b_nbrs
    assert "D" not in b_nbrs and "E" not in b_nbrs
    rise = [n for n in b_nbrs if n not in ("A", "C")]
    assert len(rise) == 1
    rid = rise[0]
    assert abs(float(xn[rid]["coords"][2]) - 0.3) < 1e-9
    r_nbrs = set()
    for p in xp.values():
        if p["start"] == rid:
            r_nbrs.add(p["end"])
        elif p["end"] == rid:
            r_nbrs.add(p["start"])
    assert "D" in r_nbrs and "E" in r_nbrs

    valve_kfp = {
        "nodes_meta_runtime": {
            "A": _n("A", 0.0, 0.0),
            "B": _n("B", 2.0, 0.0),
            "H": _n("H", 2.0, 0.0, "head"),
        },
        "pipe_data": {
            "AB": _p("A", "B", 2.0),
            "BH": _p("B", "H", 0.2),
        },
        "node_counter": {"N": 0}, "pipe_id_counter": 0,
    }
    vout = convert_to_kfp({
        "kfp": valve_kfp,
        "node_head_kinds": {"H": "상향식"},
        "head_kinds": [{"c": [2.0, 0.0], "kind": "상향식"}],
        "valve_picks": [{"xy": [1.0, 0.0]}],
    }, valve_1_m=2.5, valve_2_m=0.5)
    assert vout["ok"] is True
    assert vout["stats"]["n_valve"] == 1
    vn = vout["kfp"]["nodes_meta_runtime"]
    valves = [n for n in vn.values() if n.get("type_id") == "valve"]
    assert len(valves) == 1
    vc = valves[0]["coords"]
    assert abs(float(vc[0]) - 1.0) < 1e-3
    assert abs(float(vc[2]) + 2.5) < 1e-9
    assert valves[0].get("category_id") == "alarm_valve"
    bottoms = [n for n in vn.values()
               if n.get("type_id") == "base"
               and abs(float(n["coords"][2]) + 3.0) < 1e-9]
    assert len(bottoms) == 1
    assert abs(float(vn["A"]["coords"][2])) < 1e-9
    none_v = convert_to_kfp({
        "kfp": {
            "nodes_meta_runtime": {
                "A": _n("A", 0.0, 0.0),
                "H": _n("H", 1.0, 0.0, "head"),
            },
            "pipe_data": {"AH": _p("A", "H", 1.0)},
            "node_counter": {"N": 0}, "pipe_id_counter": 0,
        },
        "node_head_kinds": {"H": "상향식"},
        "head_kinds": [{"c": [1.0, 0.0], "kind": "상향식"}],
    }, valve_1_m=2.5, valve_2_m=0.5)
    assert none_v["ok"] is True
    assert none_v["stats"]["n_valve"] == 0

    # ho(mm)+origin → kfp. key 만 있고 ho 없으면 DXF 안 열고 전부 메인.
    ho_mm, src = _prepare_ho({
        "ho": [{"cx": 0.0, "cy": 0.0, "r": 200.0, "sa": 135.0, "sweep": 270.0}],
        "origin_mm": (0.0, 1000.0),
        "sources": [{"xy": [0.0, 0.0]}],
    }, None)
    assert abs(ho_mm[0]["cx"] - 1.0) < 1e-9
    assert abs(ho_mm[0]["cy"] - 0.0) < 1e-9
    assert abs(ho_mm[0]["r"] - 0.2) < 1e-9
    assert ho_mm[0]["sa"] == 135.0
    assert abs(src[0] - 1.0) < 1e-9
    assert abs(src[1] - 0.0) < 1e-9
    assert _prepare_ho({"key": "__no_dxf__"}, None) == (None, None)

    from services.cad_import.pipeline.stage1 import explode
    from services.cad_import.pipeline.flow import ho_from_spots, ho_from_spec
    w, _n = explode({}, [{
        "t": "ARC", "8": "SP", "10": 0.0, "20": 0.0, "40": 100.0,
        "50": 0.0, "51": 180.0,
    }], {})
    assert len(w.arcs) == 1 and w.arc_ang[0] is not None
    assert abs(w.arc_ang[0][0]) < 1e-9
    assert abs(w.arc_ang[0][1] - 180.0) < 1e-9
    assert ho_from_spots([
        {"k": "호", "cx": 1, "cy": 2, "r": 3, "sa": 10, "sweep": 20},
        {"k": "원", "cx": 0, "cy": 0, "r": 1},
    ]) == [{"cx": 1.0, "cy": 2.0, "r": 3.0, "sa": 10.0, "sweep": 20.0}]
    assert ho_from_spec({"material_picks": []}) == []
    assert ho_from_spec({
        "format": "v2",
        "ho": [{"cx": 1, "cy": 2, "r": 3, "sa": 10, "sweep": 20}],
    }) == [{"cx": 1.0, "cy": 2.0, "r": 3.0, "sa": 10.0, "sweep": 20.0}]
    assert ho_from_spec({
        "format": "v2",
        "ho": [{"cx": 1, "cy": 2, "r": 3}],
    }) == []

    print("gate: convert preflight_block=미지정 "
          "planar_missing=no_kfp ensure_passthrough "
          "no_ho_branch_z=0 upright=0.3 pendant=0 "
          "pendant_1_zero combo_two_heads=9 flex_pendant2·combo4 "
          "arm_min=0.15 pump_valve_unchanged source_unmodified "
          "arc_walk_elbow same_m_vert=1 ho_mm_origin no_key_dxf "
          "pendant_1_at_tee end2_through_branch_flat "
          "multi_source_blocked")


def run_convert_cli(argv=None):
    """python _tmp_kfp_convert.py [--smoke] [KEY [OUT]] [--source Z1]"""
    import sys
    argv = list(sys.argv[1:] if argv is None else argv)
    if "--smoke" in argv:
        smoke()
        return 0
    selected = None
    positional = []
    i = 0
    while i < len(argv):
        a = argv[i]
        if a == "--source":
            i += 1
            if i >= len(argv):
                print("[KFP 변환] --source 값이 없습니다.")
                return 1
            selected = argv[i]
        elif a.startswith("--source="):
            selected = a.split("=", 1)[1]
        elif a.startswith("--"):
            print(f"[KFP 변환] 알 수 없는 옵션 {a}")
            return 1
        else:
            positional.append(a)
        i += 1
    if not positional:
        smoke()
        return 0
    key = positional[0]
    dest = positional[1] if len(positional) > 1 else None
    result = convert_drawing(key, dest, selected_source=selected)
    if not result["ok"]:
        for b in result.get("blockers") or []:
            print(f"[KFP 변환] {b.get('message') or b.get('code')}")
        return 1
    print(f"[KFP 변환] {result['path']}")
    st = result.get("stats") or {}
    print(f"  가지수직 {st.get('n_vert_branch')} · "
          f"벗김 {st.get('n_heads_peeled')}/"
          f"{st.get('n_heads_in')} · "
          f"상하향 {st.get('n_combo')} · "
          f"헤드 {st.get('n_heads')}")
    reason = st.get("main_walk_reason")
    if reason:
        kr = {"seed_snap_fail": "급수 스냅 실패",
              "no_arc_seated": "호가 배관 노드에 앉지 않음"}.get(reason, reason)
        print(f"  ★호가 있지만 메인/가지 구분을 하지 못했습니다 — {kr}")
    return 0


if __name__ == "__main__":
    import sys
    raise SystemExit(run_convert_cli(sys.argv[1:]))
