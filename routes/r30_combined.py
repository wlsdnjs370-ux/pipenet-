# -*- coding: utf-8 -*-
"""Remote 30 통합(combined) 도메인 — 평면+아이소메트릭 결합 SDF 빌드/재빌드/미리보기.

`대조 서버.py` 에서 `register(app, ...)` 로 등록. 공유 헬퍼·전역 상태는 main 에
두고 참조 주입, combined 전용 지오메트리 헬퍼만 이 모듈에 동봉. 라이저 상수
(_RISER_*) 를 참조하는 _remap_riser_to_head_av 는 register 안에 중첩.
라우트 본문·엔드포인트명 원본 그대로.
"""
from __future__ import annotations

import json
import math
import warnings
from pathlib import Path

from flask import jsonify, request
from werkzeug.utils import secure_filename


def _bake_isometric_node_coords(nodes: list[dict], iso_z_scale: float = 1.0,
                                no_lift_labels: set | None = None) -> None:
    """통합망 노드 dict 의 (x,y) 를 30° 등각투영 좌표로 in-place 변환.

    등각 형태로 보이는 SDF/KFP/HAS 출력을 위해 emit 전에 적용한다. 노드 좌표는
    표시 전용(수리계산은 length·elevation 사용)이라 결과는 불변. 공식은
    has_converter.emit_has(isometric=True) 와 동일: X=(x−y)·cos30,
    Y=(x+y)·sin30 + (elev−eMid)·lift. lift 는 평면 대각선의 절반에 정규화.

    no_lift_labels 노드(라이저/기계실 계통도)는 lift 를 건너뛴다 — schematic y 가
    이미 수직을 인코딩하므로 elevation lift 를 다시 더하면 이중부호로 계통도가
    구부러진다. 헤드 z-돌출은 여기서 적용하지 않는다(평면 Y 를 기울여 가지배관을
    꼬이게 함; 3D 프리뷰·KFP/HAS 는 display_z 로 별도 돌출).
    """
    if not nodes:
        return
    COS30, SIN30 = 0.8660254037844387, 0.5
    no_lift = {str(l) for l in (no_lift_labels or set())}
    xs = [float(n.get("x", 0) or 0) for n in nodes]
    ys = [float(n.get("y", 0) or 0) for n in nodes]
    elevs = [float(n.get("elevation", 0) or 0) for n in nodes]
    e_min, e_max = min(elevs), max(elevs)
    e_mid = (e_min + e_max) / 2.0
    e_range = e_max - e_min
    diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) if xs else 0.0
    lift = (diag * 0.5 * iso_z_scale / e_range) if e_range > 0 else 0.0
    for n in nodes:
        x = float(n.get("x", 0) or 0)
        y = float(n.get("y", 0) or 0)
        e = float(n.get("elevation", 0) or 0)
        _lift = 0.0 if str(n.get("label")) in no_lift else (e - e_mid) * lift
        n["x"] = (x - y) * COS30
        n["y"] = (x + y) * SIN30 + _lift


def _tidy_head_plane_layout(nodes, pipes, root_label, exclude_labels):
    """헤드평면(가지·교차배관)을 **균일 격자 tree-packing**으로 재배치 — 표시 (x,y) 만.

    추출 원본을 그대로 두거나 edge 지배성분만 스냅하면, 형제 가지 서브트리의 폭 구간이
    겹쳐 등각투영 시 가지배관이 서로 꼬이고 겹쳐 보였다. 해법: AV(root)부터 스패닝
    트리를 만든 뒤 각 노드에서 자식 서브트리를 4직교 방향(직진→좌/우→후진)으로 팬아웃
    하고, **각 서브트리의 실제 bounding box 만큼 간격을 확보**해 배치한다. 형제 서브트리
    폭 구간이 완전히 분리되므로 평면 교차 0, 모든 segment 가 0°/90° → 30° 등각투영에서
    30°/150° 격자가 되어 **등각도도 깔끔**해진다.
      · root 는 **실위치 고정**, 각 자식은 실제 방위(부모→자식 벡터)에 가장 가까운 축에
        1:1 배정 → 평면도의 방향감(동/서/남/북)이 유지돼 통합망이 평면도와 닮은 형태.
      · 간격 STEP 은 실 도면 대표 edge 길이(중앙값) → 스케일 보존(붕괴·압축 없음).

    표시 좌표만 바꾼다. 파이프 length·elevation 등 수리값은 emit 단계에서 좌표와
    분리되어 직렬화되므로(emit_sdf 가 p["length"] 사용, 좌표거리 무관) **수리계산 결과
    불변**. 라이저/기계실 노드(exclude_labels)는 손대지 않는다. nodes in-place 수정,
    반환: 재배치된 노드 수.
    """
    from collections import defaultdict as _dd, deque as _deque
    import math as _math

    by_label = {str(n["label"]): n for n in nodes}
    root = str(root_label)
    if root not in by_label:
        return 0
    excl = {str(l) for l in exclude_labels} - {root}
    movable = {lbl for lbl in by_label if lbl not in excl and lbl != root}
    if not movable:
        return 0
    tree_set = movable | {root}

    def _xy(lbl):
        nd = by_label[lbl]
        return float(nd["x"]), float(nd["y"])

    adj = _dd(set)
    for p in pipes:
        a, b = str(p.get("in", "")), str(p.get("out", ""))
        if a in tree_set and b in tree_set and a != b:
            adj[a].add(b); adj[b].add(a)
    if not adj.get(root):
        return 0  # AV 가 헤드평면과 안 이어짐 — 건드리지 않음

    # 스패닝 트리 (BFS, root 부터). children·order 확보, 도달 못한 노드는 원위치 유지.
    children = _dd(list)
    order = [root]
    q = _deque([root]); seen = {root}
    while q:
        u = q.popleft()
        for v in sorted(adj[u]):
            if v not in seen:
                seen.add(v); children[u].append(v); order.append(v); q.append(v)

    # 균일 격자 tree-packing: 각 서브트리를 4직교 방향으로 팬아웃하고, 실제 bounding
    # box 만큼 간격을 확보해 배치 → 형제 서브트리 폭 구간이 완전히 분리되어 교차 0.
    # 간격 STEP 은 실 도면의 대표 edge 길이(중앙값)로 잡아 스케일을 보존(붕괴 방지).
    size = {}
    for u in reversed(order):
        size[u] = 1 + sum(size[c] for c in children[u])

    edge_lens = []
    for u in order:
        ux, uy = _xy(u)
        for c in children[u]:
            vx, vy = _xy(c)
            edge_lens.append(_math.hypot(vx - ux, vy - uy))
    edge_lens.sort()
    STEP = edge_lens[len(edge_lens) // 2] if edge_lens else 1.0
    if STEP <= 0:
        STEP = 1.0
    PAD = STEP  # 균일 간격

    OPP = {"+x": "-x", "-x": "+x", "+y": "-y", "-y": "+y"}
    PERP = {"+x": ["+y", "-y"], "-x": ["+y", "-y"],
            "+y": ["+x", "-x"], "-y": ["+x", "-x"]}
    AXES = ("+x", "+y", "-x", "-y")
    AX_ANG = {"+x": 0.0, "+y": 90.0, "-x": 180.0, "-y": 270.0}

    def _ang_gap(a, b):
        d = abs(a - b) % 360.0
        return min(d, 360.0 - d)

    from itertools import permutations as _perms

    def _assign_dirs(node, incoming):
        """각 자식을 실제 방위(부모→자식 벡터)에 가장 가까운 축에 1:1 배정.
        가용 축 = 전체 4개(root) 또는 incoming 제외 3개. 자식 수 ≤ 가용 축이면 축이
        겹치지 않는 배정 중 방위 오차 합이 최소인 것을 택함(트리 최대차수 4 → 항상 성립).
        예외적으로 자식이 더 많으면 크기순 fallback(직진→좌/우→후진)."""
        kids = children[node]
        avail = list(AXES) if incoming is None else [ax for ax in AXES if ax != incoming]
        ux, uy = _xy(node)
        ang = {c: _math.degrees(_math.atan2(_xy(c)[1] - uy, _xy(c)[0] - ux)) % 360.0
               for c in kids}
        if len(kids) <= len(avail):
            best, best_cost = None, float("inf")
            for combo in _perms(avail, len(kids)):
                cost = sum(_ang_gap(ang[kids[i]], AX_ANG[combo[i]])
                           for i in range(len(kids)))
                if cost < best_cost:
                    best_cost, best = cost, combo
            return dict(zip(kids, best))
        base = list(avail) if incoming is None else [OPP[incoming]] + PERP[incoming] + [incoming]
        ordered = sorted(kids, key=lambda c: size[c], reverse=True)
        return {c: (base[i] if i < len(base) else base[-1]) for i, c in enumerate(ordered)}

    def _layout(node, incoming):
        """node 를 원점에 두고 서브트리 배치. 반환 (pos, bbox[minx,miny,maxx,maxy])."""
        lpos = {node: (0.0, 0.0)}
        bbox = [0.0, 0.0, 0.0, 0.0]
        if not children[node]:
            return lpos, bbox
        assign = _assign_dirs(node, incoming)
        # 큰 서브트리부터 배치(간격 균일). 축이 자식마다 유일하므로 순서는 교차에 무관.
        for c in sorted(children[node], key=lambda c: size[c], reverse=True):
            dr = assign[c]
            csub, cbb = _layout(c, OPP[dr])
            if dr == "+x":
                base = bbox[2] + PAD + STEP; off = (base - cbb[0], 0.0)
            elif dr == "-x":
                base = bbox[0] - PAD - STEP; off = (base - cbb[2], 0.0)
            elif dr == "+y":
                base = bbox[3] + PAD + STEP; off = (0.0, base - cbb[1])
            else:  # -y
                base = bbox[1] - PAD - STEP; off = (0.0, base - cbb[3])
            for n, (x, y) in csub.items():
                nx, ny = x + off[0], y + off[1]
                lpos[n] = (nx, ny)
                bbox[0] = min(bbox[0], nx); bbox[1] = min(bbox[1], ny)
                bbox[2] = max(bbox[2], nx); bbox[3] = max(bbox[3], ny)
        return lpos, bbox

    rx, ry = _xy(root)
    import sys as _sys
    _old_limit = _sys.getrecursionlimit()
    _sys.setrecursionlimit(max(_old_limit, len(order) + 1000))
    try:
        pos, _ = _layout(root, None)   # root: 가용 축 4개, 자식마다 실제 방위로 배정
    finally:
        _sys.setrecursionlimit(_old_limit)

    # 패킹 결과를 원 도면 스팬에 맞춰 균일 스케일 — 형태·간격비·교차0 불변(스케일 무관),
    # 절대 스케일만 정합해 라이저/기계실(exclude, 미이동)과의 크기 붕괴/과확장 방지.
    orig_xs = [_xy(u)[0] for u in seen]; orig_ys = [_xy(u)[1] for u in seen]
    orig_span = max(max(orig_xs) - min(orig_xs), max(orig_ys) - min(orig_ys))
    pk_xs = [p[0] for p in pos.values()]; pk_ys = [p[1] for p in pos.values()]
    pk_span = max(max(pk_xs) - min(pk_xs), max(pk_ys) - min(pk_ys))
    s = (orig_span / pk_span) if (pk_span > 0 and orig_span > 0) else 1.0

    moved = 0
    for u, (x, y) in pos.items():
        if u == root:
            continue
        nd = by_label[u]
        nd["x"], nd["y"] = x * s + rx, y * s + ry   # root 실위치 고정(pos[root]=(0,0))
        moved += 1
    return moved


def _build_combined_geometry(combined, riser, riser_labels, head_label_set,
                             machine_room_labels, pump_junction_label,
                             is_pump, head_orientation, head_z_frac) -> dict:
    """통합망(헤드+라이저) 캔버스 시각화용 geometry dict.

    machine_room_at_bottom: 펌프방식이면 수원/기계실이 망 최하부 → 3D 아이소뷰가 Z 방향
    (아래로)을 뒤집는다. 고가수조(gravity, 기본)면 수원이 옥상(위). DXF 에 z 가 없어
    토폴로지 autoSpread 로 Z 를 만들므로 방향만 이 플래그로 결정한다.
    head_labels: nozzle 부착(input) 노드 — 클라이언트가 여기서 짧은 니플 스텁을 ±z 로 그린다.
    """
    return {
        "av_node_label": riser.av_node_label,
        "riser_labels": list(riser_labels),
        "machine_room_labels": machine_room_labels,
        "pump_junction_label": pump_junction_label,
        "machine_room_at_bottom": is_pump,
        "machine_room_plan_edges": combined.machine_room_plan_edges,
        "head_labels": sorted(head_label_set),
        "head_orientation": head_orientation,
        "head_z_frac": head_z_frac,
        "nodes": [
            {"label": str(n["label"]),
             "x": float(n.get("x", 0)), "y": float(n.get("y", 0)),
             "z": float(n.get("elevation", 0)),
             "io": n.get("io_node", "No"),
             # 층 단위 편집용 태그 — 계통도 라이저 노드가 어느 층인지(있을 때만).
             # 에디터가 층별 노드 분리/삭제/재연결에 사용하고 rebuild 왕복에서 보존한다.
             "floor": n.get("floor"),
             "floor_idx": (int(n["floor_idx"]) if n.get("floor_idx") is not None else None),
             # 아이소 3D 방수 시뮬레이션용 — Input 노드 공급압(Pa). 없으면 None.
             "pressure_pa": (float(n["pressure_pa"])
                             if n.get("pressure_pa") is not None else None)}
            for n in combined.nodes
        ],
        "pipes": [
            {"label": str(p.get("label", "")),
             "in": str(p["in"]), "out": str(p["out"]),
             "dia": p.get("dia", 0),
             # 아이소 3D 방수 시뮬레이션용 하젠-윌리엄스 파라미터.
             "length": float(p.get("length", 0) or 0),   # m
             "c": float(p.get("c", 120) or 120),          # Hazen-Williams C
             "elev": float(p.get("elev", 0) or 0),        # 상승고 m (in→out)
             # 계산(유속) 오버레이 — annotate_pipe_velocity 가 stamp (표시 전용).
             "flow_lpm": p.get("flow_lpm"),
             "velocity_mps": p.get("velocity_mps"),
             "v_limit": p.get("v_limit"),
             "v_over": p.get("v_over")}
            for p in combined.pipes
        ],
        "pumps": [
            {"label": str(p["label"]), "in": str(p["in"]), "out": str(p["out"])}
            for p in combined.pumps
        ],
        "valves": [
            {"label": str(v["label"]), "in": str(v["in"]), "out": str(v["out"]),
             "target_pa": v.get("target_value", 0)}
            for v in combined.valves
        ],
    }


def _build_roles_sidecar(combined, riser_labels, head_label_set, av_node_label,
                         machine_room_labels, pump_junction_label,
                         is_pump, head_orientation, head_z_frac) -> dict:
    """포맷 미리보기(round-trip)용 역할 사이드카.

    KFP/SDF/HAS emit→재파싱은 라벨을 N1..Nn 으로 개명하고 라이저/헤드/AV 구조 메타를
    잃는다. 노드 순서별 원본 라벨(combined.nodes 순서 = parse 순서, 검증됨)과 역할 집합을
    저장해두면 미리보기에서 개명 라벨로 재매핑해 동일 모양을 복원할 수 있다.
    pumps/valves 는 round-trip 시 _common_network_to_geometry 가 비우므로 원본 라벨로 보존.
    """
    return {
        "order_labels": [str(n["label"]) for n in combined.nodes],
        "order_io": [str(n.get("io_node", "No")) for n in combined.nodes],
        "riser_labels": sorted(riser_labels),
        "head_labels": sorted(head_label_set),
        "av_node_label": av_node_label,
        "machine_room_labels": list(machine_room_labels),
        "pump_junction_label": pump_junction_label,
        "head_orientation": head_orientation,
        "head_z_frac": head_z_frac,
        "machine_room_at_bottom": is_pump,
        "pumps": [{"label": str(p["label"]), "in": str(p["in"]), "out": str(p["out"])}
                  for p in combined.pumps],
        "valves": [{"label": str(v["label"]), "in": str(v["in"]), "out": str(v["out"]),
                    "target_pa": v.get("target_value", 0)} for v in combined.valves],
    }


def _patch_combined_from_geometry(combined, geom: dict) -> None:
    """캐시된 CombinedTables 를 브라우저 편집 geometry 로 in-place 패치한다.

    편집 geometry(state.combined_geometry)는 노드(label,x,y,z,io,pressure_pa)와
    배관(label,in,out,dia,length,c,elev)만 담는다. fittings/equipment/nozzle 유량 등
    리치 필드는 combined 원본에 보존돼 있으므로, 여기서는 노드/배관만 덮어쓰고
    삭제된 노드를 참조하는 nozzle/fitting/pump/valve 는 잘라낸다(SDF 무결성 유지).

    좌표계: x,y = mm(int) / elevation·length·elev = m. 클라이언트와 동일.
    """
    g_nodes = {str(n.get("label")): n for n in (geom.get("nodes") or []) if n.get("label") is not None}
    g_pipes = {str(p.get("label")): p for p in (geom.get("pipes") or []) if p.get("label")}

    def _f(v, d=0.0):
        try:
            return float(v)
        except (TypeError, ValueError):
            return float(d)

    # ── 노드: 유지/갱신 + 신규 추가, 편집본에 없는 노드는 삭제 ──
    kept_nodes, seen = [], set()
    for n in combined.nodes:
        lbl = str(n.get("label"))
        gn = g_nodes.get(lbl)
        if gn is None:
            continue  # 편집본에서 삭제됨
        n["x"] = int(round(_f(gn.get("x", n.get("x", 0)))))
        n["y"] = int(round(_f(gn.get("y", n.get("y", 0)))))
        n["elevation"] = _f(gn.get("z", n.get("elevation", 0)))
        if gn.get("io") is not None:
            n["io_node"] = gn["io"]
        if "pressure_pa" in gn:
            n["pressure_pa"] = gn["pressure_pa"]
        # 층 태그 — 편집본이 값을 실어 보내면 반영(분리 노드가 부모 층 상속 등), None 이면 제거.
        if "floor" in gn:
            if gn["floor"] is None:
                n.pop("floor", None)
            else:
                n["floor"] = gn["floor"]
        if "floor_idx" in gn:
            if gn["floor_idx"] is None:
                n.pop("floor_idx", None)
            else:
                n["floor_idx"] = int(gn["floor_idx"])
        kept_nodes.append(n)
        seen.add(lbl)
    for lbl, gn in g_nodes.items():
        if lbl in seen:
            continue
        new_node = {
            "label": lbl,
            "x": int(round(_f(gn.get("x", 0)))),
            "y": int(round(_f(gn.get("y", 0)))),
            "elevation": _f(gn.get("z", 0)),
            "io_node": gn.get("io", "No"),
            "pressure_pa": gn.get("pressure_pa"),
        }
        if gn.get("floor") is not None:
            new_node["floor"] = gn["floor"]
        if gn.get("floor_idx") is not None:
            new_node["floor_idx"] = int(gn["floor_idx"])
        kept_nodes.append(new_node)
    combined.nodes = kept_nodes
    valid = {str(n.get("label")) for n in combined.nodes}

    # 신규 배관 기본 type — 기존 배관 컨벤션을 따른다(없으면 "Pipe").
    default_type = "Pipe"
    for p in combined.pipes:
        if p.get("type"):
            default_type = p["type"]
            break

    # ── 배관: 유지/갱신 + 신규 추가, 삭제/끊긴 끝점 제거 ──
    kept_pipes, seen_p = [], set()
    for p in combined.pipes:
        lbl = str(p.get("label", ""))
        gp = g_pipes.get(lbl)
        if gp is None:
            continue  # 편집본에서 삭제됨
        p["in"] = str(gp.get("in", p.get("in")))
        p["out"] = str(gp.get("out", p.get("out")))
        p["dia"] = int(round(_f(gp.get("dia", p.get("dia", 0)))))
        p["length"] = _f(gp.get("length", p.get("length", 0)))
        p["elev"] = _f(gp.get("elev", p.get("elev", 0)))
        if gp.get("c") is not None:
            p["c"] = gp["c"]
        seen_p.add(lbl)
        if p["in"] in valid and p["out"] in valid:
            kept_pipes.append(p)
    for lbl, gp in g_pipes.items():
        if lbl in seen_p:
            continue
        pin, pout = str(gp.get("in", "")), str(gp.get("out", ""))
        if pin not in valid or pout not in valid:
            continue
        kept_pipes.append({
            "label": lbl, "in": pin, "out": pout, "type": default_type,
            "dia": int(round(_f(gp.get("dia", 50)))),
            "length": _f(gp.get("length", 0)),
            "elev": _f(gp.get("elev", 0)),
            "c": gp.get("c", 120),
            "status": "", "group": "",
        })
    combined.pipes = kept_pipes

    # ── 삭제된 노드를 참조하는 부속(nozzle/fitting/pump/valve) 정리 ──
    def _nz_in(nz):
        return str(nz.get("in") or nz.get("input_node") or nz.get("input") or "")
    combined.nozzles = [nz for nz in combined.nozzles if _nz_in(nz) in valid]
    if getattr(combined, "fittings", None):
        combined.fittings = [f for f in combined.fittings
                             if str(f.get("in", "")) in valid and str(f.get("out", "")) in valid]
    combined.pumps = [pm for pm in combined.pumps
                      if str(pm.get("in", "")) in valid and str(pm.get("out", "")) in valid]
    combined.valves = [vv for vv in combined.valves
                       if str(vv.get("in", "")) in valid and str(vv.get("out", "")) in valid]


def register(app, *, COMBINED_OUTPUT_DIR, OVERALL_OUTPUT_DIR, PROTOTYPE_OUTPUT_DIR, _COMBINED_JOBS, _COMBINED_JOBS_CAP, _PROTOTYPE_JOBS, _RISER_HEIGHT_FRAC, _RISER_SCHEMATIC_SPAN_MM, _common_network_to_geometry, _err500, _serve_run_file, _sweep_old_run_dirs, _to_float):

    def _remap_riser_to_head_av(system_riser: dict, head_av_node: dict, av_label: str,
                                head_nodes: list[dict] | None = None):
        """계통도 라이저를 실좌표 정규화 → 헤드망 AV 기준으로 배치 (RiserTables).

        계통도 픽 좌표(수십만 mm)와 헤드망 좌표(평면 DXF, 수만 mm)는 도메인이 달라 emit_sdf
        정규화 시 라이저가 한쪽에 압축된다. 하드코딩 답안(28F) 좌표를 차용하던 방식을 버리고,
        라이저 **자체의 상대 형상(층별 노드 전부 포함)** 을 유지한 채 헤드망 크기에 맞춰 균일
        스케일하고, AV 노드를 헤드망 AV 위치에 정합시켜 라이저를 AV 위쪽에 배치한다.
        → 층 단위 노드가 몇 개든, 어느 건물이든 일반화 (28F 전용 하드코딩 제거).
        좌표가 비숫자/누락이면 (KeyError, TypeError, ValueError) 를 올린다(호출자가 400 처리).
        """
        from remote30_full_network import RiserTables
        nodes = system_riser["nodes"]
        if not nodes:
            raise ValueError("라이저 노드가 비어 있음")
        head_av_x = float(head_av_node["x"])
        head_av_y = float(head_av_node["y"])

        # 라이저 native 좌표에서 AV 위치 + 자체 bbox 스팬 산출.
        src_av = next((n for n in nodes if str(n.get("label", "")) == str(av_label)), None)
        if src_av is None:
            src_av = nodes[-1]   # AV 라벨 부재 시 마지막 노드를 AV 로 간주(폴백)
        src_av_x = float(src_av["x"])
        src_av_y = float(src_av["y"])
        xs = [float(n["x"]) for n in nodes]
        ys = [float(n["y"]) for n in nodes]
        riser_span = max(max(xs) - min(xs), max(ys) - min(ys), 1.0)

        # 헤드망 특성 크기(bbox 대각선) — 라이저를 이 스케일의 일정 비율로 그려 joint
        # 정규화 시 라이저·헤드망이 서로 압축되지 않게 한다. 없으면 폴백 스팬.
        head_char = _RISER_SCHEMATIC_SPAN_MM
        if head_nodes:
            hxs = [float(n["x"]) for n in head_nodes if n.get("x") is not None]
            hys = [float(n["y"]) for n in head_nodes if n.get("y") is not None]
            if hxs and hys:
                hd = math.hypot(max(hxs) - min(hxs), max(hys) - min(hys))
                if hd > 1.0:
                    head_char = hd
        scale = (head_char * _RISER_HEIGHT_FRAC) / riser_span

        remapped_nodes: list[dict] = []
        for n in nodes:
            new_n = dict(n)
            new_n["x"] = int(round(head_av_x + (float(n["x"]) - src_av_x) * scale))
            new_n["y"] = int(round(head_av_y + (float(n["y"]) - src_av_y) * scale))
            remapped_nodes.append(new_n)
        return RiserTables(
            nodes=remapped_nodes,
            pipes=list(system_riser["pipes"]),
            pumps=list(system_riser.get("pumps", [])),
            valves=list(system_riser.get("valves", [])),
            av_node_label=av_label,
        )

    @app.post("/api/remote30/combined/build")
    def remote30_combined_build():
        """평면도 헤드망 + 계통도 라이저 → 결합 SDF 생성.

        Body (JSON):
            plane_job_id   : Remote 30 프로토타입 평면도 모드의 job_id
            plane_edit     : { added_heads:[[x,y],...], deleted_indices:[int,...],
                               zones:[[x1,y1,x2,y2],...], alarm_x, alarm_y }
            system_riser   : extract_riser_msp_28f 의 출력 그대로 (nodes/pipes/pumps/valves/av_node_label)

        Returns:
            { ok, job_id, sdf, nodes, pipes, pumps, valves, nozzles, download_url }
        """
        import secrets
        body = request.get_json(silent=True) or {}
        plane_job_id = (body.get("plane_job_id") or "").strip()
        if not plane_job_id:
            return jsonify({"ok": False, "message": "plane_job_id 가 필요합니다 (평면도 추출 먼저)"}), 400
        plane_job = _PROTOTYPE_JOBS.get(plane_job_id)
        if not plane_job:
            return jsonify({"ok": False, "message": f"unknown plane_job_id {plane_job_id}"}), 404
        if "detected_heads" not in plane_job:
            return jsonify({"ok": False, "message": "평면도 Stage A (run/stream) 가 아직 완료되지 않았습니다."}), 400

        system_riser = body.get("system_riser")
        if not system_riser or not system_riser.get("nodes") or not system_riser.get("pipes"):
            return jsonify({"ok": False, "message": "system_riser (계통도 추출) 가 필요합니다"}), 400

        plane_edit = body.get("plane_edit") or {}
        added = [tuple(p) for p in plane_edit.get("added_heads", [])]
        deleted = set(int(i) for i in plane_edit.get("deleted_indices", []))
        zones = [tuple(z) for z in plane_edit.get("zones", [])]
        alarm_xy = plane_job.get("alarm_xy")
        ax, ay = plane_edit.get("alarm_x"), plane_edit.get("alarm_y")
        if ax is not None and ay is not None:
            try:
                alarm_xy = (float(ax), float(ay))
            except (TypeError, ValueError):
                pass

        from remote30_prototype import select_worst30_heads, build_input_tables
        from remote30_full_network import (
            stitch_riser_and_heads, emit_full_sdf,
            prepend_machine_room_to_riser, insert_source_pump,
            normalize_pipe_bores, size_pipes_by_velocity, annotate_pipe_velocity,
        )

        # ── 가압 방식 — "gravity"(자연낙차/고가수조, 기본) | "pump"(펌프 가압).
        # 펌프 가압이면 (1) 기계실/수원을 망 최하부로 배치·재고도(고저차 lift 반영),
        # (2) stitch 후 수원(Input) 직후에 펌프 요소를 삽입한다.
        pressurization = str(body.get("pressurization") or "gravity").strip().lower()
        pump_spec = body.get("pump") or {}
        is_pump = pressurization == "pump"
        # 등각 세트의 고도 펼침 배율 — 평면/등각 두 세트를 항상 함께 emit 하므로,
        # 등각 좌표 베이크(_bake_isometric_node_coords)의 lift 강도만 받는다.
        has_iso_z_scale = _to_float(body.get("has_iso_z_scale"), 1.0)
        # KFP 표시좌표 배율 — K-Fire Solver 에서 노드가 작/크게 보일 때 조정(기본 1.0).
        # 표시 전용(length_m·elevation_m 불변)이라 유압계산 결과는 동일. KFP 에만 적용.
        kfp_coord_scale = min(max(_to_float(body.get("kfp_coord_scale"), 1.0), 0.05), 20.0)
        # 수원(기계실)이 최저헤드보다 몇 m 아래인지 — 펌프 흡입측 실양정(>0). DXF 에
        # z 가 없어 도출 불가 → 사용자 입력(미지정 0). 0 이면 고저차 lift 없음.
        source_drop_m = abs(_to_float(pump_spec.get("source_drop_m"), 0.0))
        # 헤드 설치방향(전역) — 상향식(upright)=가지배관 위로 돌출, 하향식(pendent)=아래로.
        # DXF 에 상/하향 정보가 없어(재질과 동일) 전역 토글로 받는다. 표시 전용 — 헤드를
        # 짧은 니플로 ±z 띄워 그릴 뿐, 수리 elevation·length 는 불변. head_z_frac 은 도면
        # 대각선 대비 비율(스케일 무관) — m/mm 좌표계 모두에서 일관되게 보이도록.
        head_orientation = str(body.get("head_orientation") or "pendent").strip().lower()
        if head_orientation not in ("upright", "pendent"):
            head_orientation = "pendent"
        head_z_frac = _to_float(body.get("head_z_frac"), 0.04)
        if head_z_frac < 0:
            head_z_frac = 0.0
        # 불리한 헤드 개수 N — 평면도에서 고른 값(통합 빌드 body 의 n_heads, 없으면
        # finalize 단계에서 저장된 plane_job["k_heads"]). 미지정이면 select_worst30_heads
        # 기본(30)을 따른다. 표시 전용이 아니라 망에 포함될 헤드 수를 결정 → 수리계산에도 반영.
        n_heads_raw = body.get("n_heads", plane_job.get("k_heads"))
        k_heads: int | None = None
        if n_heads_raw is not None:
            try:
                _kv = int(n_heads_raw)
                if _kv >= 1:
                    k_heads = _kv
            except (TypeError, ValueError):
                pass

        # ── Stage A 마무리 — 평면도 헤드 선정 + PipeTables 생성
        detected_pos = [tuple(d["pos"]) for d in plane_job.get("detected_heads", [])]
        manual_heads = [p for i, p in enumerate(detected_pos) if i not in deleted]
        manual_heads.extend(added)

        try:
            selection = select_worst30_heads(
                pipe_entities=plane_job.get("pipe_ents", []),
                layer_categories=plane_job.get("layer_cat", {}),
                manual_source=alarm_xy,
                manual_heads=manual_heads if (manual_heads or deleted or added) else None,
                zones=zones if zones else None,
                **({"k": k_heads} if k_heads is not None else {}),
            )
            head_tables = build_input_tables(
                selection,
                pipe_entities=plane_job.get("pipe_ents", []),
                project_title=Path(plane_job["dxf_path"]).stem,
            )
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

        # ── 계통도 라이저 → RiserTables.
        # ★ 실좌표 정규화: system_riser 의 노드 좌표(사용자 계통도 픽 — 수십만 mm) 와 헤드망 노드
        # 좌표(평면도 DXF — 수만 mm)가 도메인이 달라 emit_sdf 의 정규화 시 라이저가 한쪽에 압축됨.
        # 라이저 자체 형상(층별 노드 포함)을 헤드망 크기에 맞춰 균일 스케일 + 헤드망 AV 위치로
        # translate → 라이저가 헤드망 AV 위쪽에 자연스럽게 배치, 좌표 단위 일치 (28F 하드코딩 제거).
        av_label = str(system_riser.get("av_node_label", "10"))
        head_av_node = next((n for n in head_tables.nodes if n["label"] == av_label), None)
        if head_av_node is None:
            return jsonify({"ok": False,
                            "message": f"헤드망에 AV(label={av_label}) 노드가 없음 — 평면도 추출 다시 확인"}), 500
        # 좌표 정규화 — head_av/노드 좌표가 비숫자·누락이면 500+traceback 대신 깔끔한 400.
        try:
            riser = _remap_riser_to_head_av(system_riser, head_av_node, av_label,
                                            head_nodes=head_tables.nodes)
        except (KeyError, TypeError, ValueError) as exc:
            return jsonify({"ok": False,
                            "message": f"계통도/헤드망 노드 좌표가 올바르지 않습니다: {exc}"}), 400

        # ── 기계실(옥상수조) 경로 (선택) → 라이저 Input 앞에 prepend.
        # 있으면 수원 경계가 탱크(m1)로 이동하고 옥상부 마찰손실이 통합망에 반영됨.
        machine_room = body.get("machine_room")
        mr_attached = False
        machine_room_labels: list[str] = []
        pump_junction_label: str | None = None
        if machine_room and machine_room.get("nodes") and machine_room.get("pipes"):
            # 펌프 junction = 기계실이 병합되는 라이저 Input ("1"). prepend 후엔
            # io 가 No 로 강등되므로 prepend 전에 미리 라벨을 기록해 둔다. 캔버스에서
            # 이 노드를 "펌프" 로 명시해 기계실↔계통도 경계를 시각적으로 분리.
            _ri = next((n for n in riser.nodes
                        if str(n.get("io_node", "")).lower() == "input"), None)
            if _ri is None:
                _ri = next((n for n in riser.nodes if str(n["label"]) == "1"), None)
            pump_junction_label = str(_ri["label"]) if _ri else None
            # 기계실 노드 라벨 (conn=mK 은 라이저 Input 과 병합돼 사라지므로 제외)
            _conn = str(machine_room.get("conn_node_label")
                        or machine_room["nodes"][-1]["label"])
            machine_room_labels = [str(n["label"]) for n in machine_room["nodes"]
                                   if str(n["label"]) != _conn]
            try:
                riser, mr_attached = prepend_machine_room_to_riser(
                    machine_room, riser,
                    at_bottom=is_pump, source_drop_below_lowest_m=source_drop_m)
            except Exception as _mr_exc:  # noqa: BLE001 — 기계실 실패가 통합을 막지 않도록
                warnings.warn(f"[combined] 기계실 prepend 실패 (라이저만 사용): {_mr_exc}",
                               RuntimeWarning, stacklevel=2)
            if not mr_attached:
                machine_room_labels = []
                pump_junction_label = None

        # ── Stitch + emit
        try:
            combined = stitch_riser_and_heads(
                riser, head_tables,
                machine_room_labels=machine_room_labels,
                pump_junction_label=pump_junction_label,
                machine_room_plan_edges=(machine_room.get("plan_edges") if mr_attached else None),
                machine_room_at_bottom=is_pump,
            )
            # ── 가압 방식: 펌프 선택 시 수원 경계에 펌프 삽입 (자연낙차는 기본값, 무변경)
            if is_pump:
                rated_q = float(pump_spec.get("rated_q") or 2400)
                rated_h = float(pump_spec.get("rated_h") or 100)
                count = int(pump_spec.get("count") or 1)
                insert_source_pump(combined, rated_q_lpm=rated_q, rated_h_m=rated_h, count=count)
            # ── 내경 정규화: 상류(입상관)→하류(가지) 단조 비증가로 꼬임 해소 + 전 구간 한 치수 승급.
            #   build 시 1회만 승급(bump_one_size=True). rebuild 는 승급 없이 detangle 만(멱등).
            try:
                _bore_ch = normalize_pipe_bores(
                    combined.nodes, combined.pipes, bump_one_size=True)
                app.logger.info("combined/build: pipe bores normalized (+1 size), changed=%d", _bore_ch)
            except Exception as _e:
                app.logger.warning("combined/build: bore normalize skipped: %s", _e)
            # ── 유속 상한(≤50A 6 m/s, ≥65A 10 m/s) 보장: 과토출 대비 safety 여유를 두고
            #   유량 기준 최소 내경으로 승급(never-shrink). 정규화 뒤에 두어 승급분을 보존.
            try:
                _vel = size_pipes_by_velocity(
                    combined.nodes, combined.pipes, combined.nozzles,
                    safety=1.2, keep_existing=True)
                app.logger.info(
                    "combined/build: velocity sizing changed=%d, max v %.2f->%.2f, viol %d->%d",
                    _vel["changed"], _vel["max_velocity_before"], _vel["max_velocity_after"],
                    _vel["violations_before"], _vel["violations_after"])
            except Exception as _e:
                app.logger.warning("combined/build: velocity sizing skipped: %s", _e)
            # ── 통합 뷰 계산(유속) 오버레이용 배관 stamp — 최종 내경 기준.
            try:
                _va = annotate_pipe_velocity(combined.nodes, combined.pipes, combined.nozzles)
                app.logger.info("combined/build: velocity annotated, max %.2f m/s, over=%d",
                                _va["max_velocity"], _va["violations"])
            except Exception as _e:
                app.logger.warning("combined/build: velocity annotate skipped: %s", _e)
            job_id = secrets.token_hex(6)
            _sweep_old_run_dirs(PROTOTYPE_OUTPUT_DIR, OVERALL_OUTPUT_DIR, COMBINED_OUTPUT_DIR)
            out_dir = COMBINED_OUTPUT_DIR / job_id
            out_dir.mkdir(parents=True, exist_ok=True)
            # 통합 Title — 업로드한 평면도 파일명(건물/도면명)을 따른다. 답안지 Title
            # 컨벤션이 건물명(예: "Officetell")이라, 내부 식별자(SYSTEM_EXTRACT_V1)나
            # 도구 브랜딩이 그대로 노출되지 않도록 도면 stem 을 쓴다.
            title = (Path(plane_job.get("dxf_path", "")).stem
                     or system_riser.get("title")
                     or "Combined")

            import zipfile as _zipfile
            import copy as _copy
            from remote30_prototype import emit_kfp as _emit_kfp, emit_has as _emit_has

            # ── z-aware 도구(K-solver/HASS)용 라이저 "참 3D 축정렬" 좌표 — KFP/HAS 만 적용.
            # 라이저 막대는 schematic 으로 y 가 인위적으로 펼쳐져 있다(_layout_riser_as_schematic).
            # SDF/PIPENET 은 선언 length 와 자체 schematic 을 써 이 y-spread 가 문제없지만,
            # z(고도)까지 쓰는 K-solver/HASS 에서는 (1) 라이저가 대각선 지그재그로 깨지고,
            # (2) KFP length 가 3D 좌표거리(=√(Δx²+Δy²+Δz²))라 인위적 Δy 가 라이저 배관장을
            # 부풀린다.
            #
            # 단순히 라이저 전체를 AV 한 점으로 모으면 수직 입상관은 맞지만 옥상 수평 헤더
            # (등고선 z 가 같고 수평으로 뻗는 구간)가 한 점에 뭉개져 길이 0 배관이 생기고
            # 계통도가 찌그러진다. → 라이저를 하나의 수직 평면(y=AV.y 고정) 안에 축정렬로
            # 재구성한다: 각 배관의 실제 선언 length 와 고도차 rise 로 수평 run =
            # √(length²−rise²) 을 구하고, z 는 노드 고도(권위값)를 그대로 둔다. 수직관은
            # rise≈length → 수평 run≈0(기둥), 수평 헤더는 rise≈0 → run=실제 길이(수평선).
            # 방향(±x)은 실제 계통도 DXF x 차의 부호를 따라 자연스러운 배치를 유지한다.
            # 기계실/헤드망(실제 평면 좌표)은 보존.
            _mr_set = {str(l) for l in (machine_room_labels or [])}
            riser_collapse_labels = {
                str(n["label"]) for n in riser.nodes if str(n["label"]) not in _mr_set
            }
            # 라이저 고도 z(m) — riser.nodes 의 원좌표에서(_auto_spread 판정용).
            _riser_elev = {
                str(n["label"]): float(n.get("elevation", 0.0)) for n in riser.nodes
            }
            # 라이저 인접리스트 (라벨→[(이웃, 선언length_m), ...]) — riser.pipes 기준,
            # collapse 대상(=계통도) 노드 사이 간선만. 기계실 간선은 제외.
            _riser_adj: dict[str, list[tuple[str, float]]] = {}
            for _p in riser.pipes:
                _a = str(_p.get("in", "")); _b = str(_p.get("out", ""))
                if _a not in riser_collapse_labels or _b not in riser_collapse_labels:
                    continue
                try:
                    _ln = float(_p.get("length", 0.0) or 0.0)
                except (TypeError, ValueError):
                    _ln = 0.0
                _riser_adj.setdefault(_a, []).append((_b, _ln))
                _riser_adj.setdefault(_b, []).append((_a, _ln))

            # 라이저 고도(elevation) 변화폭 — 미리보기 autoRiserSpread 와 동일 판정.
            # < 1.0 m 이면 "실제 고도 정보 없음(단층 도면·하드코드 elev)" → 토폴로지로 강제 펼침.
            _riser_elev_vals = [v for k, v in _riser_elev.items()
                                if k in riser_collapse_labels]
            _riser_elev_spread = ((max(_riser_elev_vals) - min(_riser_elev_vals))
                                  if _riser_elev_vals else 0.0)
            _auto_spread = _riser_elev_spread < 1.0

            # ── 평면 세트/미리보기 geometry 는 실 DXF 좌표를 쓴다(교차·주배관은 도면 그대로,
            #    가지배관만 build_input_tables 에서 이미 직각 스냅됨). tree-packing 스키매틱
            #    재배치는 평면을 도면과 동떨어진 기괴한 형태로 만들어 폐지 — 대신 iso(등각)
            #    세트에만 tree-packing 을 적용해 아이소 꼬임을 방지한다(아래 combined_iso).
            #    수리값(length·elevation)은 어느 경로든 불변.

            # ── 헤드(스프링클러) z 돌출 — 상향식(+)/하향식(−)을 표시 전용 display_z 로 베이크.
            #   헤드 = nozzle 부착(input) 노드. 돌출량은 평면 대각선 비율(head_z_frac)이라
            #   좌표 단위(m/mm)에 무관. elevation(수리 실표고)은 일절 건드리지 않아 결과 불변.
            #   여기서 한 번 계산해 KFP/HAS(display_z)·iso SDF(iso 베이크)·geometry 가 공유.
            _hd_xs = [float(n.get("x", 0) or 0) for n in combined.nodes]
            _hd_ys = [float(n.get("y", 0) or 0) for n in combined.nodes]
            _plan_diag = (math.hypot(max(_hd_xs) - min(_hd_xs), max(_hd_ys) - min(_hd_ys))
                          if _hd_xs else 0.0)
            head_label_set = {
                str(nz.get("in") or nz.get("input_node") or nz.get("input") or "")
                for nz in combined.nozzles
                if (nz.get("in") or nz.get("input_node") or nz.get("input"))
            }
            head_label_set.discard("")
            _head_sign = 1.0 if head_orientation == "upright" else -1.0
            head_disp_z = _head_sign * _plan_diag * head_z_frac

            def _collapse_riser_to_column(net_obj):
                """net_obj 사본에서 라이저를 수직 입상관으로 재배치 (미리보기 3D 뷰와 정합).

                ★ 단위 일치 핵심: 라이저 x,y 를 AV 한 점으로 모으고(미리보기 riserSet→avX,avY
                와 동일), 모든 노드의 표시 z 를 **평면 좌표(mm) 단위인 display_z** 로 베이크한다.
                display_z 는 emit_sdf 가 x,y 와 동일 _scale 로 정규화 → KFP/HAS 변환기가 x,y 와
                같은 배율을 타 평면과 자동 비례한다. (display_z 를 안 박으면 변환기가 Position z
                부재로 raw elevation[미터]으로 fallback 하는데, x,y 는 mm·z 는 m 라 단위가
                1000× 어긋나 라이저 기둥이 평면 대비 폭주했다 — 이 버그 수정.)

                기둥 내부 고도 분배 t∈[0,1] (AV=0 바닥 → 최상류=1 꼭대기) 만 두 방식:
                (A) 실측 고도폭 < 1 m (단층/하드코드) → 토폴로지 BFS 순서로 균등 분배.
                (B) 실측 고도폭 ≥ 1 m → elevation 을 [0,1] 정규화(실 층고 비율 보존).
                어느 쪽이든 기둥 전체 높이 = 0.5 × 평면 대각선 으로 통일해 항상 평면과 비례.
                elevation(수리 실표고)은 일절 변경하지 않아 head 계산 불변.

                반환: (사본/원본, 변경여부). AV 를 못 찾으면 원본·False.
                """
                if not riser_collapse_labels:
                    return net_obj, False
                av = next((n for n in net_obj.nodes
                           if str(n.get("label")) == str(av_label)), None)
                if av is None or av.get("x") is None or av.get("y") is None:
                    return net_obj, False
                cx = float(av["x"])
                cy = int(round(float(av["y"])))
                import math as _math
                z_net = _copy.deepcopy(net_obj)

                z_scale = 3.0  # 미리보기 opts.zScale 기본값 — 기둥 높이 배율(평면 대비).
                dir_sign = -1.0 if is_pump else 1.0
                # ── 공통 1: 라이저 x,y 를 AV 한 점으로 모은다(순수 수직 기둥) ──
                #   미리보기 riserSet→(avX,avY)(remote30_prototype.html) 와 동일. 옛 branch-B
                #   의 수평 run 재구성(라이저 꼬임 유발)을 폐지해 미리보기와 규격 일치.
                for n in z_net.nodes:
                    if str(n.get("label")) in riser_collapse_labels:
                        n["x"] = int(round(cx))
                        n["y"] = cy
                # ── 공통 1.5: 헤드(nozzle 부착 leaf)의 x,y 를 가지배관 이웃 노드에 스냅.
                #   헤드는 display_z 로만 ±수직 돌출하는데, 평면 x,y 가 이웃(가지 tee)과
                #   어긋나 있으면 드롭 배관이 대각선으로 보인다(예: head N48 35°, N54 48°).
                #   이웃 x,y 로 맞추면 Δxy=0 → 완전 90° 수직 드롭. leaf(이웃 1개)만 대상으로
                #   해 가지 중간 통과 헤드는 건드리지 않는다. 표시 전용(elevation·length_m
                #   불변 → 수리 결과 동일).
                _znode = {str(n.get("label")): n for n in z_net.nodes}
                _head_nbrs: dict[str, set] = {}
                for _p in z_net.pipes:
                    _a = str(_p.get("in", "")); _b = str(_p.get("out", ""))
                    if _a in head_label_set and _b:
                        _head_nbrs.setdefault(_a, set()).add(_b)
                    if _b in head_label_set and _a:
                        _head_nbrs.setdefault(_b, set()).add(_a)
                for _hl, _nbs in _head_nbrs.items():
                    if len(_nbs) != 1:  # leaf 헤드만 (통과 헤드 제외)
                        continue
                    _hn = _znode.get(_hl)
                    _tn = _znode.get(next(iter(_nbs)))
                    if (_hn is None or _tn is None
                            or _tn.get("x") is None or _tn.get("y") is None):
                        continue
                    _hn["x"] = _tn["x"]
                    _hn["y"] = _tn["y"]
                # ── 공통 2: 기둥 높이 = 0.5 × 전체 평면 대각선 × zScale (x,y 가 AV 로 모인 뒤
                #   bbox — 미리보기 spreadHeight/_spreadH 와 정합). display_z 는 emit_sdf 가 x,y
                #   와 동일 _scale 로 정규화 → 평면과 자동 비례(3/longest 보정 불필요).
                _full_xs = [float(n["x"]) for n in z_net.nodes if n.get("x") is not None]
                _full_ys = [float(n["y"]) for n in z_net.nodes if n.get("y") is not None]
                _x_span = (max(_full_xs) - min(_full_xs)) if _full_xs else 0.0
                _y_span = (max(_full_ys) - min(_full_ys)) if _full_ys else 0.0
                _full_diag = _math.hypot(_x_span, _y_span)
                spread_h = (0.5 * _full_diag * z_scale) if _full_diag > 1e-9 else 1500.0

                if _auto_spread:
                    # (A) 실측 고도폭 < 1 m(단층·하드코드 elev) → 토폴로지 BFS 순서로 펼친다.
                    #   root(수원/Input)=꼭대기(spread_h) → AV(헤드평면 anchor)=바닥(0).
                    import collections as _collections
                    root = next((str(n["label"]) for n in z_net.nodes
                                 if str(n.get("label")) in riser_collapse_labels
                                 and n.get("io_node") == "Input"), None)
                    if root is None:  # Input 없으면 AV 에서 가장 먼 노드(펌프 후보)를 root 로.
                        _vis = {str(av_label): 0}
                        _q = _collections.deque([str(av_label)])
                        _far, _fd = str(av_label), 0
                        while _q:
                            u = _q.popleft()
                            for v, _ln in _riser_adj.get(u, ()):
                                if v in _vis:
                                    continue
                                _vis[v] = _vis[u] + 1
                                if _vis[v] > _fd:
                                    _fd, _far = _vis[v], v
                                _q.append(v)
                        root = _far
                    order: list[str] = []
                    seen = {root}
                    q = _collections.deque([root])
                    while q:
                        u = q.popleft()
                        order.append(u)
                        for v, _ln in _riser_adj.get(u, ()):
                            if v in seen:
                                continue
                            seen.add(v)
                            q.append(v)
                    for lbl in riser_collapse_labels:  # 다른 component 는 끝에 append
                        if lbl not in seen:
                            order.append(lbl)
                    _av_s = str(av_label)  # AV 를 강제로 마지막(헤드평면 anchor=바닥)으로.
                    if _av_s in order and order[-1] != _av_s:
                        order.remove(_av_s)
                        order.append(_av_s)
                    _n_order = max(1, len(order) - 1)
                    _riser_z = {
                        lbl: dir_sign * (1.0 - i / _n_order) * spread_h
                        for i, lbl in enumerate(order)
                    }
                    _mr_z = dir_sign * spread_h * 1.18  # 라이저 극단 너머 수원 평면(미리보기 _mrZ).
                    for n in z_net.nodes:
                        lbl = str(n.get("label"))
                        if lbl in _mr_set:
                            n["display_z"] = _mr_z
                        elif lbl in riser_collapse_labels:
                            n["display_z"] = _riser_z.get(lbl, 0.0)
                        elif lbl in head_label_set:
                            n["display_z"] = head_disp_z
                        else:
                            n["display_z"] = 0.0
                    return z_net, True

                # (B) 실측 고도폭 ≥ 1 m → elevation(m)을 표시 z(mm)로 펼친다(실 층고 비율 보존).
                #   미리보기 z=(e-eMid)*1000*zScale 와 동일: ×1000=m→mm 로 평면 x,y(mm) 단위와
                #   일치, eMid 중심. 라이저·헤드·평면 모두 같은 식. 기계실은 라이저 극단 너머
                #   수원 평면(_mrZ)에. (옛 코드는 display_z 미베이크 → 변환기가 raw elevation[m]
                #   으로 fallback, x,y[mm] 와 1000× 어긋나 기둥이 폭주했다 — 이 버그 수정.)
                _all_e = [float(n.get("elevation", 0.0) or 0.0) for n in z_net.nodes]
                _e_lo, _e_hi = (min(_all_e), max(_all_e)) if _all_e else (0.0, 0.0)
                _e_mid = (_e_lo + _e_hi) / 2.0
                _riser_extreme = (((_e_lo - _e_mid) if is_pump else (_e_hi - _e_mid))
                                  * 1000.0 * z_scale * dir_sign)
                _mr_z = _riser_extreme + dir_sign * spread_h * 0.18
                for n in z_net.nodes:
                    lbl = str(n.get("label"))
                    _ez = (float(n.get("elevation", 0.0) or 0.0) - _e_mid) * 1000.0 * z_scale
                    if lbl in _mr_set:
                        n["display_z"] = _mr_z
                    elif lbl in head_label_set:
                        n["display_z"] = _ez + head_disp_z
                    else:
                        n["display_z"] = _ez
                return z_net, True

            def _emit_bundle(net_obj, suffix: str) -> dict:
                """net_obj → SDF(+SLF)/KFP/HAS/ZIP 한 세트 생성. suffix 로 평면("")/등각("_iso") 구분.

                net_obj 의 노드 좌표를 그대로 베이크하므로, 등각 세트는 호출 전에
                _bake_isometric_node_coords 로 (x,y) 를 등각투영해 넘긴다. HAS 는 좌표가
                이미 정해져 있으니 isometric=False (이중 투영 방지).
                """
                b_sdf = out_dir / f"combined_{job_id}{suffix}.sdf"
                emit_full_sdf(net_obj, b_sdf, project_title=title)
                b_slf = out_dir / f"combined_{job_id}{suffix}.slf"
                # ── KFP 와 HAS 는 표시 규약이 다르다(둘 다 등각 베이크 안 된 원본 combined
                #    에서 파생, 라이저만 수직 기둥으로 collapse — z_net):
                #  · KFP: 참 3D 직교좌표 [x,y,z]. 노드 z = display_z(라이저=기둥, 헤드평면=0).
                #    K-Fire Solver 가 화면에서 자체 등각투영하므로 우리는 베이크하지 않는다.
                #  · HAS: HASS 는 InsertionPoint 2D 좌표를 **그대로** 표시(재투영 안 함, 참조
                #    계통도도 30° 베이크본). 따라서 emit_has(isometric=True) 로 display_z 를
                #    화면 Y 에 lift 베이크해야 라이저가 기둥으로 보인다. Height(수리표고)는
                #    elevation_m 분리 보존. (SDF 는 PIPENET 2D 스키매틱 — net_obj 좌표 그대로.)
                z_sdf = b_sdf
                z_net, _z_done = _collapse_riser_to_column(combined)
                if _z_done:
                    z_sdf = out_dir / f"combined_{job_id}{suffix}_z.sdf"
                    try:
                        emit_full_sdf(z_net, z_sdf, project_title=title)
                    except Exception as _z_exc:  # noqa: BLE001 — 사본 실패 시 원본 좌표로 폴백
                        warnings.warn(f"[combined{suffix}] z-aware SDF emit 실패 (원본 좌표 사용): {_z_exc}", RuntimeWarning, stacklevel=2)
                        z_sdf = b_sdf
                b_kfp = out_dir / f"combined_{job_id}{suffix}.kfp"
                b_kfp_ok = False
                try:
                    # 통합망 KFP — 미리보기와 동일한 스키매틱 표시좌표(라이저=기둥,
                    # display_z)로 비율 일치. length_m·elevation_m 는 실값 보존(수리 권위값).
                    _emit_kfp(z_sdf, b_kfp, coord_scale=kfp_coord_scale,
                              display_geometry=True)
                    b_kfp_ok = b_kfp.is_file()
                except Exception as _kfp_exc:  # noqa: BLE001 — KFP 실패가 통합 출력을 막지 않도록
                    warnings.warn(f"[combined{suffix}] KFP emit 실패 (SDF 는 정상): {_kfp_exc}", RuntimeWarning, stacklevel=2)
                b_has = out_dir / f"combined_{job_id}{suffix}.has"
                b_has_ok = False
                try:
                    _emit_has(z_sdf, b_has, isometric=True, iso_z_scale=has_iso_z_scale)
                    b_has_ok = b_has.is_file()
                except Exception as _has_exc:  # noqa: BLE001 — HAS 실패가 통합 출력을 막지 않도록
                    warnings.warn(f"[combined{suffix}] HAS emit 실패 (SDF 는 정상): {_has_exc}", RuntimeWarning, stacklevel=2)
                # 임시 z-aware SDF/SLF 정리 — ZIP·다운로드에는 원본 b_sdf 만 포함.
                if z_sdf != b_sdf:
                    for _tmp in (z_sdf, z_sdf.with_suffix(".slf")):
                        try:
                            if _tmp.is_file():
                                _tmp.unlink()
                        except OSError:
                            pass
                # 전체 ZIP — SDF + SLF + KFP + HAS 를 한 번에 (모든 포맷 묶음).
                b_zip = out_dir / f"combined_{job_id}{suffix}.zip"
                with _zipfile.ZipFile(b_zip, "w", _zipfile.ZIP_DEFLATED) as zf:
                    zf.write(b_sdf, arcname=b_sdf.name)
                    if b_slf.is_file():
                        zf.write(b_slf, arcname=b_slf.name)
                    if b_kfp_ok:
                        zf.write(b_kfp, arcname=b_kfp.name)
                    if b_has_ok:
                        zf.write(b_has, arcname=b_has.name)
                # PIPENET-native ZIP — .sdf + .slf 만. PIPENET 은 두 파일이 같은 폴더에
                # 있어야 호칭경↔내경 lookup 이 되므로 SDF 버튼은 이 쌍을 묶어 내보낸다
                # (KFP/HAS 가 딸려나오지 않도록). .xml 결과파일은 PIPENET 이 연산 후 생성하는
                # 산출물이라 입력 번들에 포함하지 않는다.
                b_zip_sdf = out_dir / f"combined_{job_id}{suffix}_pipenet.zip"
                with _zipfile.ZipFile(b_zip_sdf, "w", _zipfile.ZIP_DEFLATED) as zf:
                    zf.write(b_sdf, arcname=b_sdf.name)
                    if b_slf.is_file():
                        zf.write(b_slf, arcname=b_slf.name)
                return {"sdf": b_sdf, "slf": b_slf, "kfp": b_kfp, "has": b_has, "zip": b_zip,
                        "zip_sdf": b_zip_sdf, "kfp_ok": b_kfp_ok, "has_ok": b_has_ok}

            # 평면 세트(실 DXF 좌표) — 캔버스 geometry 와 동일한 평면도 좌표(가지만 직각화).
            plan_bundle = _emit_bundle(combined, "")
            # 등각 세트 — 사본에 tree-packing 스키매틱 재배치 후 30° 등각투영 베이크.
            # 등각은 헤드평면을 균일 격자로 재배치해야 아이소 꼬임이 안 생기므로(평면과 달리)
            # 여기서만 정돈한다. 표시 전용 변환이라 SDF/KFP/HAS 의 수리계산 결과는 평면과 동일.
            combined_iso = _copy.deepcopy(combined)
            try:
                _tidied = _tidy_head_plane_layout(
                    combined_iso.nodes, combined_iso.pipes, av_label,
                    riser_collapse_labels | _mr_set)
                app.logger.info("combined/build: iso head-plane tidied nodes=%d", _tidied)
            except Exception as _e:  # 정돈 실패는 치명적이지 않음 — 원좌표로 진행.
                app.logger.warning("combined/build: iso head-plane tidy skipped: %s", _e)
            # 라이저·기계실 계통도는 lift 제외 — schematic y 가 이미 수직을 인코딩하므로
            # elevation lift 를 다시 더하면 이중부호로 계통도가 구부러진다. 헤드 z-돌출도
            # 여기선 안 씀(평면 Y 를 기울여 가지배관을 꼬이게 함; 3D·KFP/HAS 는 display_z).
            _bake_isometric_node_coords(combined_iso.nodes, has_iso_z_scale,
                                        no_lift_labels=riser_collapse_labels | _mr_set)
            iso_bundle = _emit_bundle(combined_iso, "_iso")

            out_sdf, out_slf = plan_bundle["sdf"], plan_bundle["slf"]
            out_kfp, out_has, out_zip = plan_bundle["kfp"], plan_bundle["has"], plan_bundle["zip"]
            out_zip_sdf = plan_bundle["zip_sdf"]
            kfp_ok, has_ok = plan_bundle["kfp_ok"], plan_bundle["has_ok"]
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

        # ── 캔버스 시각화용 geometry 데이터 (헤드망 + 라이저 통합)
        riser_labels = {str(n["label"]) for n in riser.nodes}
        geometry = _build_combined_geometry(
            combined, riser, riser_labels, head_label_set,
            machine_room_labels, pump_junction_label,
            is_pump, head_orientation, head_z_frac)

        # ── 포맷 미리보기(round-trip)용 역할 사이드카.
        try:
            roles_sidecar = _build_roles_sidecar(
                combined, riser_labels, head_label_set, riser.av_node_label,
                machine_room_labels, pump_junction_label,
                is_pump, head_orientation, head_z_frac)
            (out_dir / f"combined_{job_id}_roles.json").write_text(
                json.dumps(roles_sidecar, ensure_ascii=False), encoding="utf-8")
        except Exception as _rs_exc:  # noqa: BLE001 — 사이드카 실패가 통합 출력을 막지 않도록
            warnings.warn(f"[combined] roles 사이드카 저장 실패 (미리보기 모양 복원 불가): {_rs_exc}",
                           RuntimeWarning, stacklevel=2)

        # ── 수동 편집 재출력(/combined/rebuild)용 원본 망 캐시. deepcopy 로 보관해
        #    이후 emit 부작용(좌표 베이크 등)이 캐시본을 오염시키지 않도록 격리한다.
        #    z-aware 표시좌표(라이저 기둥 collapse + display_z)도 라벨별로 캐시한다 —
        #    편집 재출력의 KFP/HAS 가 원본과 동일 비율이 되려면 display_z 가 필요하다
        #    (없으면 변환기가 raw elevation[m]/x,y[mm] 단위 1000× 어긋나 라이저 폭주).
        _zaware_map: dict[str, dict] = {}
        try:
            _zp, _zok = _collapse_riser_to_column(combined)
            if _zok:
                for _n in _zp.nodes:
                    _lb = str(_n.get("label"))
                    _zaware_map[_lb] = {
                        "x": _n.get("x"), "y": _n.get("y"),
                        "dz": _n.get("display_z"),
                        "riser": _lb in riser_collapse_labels,
                    }
        except Exception as _zexc:  # noqa: BLE001 — z-aware 캐시 실패는 KFP/HAS 만 영향
            warnings.warn(f"[combined] z-aware 좌표 캐시 실패 (편집 KFP/HAS 비율 저하 가능): {_zexc}",
                           RuntimeWarning, stacklevel=2)
        try:
            if len(_COMBINED_JOBS) >= _COMBINED_JOBS_CAP:
                _COMBINED_JOBS.pop(next(iter(_COMBINED_JOBS)))  # 가장 오래된 항목 제거
            _COMBINED_JOBS[job_id] = {
                "combined": _copy.deepcopy(combined),
                "title": title,
                "zaware": _zaware_map,
                "kfp_coord_scale": kfp_coord_scale,
                "has_iso_z_scale": has_iso_z_scale,
            }
        except Exception as _cache_exc:  # noqa: BLE001 — 캐시 실패가 통합 출력을 막지 않도록
            warnings.warn(f"[combined] rebuild 캐시 저장 실패 (편집 재출력 불가): {_cache_exc}",
                           RuntimeWarning, stacklevel=2)

        return jsonify({
            "ok": True, "job_id": job_id, "sdf": out_sdf.name, "zip": out_zip.name,
            "nodes": len(combined.nodes), "pipes": len(combined.pipes),
            "pumps": len(combined.pumps), "valves": len(combined.valves),
            "nozzles": len(combined.nozzles),
            # 평면 세트
            "download_url_sdf": f"/api/remote30/combined/result/{job_id}/{out_sdf.name}",
            "download_url_slf": f"/api/remote30/combined/result/{job_id}/{out_slf.name}" if out_slf.is_file() else None,
            "download_url_kfp": f"/api/remote30/combined/result/{job_id}/{out_kfp.name}" if kfp_ok else None,
            "download_url_has": f"/api/remote30/combined/result/{job_id}/{out_has.name}" if has_ok else None,
            "download_url_zip": f"/api/remote30/combined/result/{job_id}/{out_zip.name}",
            "download_url_sdf_zip": f"/api/remote30/combined/result/{job_id}/{out_zip_sdf.name}",
            # 등각 세트 (30° 등각투영 좌표 베이크)
            "download_url_sdf_iso": f"/api/remote30/combined/result/{job_id}/{iso_bundle['sdf'].name}",
            "download_url_slf_iso": f"/api/remote30/combined/result/{job_id}/{iso_bundle['slf'].name}" if iso_bundle["slf"].is_file() else None,
            "download_url_kfp_iso": f"/api/remote30/combined/result/{job_id}/{iso_bundle['kfp'].name}" if iso_bundle["kfp_ok"] else None,
            "download_url_has_iso": f"/api/remote30/combined/result/{job_id}/{iso_bundle['has'].name}" if iso_bundle["has_ok"] else None,
            "download_url_zip_iso": f"/api/remote30/combined/result/{job_id}/{iso_bundle['zip'].name}",
            "download_url_sdf_zip_iso": f"/api/remote30/combined/result/{job_id}/{iso_bundle['zip_sdf'].name}",
            "title": title,
            "machine_room_attached": mr_attached,
            "geometry": geometry,
        })

    @app.post("/api/remote30/combined/rebuild")
    def remote30_combined_rebuild():
        """브라우저에서 수동 편집한 통합망 geometry → SDF 재출력.

        Body(JSON): { job_id, geometry:{nodes:[...], pipes:[...]} }
          job_id   : 원본 통합 빌드의 job_id (_COMBINED_JOBS 캐시 키)
          geometry : 편집된 state.combined_geometry (노드/배관만)

        캐시된 원본 CombinedTables 를 deepcopy → geometry 로 패치 → emit_full_sdf 로
        새 SDF 를 생성해 다운로드 URL 을 돌려준다. 패치본을 다시 캐시해 연속 편집을 지원.
        """
        import secrets
        import copy as _copy
        body = request.get_json(silent=True) or {}
        src_job = (body.get("job_id") or "").strip()
        geom = body.get("geometry") or {}
        if not src_job:
            return jsonify({"ok": False, "message": "job_id 가 필요합니다"}), 400
        cache = _COMBINED_JOBS.get(src_job)
        if not cache:
            return jsonify({"ok": False,
                            "message": f"통합 빌드 캐시를 찾을 수 없습니다 (job_id={src_job}). "
                                       "배관망 통합을 다시 실행한 뒤 편집해 주세요."}), 404
        if not geom.get("nodes") or not geom.get("pipes"):
            return jsonify({"ok": False, "message": "편집된 geometry(nodes/pipes)가 필요합니다"}), 400

        from remote30_full_network import (emit_full_sdf, normalize_pipe_bores,
                                            size_pipes_by_velocity, annotate_pipe_velocity)
        try:
            combined = _copy.deepcopy(cache["combined"])
            _patch_combined_from_geometry(combined, geom)
            if not combined.nodes or not combined.pipes:
                return jsonify({"ok": False, "message": "편집 결과 노드/배관이 비어 재출력할 수 없습니다"}), 400
            # ── 내경 꼬임 해소만(detangle) — 편집 후 상류<하류 역전 방지. 승급(+1)은 build 시 1회뿐이라
            #    rebuild 에선 제외(멱등). 사용자 수동 편집을 얇게 줄이지 않고 상류만 ≥ 하류로 끌어올림.
            try:
                normalize_pipe_bores(combined.nodes, combined.pipes, bump_one_size=False)
            except Exception as _e:
                app.logger.warning("combined/rebuild: bore detangle skipped: %s", _e)
            # ── 유속 상한 보장(멱등, never-shrink): 편집으로 얇아진 구간이 유속 초과면 최소 내경으로 승급.
            try:
                _vel = size_pipes_by_velocity(
                    combined.nodes, combined.pipes, combined.nozzles,
                    safety=1.2, keep_existing=True)
                app.logger.info(
                    "combined/rebuild: velocity sizing changed=%d, viol %d->%d",
                    _vel["changed"], _vel["violations_before"], _vel["violations_after"])
            except Exception as _e:
                app.logger.warning("combined/rebuild: velocity sizing skipped: %s", _e)
            try:
                annotate_pipe_velocity(combined.nodes, combined.pipes, combined.nozzles)
            except Exception as _e:
                app.logger.warning("combined/rebuild: velocity annotate skipped: %s", _e)

            new_job = secrets.token_hex(6)
            _sweep_old_run_dirs(PROTOTYPE_OUTPUT_DIR, OVERALL_OUTPUT_DIR, COMBINED_OUTPUT_DIR)
            out_dir = COMBINED_OUTPUT_DIR / new_job
            out_dir.mkdir(parents=True, exist_ok=True)
            title = cache.get("title") or "Combined (edited)"
            out_sdf = out_dir / f"combined_{new_job}_edited.sdf"
            emit_full_sdf(combined, out_sdf, project_title=title)

            # ── KFP/HAS 재출력 — 원본 빌드가 캐시한 z-aware 표시좌표(라이저 기둥 collapse
            #    + display_z)를 라벨별로 재적용해 원본과 동일 비율로 emit 한다. 편집으로
            #    옮긴 노드는 새 x,y 를 유지하되 display_z 는 캐시값(헤드 돌출·고도)을 쓴다.
            #    신규 노드는 display_z 없음 → 헤드평면(0). 캐시 없으면 KFP/HAS 는 생략.
            zaware = cache.get("zaware") or {}
            kfp_scale = cache.get("kfp_coord_scale", 1.0)
            has_zs = cache.get("has_iso_z_scale", 1.0)
            out_kfp = out_dir / f"combined_{new_job}_edited.kfp"
            out_has = out_dir / f"combined_{new_job}_edited.has"
            kfp_ok = has_ok = False
            try:
                from remote30_prototype import emit_kfp as _emit_kfp, emit_has as _emit_has
                # KFP/HAS 원본 SDF 선택 — 라이저 collapse 캐시(zaware)가 있으면 display_z 를
                # 재적용한 z-aware SDF 를, 없으면(라이저 없는 헤드 전용망) 평면 SDF 를 그대로 쓴다.
                z_tmp = None
                if zaware:
                    z_net = _copy.deepcopy(combined)
                    for n in z_net.nodes:
                        zc = zaware.get(str(n.get("label")))
                        if not zc:
                            continue
                        if zc.get("riser"):  # 라이저는 캐시된 기둥 x,y 로 collapse
                            if zc.get("x") is not None:
                                n["x"] = zc["x"]
                            if zc.get("y") is not None:
                                n["y"] = zc["y"]
                        if zc.get("dz") is not None:
                            n["display_z"] = zc["dz"]
                    z_tmp = out_dir / f"combined_{new_job}_edited_z.sdf"
                    emit_full_sdf(z_net, z_tmp, project_title=title)
                    kfp_src = z_tmp
                else:
                    kfp_src = out_sdf
                try:
                    _emit_kfp(kfp_src, out_kfp, coord_scale=kfp_scale, display_geometry=True)
                    kfp_ok = out_kfp.is_file()
                except Exception as _kexc:  # noqa: BLE001 — KFP 실패가 SDF 를 막지 않도록
                    warnings.warn(f"[combined/rebuild] KFP emit 실패 (SDF 정상): {_kexc}",
                                   RuntimeWarning, stacklevel=2)
                try:
                    _emit_has(kfp_src, out_has, isometric=True, iso_z_scale=has_zs)
                    has_ok = out_has.is_file()
                except Exception as _hexc:  # noqa: BLE001 — HAS 실패가 SDF 를 막지 않도록
                    warnings.warn(f"[combined/rebuild] HAS emit 실패 (SDF 정상): {_hexc}",
                                   RuntimeWarning, stacklevel=2)
                if z_tmp is not None:  # 임시 z-aware SDF 정리(다운로드엔 평면 out_sdf 만)
                    for _tmp in (z_tmp, z_tmp.with_suffix(".slf")):
                        try:
                            if _tmp.is_file():
                                _tmp.unlink()
                        except OSError:
                            pass
            except Exception as _zexc:  # noqa: BLE001 — KFP/HAS 전체 실패도 SDF 는 유지
                warnings.warn(f"[combined/rebuild] KFP/HAS 재출력 실패 (SDF 정상): {_zexc}",
                               RuntimeWarning, stacklevel=2)

            # 패치본을 새 job_id 로 캐시 — 연속 편집(편집→재출력→더 편집) 지원.
            # z-aware/스케일도 승계해 이후 편집에서도 KFP/HAS 비율을 유지한다.
            if len(_COMBINED_JOBS) >= _COMBINED_JOBS_CAP:
                _COMBINED_JOBS.pop(next(iter(_COMBINED_JOBS)))
            _COMBINED_JOBS[new_job] = {
                "combined": _copy.deepcopy(combined), "title": title,
                "zaware": zaware, "kfp_coord_scale": kfp_scale, "has_iso_z_scale": has_zs,
            }

            base = f"/api/remote30/combined/result/{new_job}"
            return jsonify({
                "ok": True, "job_id": new_job, "sdf": out_sdf.name,
                "kfp": out_kfp.name if kfp_ok else None,
                "has": out_has.name if has_ok else None,
                "nodes": len(combined.nodes), "pipes": len(combined.pipes),
                "nozzles": len(combined.nozzles),
                "download_url_sdf": f"{base}/{out_sdf.name}",
                "download_url_kfp": f"{base}/{out_kfp.name}" if kfp_ok else None,
                "download_url_has": f"{base}/{out_has.name}" if has_ok else None,
            })
        except Exception as exc:  # noqa: BLE001
            return _err500(exc)

    @app.get("/api/remote30/combined/result/<job_id>/<path:filename>")
    def remote30_combined_result(job_id: str, filename: str):
        return _serve_run_file(COMBINED_OUTPUT_DIR, job_id, filename)

    @app.get("/api/remote30/combined/preview/<job_id>/<fmt>")
    def remote30_combined_preview(job_id: str, fmt: str):
        """통합 결과 파일(.sdf/.kfp/.has)을 실제로 다시 파싱 → 캔버스 geometry.

        다운로드 전에 "그 포맷으로 내보낸 파일을 다시 읽으면 망이 어떻게 보이는지"
        를 미리보기로 보여주기 위함. emit 결과를 진짜로 round-trip 파싱하므로,
        포맷별 라벨/노드 누락 같은 깨짐이 있으면 그대로 드러난다(진단용).

        Query: form=plan(기본)|iso — 평면 좌표 세트 vs 등각투영 세트 선택.
        """
        safe_id = secure_filename(job_id)
        if not safe_id or safe_id != job_id:
            return jsonify({"ok": False, "message": "잘못된 job_id"}), 400
        fmt = (fmt or "").lower().lstrip(".")
        if fmt not in ("sdf", "kfp", "has"):
            return jsonify({"ok": False, "message": f"지원하지 않는 포맷: {fmt}"}), 400
        form = (request.args.get("form") or "plan").lower()
        suffix = "_iso" if form == "iso" else ""

        run_dir = COMBINED_OUTPUT_DIR / safe_id
        target = run_dir / f"combined_{safe_id}{suffix}.{fmt}"
        try:
            target.resolve().relative_to(COMBINED_OUTPUT_DIR.resolve())
        except ValueError:
            return jsonify({"ok": False, "message": "잘못된 경로"}), 400
        if not target.is_file():
            return jsonify({"ok": False,
                            "message": f"{fmt.upper()} 출력 파일이 없습니다 (emit 실패했거나 job 만료)"}), 404

        try:
            if fmt == "sdf":
                from kfp_sdf_converter import parse_sdf as _parse
            elif fmt == "kfp":
                from kfp_sdf_converter import parse_kfp as _parse
            else:
                from has_converter import parse_has as _parse
            net = _parse(str(target))
            geometry = _common_network_to_geometry(net)
        except Exception as exc:  # noqa: BLE001
            return _err500(exc, message=f"{fmt.upper()} 미리보기 파싱 실패: {str(exc)[:280]}")

        # ── 역할 사이드카로 구조 메타 복원 (Option I).
        # emit→재파싱은 라벨을 N1..Nn 으로 개명하고 라이저/헤드/AV 구분을 잃는다. 빌드 때
        # 저장한 노드 순서별 원본 라벨로 개명 라벨과 매핑해, 렌더러가 모양을 재구성하는 데
        # 쓰는 riser_labels·head_labels·head_z_frac 등을 다시 채운다. 노드 순서는 parse==emit
        # ==combined 로 동일(검증됨)하므로 인덱스 정합으로 안전하게 옮길 수 있다.
        roles_path = run_dir / f"combined_{safe_id}_roles.json"
        if roles_path.is_file():
            try:
                roles = json.loads(roles_path.read_text(encoding="utf-8"))
                order_labels = roles.get("order_labels") or []
                geo_nodes = geometry["nodes"]
                if order_labels and len(order_labels) == len(geo_nodes):
                    old2new = {str(old): str(geo_nodes[i]["label"])
                               for i, old in enumerate(order_labels)}

                    def _remap(labels):
                        return [old2new[str(l)] for l in (labels or []) if str(l) in old2new]

                    geometry["riser_labels"] = _remap(roles.get("riser_labels"))
                    geometry["head_labels"] = _remap(roles.get("head_labels"))
                    geometry["machine_room_labels"] = _remap(roles.get("machine_room_labels"))
                    _av = roles.get("av_node_label")
                    geometry["av_node_label"] = old2new.get(str(_av)) if _av is not None else None
                    _pj = roles.get("pump_junction_label")
                    if _pj is not None and str(_pj) in old2new:
                        geometry["pump_junction_label"] = old2new[str(_pj)]
                    geometry["head_orientation"] = roles.get("head_orientation", "pendent")
                    geometry["head_z_frac"] = roles.get("head_z_frac", 0.04)
                    geometry["machine_room_at_bottom"] = bool(roles.get("machine_room_at_bottom", False))
                    order_io = roles.get("order_io") or []
                    if len(order_io) == len(geo_nodes):
                        for i, gn in enumerate(geo_nodes):
                            gn["io"] = str(order_io[i])
                    geometry["pumps"] = [
                        {"label": p["label"], "in": old2new.get(str(p["in"]), str(p["in"])),
                         "out": old2new.get(str(p["out"]), str(p["out"]))}
                        for p in (roles.get("pumps") or [])]
                    geometry["valves"] = [
                        {"label": v["label"], "in": old2new.get(str(v["in"]), str(v["in"])),
                         "out": old2new.get(str(v["out"]), str(v["out"])),
                         "target_pa": v.get("target_pa", 0)}
                        for v in (roles.get("valves") or [])]
            except Exception:  # noqa: BLE001 — 사이드카 손상 시 round-trip 원본 그대로
                pass

        src = next((n for n in geometry["nodes"] if n["io"] == "Input"), None)
        return jsonify({
            "ok": True,
            "fmt": fmt,
            "form": "iso" if suffix else "plan",
            "filename": target.name,
            "nodes": len(geometry["nodes"]),
            "pipes": len(geometry["pipes"]),
            "source_label": src["label"] if src else None,
            "geometry": geometry,
        })
