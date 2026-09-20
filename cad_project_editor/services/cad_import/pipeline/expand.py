# -*- coding: utf-8 -*-
"""1단계 날것으로 펼치기 — stage1 그래프를 도면 키로 만든다.

화면 없음. 이음 규칙은 여기 없다. DXF 준비 캐시는 `pipeline.handoff`
(찍기 엔진은 `services.cad_import.pick`).
"""
import json
import math
import os
from collections import defaultdict

from services.cad_import.pipeline import handoff
from services.cad_import.pipeline import stage1 as s1

DWG = s1.DWG_DIR

CASES = {
    "MF2": "MF2 sample_libredwg.dxf",
    "3F":  "3f sample_libredwg.dxf",
    "MF3": "MF3_libredwg.dxf",
    "MF":  "MF sample_libredwg.dxf",
    "BF4": "BF4 samplle_libredwg.dxf",
    "MF4": "mf4_libredwg.dxf",
    "apt": "apt.dxf",
    # 대규모 시험용 (원 7도면 밖) — 0단계 키와 동일
    "MF101_4F_libredwg": "MF101_4F_libredwg.dxf",
    "MF101_4F_SPzone_libredwg": "MF101_4F_SPzone_libredwg.dxf",
    "MF101_흰색점선범위": "MF101_흰색점선범위.dxf",
    "sgd_libredwg": "sgd_libredwg.dxf",
    "B1F_현장조사_libredwg": "B1F_현장조사_libredwg.dxf",
}


def gput(grid, cell, x, y, v):
    grid[(int(x // cell), int(y // cell))].append(v)


def gnear(grid, cell, x, y, rings=1):
    gx, gy = int(x // cell), int(y // cell)
    out = []
    for dx in range(-rings, rings + 1):
        for dy in range(-rings, rings + 1):
            out.extend(grid.get((gx + dx, gy + dy), ()))
    return out


def seg_dist(a, b, px, py):
    """점 → 선분 거리와 t(0~1)."""
    vx, vy = b[0] - a[0], b[1] - a[1]
    ln2 = vx * vx + vy * vy
    if ln2 < 1e-9:
        return math.hypot(px - a[0], py - a[1]), 0.0
    t = ((px - a[0]) * vx + (py - a[1]) * vy) / ln2
    t = max(0.0, min(1.0, t))
    cx, cy = a[0] + vx * t, a[1] + vy * t
    return math.hypot(px - cx, py - cy), t


def _spec_path(key):
    """새찍기 우선 · 없으면 DWG 옛 경로 [2026-08-08 오너 — 저장물 분리]."""
    new = os.path.join(handoff.pick_out_dir(), f"{key}_찍은스펙.json")
    old = os.path.join(DWG, f"{key}_찍은스펙.json")
    return new if os.path.exists(new) else old


def dxf_path_for(key, spec=None):
    """찍은 스펙 source_dxf → CASES → DWG 폴더. 없으면 None."""
    if spec is None:
        spath = _spec_path(key)
        if os.path.isfile(spath):
            try:
                spec = json.load(open(spath, encoding="utf-8"))
            except Exception:  # noqa: BLE001
                spec = {}
        else:
            spec = {}
    src = (spec or {}).get("source_dxf")
    if src:
        src = os.path.abspath(src)
        if os.path.isfile(src):
            return src
    if key in CASES:
        return os.path.abspath(os.path.join(DWG, CASES[key]))
    names = []
    if str(key).lower().endswith(".dxf"):
        names.append(key)
    names.append(f"{key}.dxf")
    for name in names:
        p = name if os.path.isfile(name) else os.path.join(
            DWG, os.path.basename(name))
        if os.path.isfile(p):
            return os.path.abspath(p)
    return None


def stage1_body(key):
    """본체 `process()` 1단계와 같은 순서·인자 — pipeline.stage1 소유본."""
    spath = _spec_path(key)
    print(f"  [0 찍기] spec ← {spath}")
    spec = json.load(open(spath, encoding="utf-8"))
    source_path = dxf_path_for(key, spec)
    if not source_path or not os.path.isfile(source_path):
        raise SystemExit(f"DXF를 못 찾음: {key}")
    fname = os.path.basename(source_path)
    w = handoff.load_world(key, source_path, s1.World)
    if w is not None and len(getattr(w, "arc_ang", ())) != len(w.arcs):
        print("  [0 찍기] handoff 각 없음 — DXF 다시 펼침")
        w = None
    if w is None:
        ltab, ents, bdefs = s1.read_dxf(source_path)
        w, _hid = s1.explode(ltab, ents, bdefs)
    else:
        print("  [0 찍기] DXF 준비 handoff HIT")
    knobs = dict(s1.DEFAULT_KNOBS)
    knobs.update(spec.get("knobs", {}))
    # report 를 받아 «기호 획»(작대기·관말 캡) 명단을 챙긴다 [2026-08-07].
    # 본체가 이미 모양으로 골라 재료에서 뺀 것들이다 — 시제품이 그것을
    # 접속표시로 다시 쓴다. 짧은 선을 제 눈으로 다시 고르면 «짧은 배관»까지
    # 접속표시가 되어 없는 관을 그린다(MF2 배관 2,407→3,211m 실측).
    rep6 = {}
    mat_raw, pick_set, mat_bundles, _hg = s1.material_with_headgap(
        w, spec, report=rep6)
    mat, _a1_marks, a1_rep = s1.merge_overlaps_a1(mat_raw, knobs)
    # 헤드 원 안 가로막대기: 재료·문양은 남기고 망에만 안 넣는다 [2026-08-13 오너].
    # symbol_strokes 로 빼서 접속표시로 다시 태우지 않는다(지붕 재발).
    # 확정은 «양 끝 허공» 토막만 — 짧은 선 조각은 배관으로 남긴다 [같은 날 오너].
    circs = s1.head_circles_of(w, spec)
    bar_keys = s1.head_symbol_bar_keys(mat, circs)
    # 원시 선 bbox 격자 — B1 닿음의 30mm 연속 검사가 쓴다.
    thgrid = defaultdict(list)
    for _l, _c, a3, b3 in mat_raw:
        x0, x1 = min(a3[0], b3[0]) - 30.0, max(a3[0], b3[0]) + 30.0
        y0, y1 = min(a3[1], b3[1]) - 30.0, max(a3[1], b3[1]) + 30.0
        for gx in range(int(x0 // s1._B1_CELL), int(x1 // s1._B1_CELL) + 1):
            for gy in range(int(y0 // s1._B1_CELL), int(y1 // s1._B1_CELL) + 1):
                thgrid[(gx, gy)].append((a3, b3, (_l, _c)))
    g = s1.Graph()
    ebundle = {}
    # ★짧은 실배관이 SNAP(30mm)에 접혀 변으로 안 남는 구멍 [2026-08-07 밤3].
    #   3F 실측: 헤드↔접속부 L24.5 가 mat 에는 있는데 양 끝이 한 노드로
    #   합쳐져 add_edge 가 무시됨 → 접속부 팔 1개 → 헤드 영영 안 붙음.
    #   기호 획이 먹은 게 아니다. 실길이만 있으면 접힌 쪽 끝을 살린다.
    n_short_keep = 0
    head_symbol_bars = []
    for ly, c, a, b in mat:
        if s1.head_bar_key(a, b) in bar_keys:
            head_symbol_bars.append((a, b))
            continue
        L = math.hypot(b[0] - a[0], b[1] - a[1])
        i, j = g.node(*a), g.node(*b)
        if i == j and L > 1e-9:
            px, py = g.pts[i]
            if (math.hypot(px - a[0], py - a[1])
                    <= math.hypot(px - b[0], py - b[1])):
                j = g.force_node(*b)
            else:
                i = g.force_node(*a)
            n_short_keep += 1
        if i == j:
            continue
        g.add_edge(i, j)
        ebundle[s1._eb_key(i, j)] = (ly, c)
    raw_bar_keys = s1.head_symbol_bar_keys(mat_raw, circs)
    mat_b1 = [m for m in mat_raw
              if s1.head_bar_key(m[2], m[3]) not in raw_bar_keys]
    n_touch, _b1_side = s1.touch_tee_b1(g, mat_b1, knobs, thgrid, ebundle)
    return dict(fname=fname, spec=spec, w=w, g=g, knobs=knobs,
                mat_raw=mat_raw, pick_set=pick_set, mat_bundles=mat_bundles,
                sym_strokes=rep6.get("symbol_strokes") or [],
                ebundle=ebundle,     # 4단계(끊긴 배관 잇기)가 쓴다
                n_touch=n_touch, a1_groups=a1_rep["n_groups"],
                head_symbol_bars=head_symbol_bars)


def elen(g, edges=None):
    """m 단위 길이. edges 를 안 주면 그래프 전체."""
    return sum(math.hypot(g.pts[j][0] - g.pts[i][0],
                          g.pts[j][1] - g.pts[i][1])
               for i, j in (g.edges if edges is None else edges)) / 1000.0


def comps(g):
    """덩어리 — stage1 graph_comps(간선 묶음)를 노드 묶음으로 바꿔 쓴다."""
    return [{n for e in es for n in e} for es in s1.graph_comps(g)]
