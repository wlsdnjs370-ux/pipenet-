# -*- coding: utf-8 -*-
"""가지관이 «얼마나» 떨어져 있나 — 틈의 실제 분포를 잰다.

R7(헤드틈 접속)이 0건인 이유가 「규칙이 안 돈다」가 아니라 「상한 400mm 안에
마주 보는 끝점이 없다」였다(끝점쌍 2개). 그러면 진짜 틈은 얼마인가?

두 가지를 따로 잰다 — 고쳐야 할 규칙이 다르기 때문이다:

    끝점 ↔ 끝점        R7 이 담당. 상한을 올리면 잡힌다.
    끝점 ↔ 배관 중간    R7 이 못 잡는다(T 접속). `_split_tee_branches` 담당이고
                       그쪽 상한은 TEE_SPLIT_MAX_MM 이다.

BLOCKED §14 의 미해결 「950mm 직교 단절」이 후자다. 어느 쪽이 얼마나 있는지
보고 나서 손댈 자리를 정한다.

    python scripts/_probe_gap_dist.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
OUT: dict = {}


def _hist(label, vals, edges=(200, 300, 400, 600, 800, 1000, 1500, 3000)):
    if not vals:
        print(f"  {label}: 없음")
        return
    vals = sorted(vals)
    print(f"  {label} — {len(vals):,}건 · "
          f"중앙값 {vals[len(vals) // 2]:.0f}mm")
    prev = 0
    for e in edges:
        n = sum(1 for v in vals if prev <= v < e)
        bar = "█" * int(46 * n / max(1, len(vals)))
        print(f"    {prev:>5}~{e:<5} {n:>6,}  {bar}")
        prev = e
    n = sum(1 for v in vals if v >= prev)
    print(f"    {prev:>5}~      {n:>6,}  " + "█" * int(46 * n / max(1, len(vals))))


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    import remote30_prototype as A
    from remote30_graph import HeadRegion, _point_to_segment_dist

    # 파이프라인 한복판의 그래프를 그대로 붙잡는다 — 재구성하면 입력이 달라진다.
    real = A._join_head_gap_endpoints

    def grab(graph, edge_len, head_pts, *args, **kw):
        OUT["graph"] = {k: set(v) for k, v in graph.items()}
        OUT["edge_len"] = dict(edge_len)
        OUT["heads"] = list(head_pts or ())
        return real(graph, edge_len, head_pts, *args, **kw)

    A._join_head_gap_endpoints = grab
    try:
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = bundle.entities
        cat = {}
        for nm in {str(e.get("l") or "0") for e in ents}:
            try:
                cat[nm] = A._categorize_layer(nm)
            except Exception:  # noqa: BLE001
                cat[nm] = "OTHER"
        heads = A.detect_heads(ents, cat)
        pts = [(h.pos[0], h.pos[1]) for h in heads]
        sheet = A.sheet_frame_at(pts)
        inside = pts
        if sheet is not None:
            x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
            inside = [q for q in pts if x0 <= q[0] <= x1 and y0 <= q[1] <= y1]
        cx = sum(q[0] for q in inside) / len(inside)
        cy = sum(q[1] for q in inside) / len(inside)
        alarm = min(inside, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
        zones = A.head_bbox_for_region(pts, alarm)
        A.select_worst30_heads_anchored(
            pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
            head_region=HeadRegion.from_rects(zones), zones=zones, k=30)
    finally:
        A._join_head_gap_endpoints = real

    graph = OUT.get("graph") or {}
    edge_len = OUT.get("edge_len") or {}
    hpts = OUT.get("heads") or []
    ends = [n for n, nb in graph.items() if len(nb) == 1]
    print(f"\n{dxf.name}")
    print(f"  절점 {len(graph):,} · 간선 {len(edge_len):,} · "
          f"끝점(차수1) {len(ends):,} · 헤드 {len(hpts):,}\n")

    # 연결요소 — 조각끼리의 틈만 재야 한다(같은 조각 안은 이미 이어져 있다).
    comp: dict = {}

    def find(x):
        r = x
        while comp.get(r, r) != r:
            r = comp[r]
        while comp.get(x, x) != x:
            comp[x], x = r, comp[x]
        return r

    for u, v in edge_len:
        ru, rv = find(u), find(v)
        if ru != rv:
            comp[ru] = rv

    CELL = 4000.0
    egrid = defaultdict(list)
    for n in ends:
        egrid[(int(n[0] // CELL), int(n[1] // CELL))].append(n)
    sgrid = defaultdict(list)
    for (u, v) in edge_len:
        x0, x1 = sorted((u[0], v[0]))
        y0, y1 = sorted((u[1], v[1]))
        for gx in range(int(x0 // CELL), int(x1 // CELL) + 1):
            for gy in range(int(y0 // CELL), int(y1 // CELL) + 1):
                sgrid[(gx, gy)].append((u, v))

    e2e, e2seg = [], []
    for u in ends:
        gx, gy = int(u[0] // CELL), int(u[1] // CELL)
        cells = [(gx + dx, gy + dy) for dx in (-1, 0, 1) for dy in (-1, 0, 1)]
        ru = find(u)
        best_e = None
        for v in (n for c in cells for n in egrid.get(c, ())):
            if v == u or find(v) == ru:
                continue
            d = math.hypot(v[0] - u[0], v[1] - u[1])
            if best_e is None or d < best_e:
                best_e = d
        if best_e is not None:
            e2e.append(best_e)
        best_s = None
        for (p, q) in (s for c in cells for s in sgrid.get(c, ())):
            if find(p) == ru:
                continue
            d = _point_to_segment_dist(u[0], u[1], p[0], p[1], q[0], q[1])
            if best_s is None or d < best_s:
                best_s = d
        if best_s is not None:
            e2seg.append(best_s)

    print("■ 조각과 조각 사이 — 끝점에서 가장 가까운 «다른 조각»")
    _hist("끝점 ↔ 끝점   (R7 담당)", e2e)
    print()
    _hist("끝점 ↔ 배관중간 (T접속 담당)", e2seg)

    print(f"\n  현재 상한 — R7 {A.HEAD_GAP_JOIN_MAX_MM:.0f}mm · "
          f"T접속 {A.TEE_SPLIT_MAX_MM:.0f}mm · "
          f"헤드결합 {A.HEAD_DROP_MAX_MM:.0f}mm")
    for lim in (400, 600, 800, 1000, 1500):
        n1 = sum(1 for v in e2e if v <= lim)
        n2 = sum(1 for v in e2seg if v <= lim)
        print(f"    상한 {lim:>5}mm → 끝점쌍 {n1:>6,} · 배관중간 {n2:>6,}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
