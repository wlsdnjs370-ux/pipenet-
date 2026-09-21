# -*- coding: utf-8 -*-
"""그래프가 «무엇으로» 이루어져 있나 — 배관인가 기호 획인가.

B1F 는 배관 레이어 선분의 66%가 300mm 미만(중앙값 149mm)이다. 그 길이는
배관이 아니라 «헤드 십자 기호» 의 획이다(BLOCKED §14 R8c 실측: 137~288mm
짜리 고립 선분이 LINE 으로 그린 헤드 기호였다).

그것들이 그래프에 들어오면:
  · 배관 위에 없는 절점이 생기고(차수 오염 → R7 이 막힌다)
  · 십자 네 획이 X 를 만들어 «분기» 로 오해되고
  · 조각이 폭증한다

그래서 그래프 간선을 길이·소속으로 갈라 «몇 %가 기호인가» 를 센다.

    python scripts/_probe_graph_mess.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import math
import statistics
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
GRAB: dict = {}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A
    from remote30_graph import HeadRegion

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    real = A._join_head_gap_endpoints

    def grab(graph, edge_len, head_pts, *ar, **kw):
        n = real(graph, edge_len, head_pts, *ar, **kw)
        GRAB["graph"] = {k: set(v) for k, v in graph.items()}
        GRAB["edge_len"] = dict(edge_len)
        GRAB["heads"] = list(head_pts or ())
        return n

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
        ins = pts
        if sheet is not None:
            x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
            ins = [q for q in pts if x0 <= q[0] <= x1 and y0 <= q[1] <= y1]
        cx = sum(q[0] for q in ins) / len(ins)
        cy = sum(q[1] for q in ins) / len(ins)
        al = min(ins, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
        zs = A.head_bbox_for_region(pts, al)
        A.select_worst30_heads_anchored(
            pipe_entities=ents, layer_categories=cat, alarm_xy=al,
            head_region=HeadRegion.from_rects(zs), zones=zs, k=30)
    finally:
        A._join_head_gap_endpoints = real

    graph, edge_len = GRAB["graph"], GRAB["edge_len"]
    hpts = GRAB["heads"]
    print(f"\n{dxf.name}")
    print(f"  절점 {len(graph):,} · 간선 {len(edge_len):,} · 헤드 {len(hpts):,}")
    print(f"  MIN_PIPE_EDGE_MM = {A.MIN_PIPE_EDGE_MM:.0f} · "
          f"SNAP_TOL_MM = {A.SNAP_TOL_MM:.0f}\n")

    # ── ① 간선 길이 분포
    ls = sorted(edge_len.values())
    print("■ 간선 길이 — 기호 획은 짧고 배관은 길다")
    edges_b = [0, 100, 200, 300, 500, 1000, 3000, 10 ** 9]
    prev = 0
    short = 0
    for e in edges_b[1:]:
        n = sum(1 for v in ls if prev <= v < e)
        if e <= 300:
            short += n
        bar = "█" * int(46 * n / max(1, len(ls)))
        lab = f"{prev}~{e}" if e < 10 ** 9 else f"{prev}~"
        print(f"    {lab:>10}mm {n:>6,}  {bar}")
        prev = e
    print(f"    중앙값 {statistics.median(ls):.0f}mm · 최대 {ls[-1]:,.0f}mm")
    print(f"    ★300mm 미만 {short:,} / {len(ls):,} "
          f"({short / max(1, len(ls)) * 100:.0f}%)\n")

    # ── ② 짧은 간선이 헤드 옆에 몰려 있나 (기호 획의 표)
    CELL = 1000.0
    hg = defaultdict(list)
    for hp in hpts:
        hg[(int(hp[0] // CELL), int(hp[1] // CELL))].append(hp)

    def near_head(p, q, r=400.0):
        mx, my = (p[0] + q[0]) / 2, (p[1] + q[1]) / 2
        gx, gy = int(mx // CELL), int(my // CELL)
        for i in (-1, 0, 1):
            for j in (-1, 0, 1):
                for hp in hg.get((gx + i, gy + j), ()):
                    if math.hypot(hp[0] - mx, hp[1] - my) <= r:
                        return True
        return False

    sh = [(k, v) for k, v in edge_len.items() if v < 300.0]
    on_head = sum(1 for (p, q), _ in sh if near_head(p, q))
    print("■ 300mm 미만 간선이 헤드 옆에 있나 (기호 획이면 그렇다)")
    print(f"    헤드 400mm 안 {on_head:,} / {len(sh):,} "
          f"({on_head / max(1, len(sh)) * 100:.0f}%)\n")

    # ── ③ 절점 차수 — 기호가 붙으면 차수가 부푼다
    deg = Counter(len(v) for v in graph.values())
    print("■ 절점 차수")
    for d in sorted(deg):
        if d > 6:
            continue
        print(f"    차수 {d}  {deg[d]:>6,}")
    hi = sum(v for k, v in deg.items() if k > 6)
    if hi:
        print(f"    차수 7+ {hi:>6,}  ★배관에 이만한 분기는 없다")

    # ── ④ 조각
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
    sizes = Counter(find(n) for n in graph)
    big = sizes.most_common(1)[0][1] if sizes else 0
    print(f"\n■ 연결 조각 {len(sizes):,}개 · 최대 조각 절점 {big:,} "
          f"({big / max(1, len(graph)) * 100:.0f}%)")
    tiny = sum(1 for v in sizes.values() if v <= 4)
    print(f"    절점 4개 이하 조각 {tiny:,}개  ★기호 획이 만든 조각일 수 있다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
