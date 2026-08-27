# -*- coding: utf-8 -*-
"""R7 이 왜 0인가 — 상한이 아니라 «반대편» 이 문제인가.

상한을 1200mm 까지 올려도 접속 0건이었다. 그러면 막는 것은 거리가 아니다.
R7 은 **양쪽이 다 차수 1** 이어야 짝으로 본다(`egrid` 를 ends 로만 만든다).
반대편이 차수 2 이상이면 후보에조차 안 든다.

그래서 여기서는 «끝점이 제 축 방향으로 계속 갔을 때 만나는 첫 그래프 노드» 를
찾아 그 **차수** 를 본다. 차수가 2 이상이면 R7 이 원리적으로 못 잡는다는 뜻이고,
그때는 상한이 아니라 규칙의 모양을 봐야 한다.

    python scripts/_probe_r7_block.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import math
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
    ap.add_argument("--reach", type=float, default=2000.0)
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

    def grab(graph, edge_len, head_pts, *args, **kw):
        n = real(graph, edge_len, head_pts, *args, **kw)
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

    graph, edge_len = GRAB["graph"], GRAB["edge_len"]
    hpts = GRAB["heads"]
    tol = A.HEAD_GAP_JOIN_TOL_MM
    ends = [n for n, nb in graph.items() if len(nb) == 1]

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

    R = a.reach
    ngrid = defaultdict(list)
    for n in graph:
        ngrid[(int(n[0] // R), int(n[1] // R))].append(n)
    hgrid = defaultdict(list)
    for hp in hpts:
        hgrid[(int(hp[0] // R), int(hp[1] // R))].append(hp)

    deg_hist = Counter()
    dist_by_deg = defaultdict(list)
    head_between = Counter()
    n_axis = 0

    for u in ends:
        w = next(iter(graph[u]))
        # 런이 계속 가는 방향 = w → u 방향
        dx, dy = u[0] - w[0], u[1] - w[1]
        L = math.hypot(dx, dy)
        if L < 1e-6:
            continue
        ax = A._axis_index(w, u, tol)
        if ax < 0:
            continue          # 축평행이 아닌 끝점은 R7 대상이 아니다
        n_axis += 1
        dx, dy = dx / L, dy / L
        gx, gy = int(u[0] // R), int(u[1] // R)
        cells = [(gx + i, gy + j) for i in (-1, 0, 1) for j in (-1, 0, 1)]
        ru = find(u)
        best, bd = None, float("inf")
        for v in (n for c in cells for n in ngrid.get(c, ())):
            if v == u or find(v) == ru:
                continue
            vx, vy = v[0] - u[0], v[1] - u[1]
            fwd = vx * dx + vy * dy            # 앞쪽으로 얼마나
            off = abs(vx * dy - vy * dx)       # 축에서 얼마나 벗어났나
            if fwd <= 0 or fwd > R or off > tol:
                continue                       # 동일선상 앞쪽만
            if fwd < bd:
                bd, best = fwd, v
        if best is None:
            continue
        d = len(graph.get(best, ()))
        deg_hist[d] += 1
        dist_by_deg[d].append(bd)
        # 그 사이에 헤드가 있나 (R7 의 ③)
        got = any(
            0.0 < ((hp[0] - u[0]) * dx + (hp[1] - u[1]) * dy) < bd
            and abs((hp[0] - u[0]) * dy - (hp[1] - u[1]) * dx) <= tol
            for c in cells for hp in hgrid.get(c, ()))
        head_between[(d, got)] += 1

    print(f"\n{dxf.name}")
    print(f"  끝점(차수1) {len(ends):,} · 그중 축평행 {n_axis:,}")
    print(f"  같은 축선 앞쪽 {R:.0f}mm 안에서 «첫 노드» 를 찾은 것 "
          f"{sum(deg_hist.values()):,}\n")
    print(f"  {'첫 노드 차수':>12} {'건수':>7} {'거리 중앙값':>12} "
          f"{'사이에 헤드':>12}")
    print("  " + "-" * 48)
    for d in sorted(deg_hist):
        ds = sorted(dist_by_deg[d])
        yes = head_between[(d, True)]
        print(f"  {d:>12} {deg_hist[d]:>7,} {ds[len(ds) // 2]:>12.0f} "
              f"{yes:>12,}")
    n1 = deg_hist.get(1, 0)
    rest = sum(v for k, v in deg_hist.items() if k != 1)
    print(f"\n  차수 1 (R7 이 짝으로 볼 수 있는 것)  {n1:,}")
    print(f"  차수 2+ (R7 이 원리적으로 못 보는 것) {rest:,}")
    if rest > n1 * 3:
        print("\n  ★반대편이 대부분 차수 2 이상이다 — 상한을 올려도 안 걸린다.")
        print("    막는 것은 거리가 아니라 «양쪽 다 끝점» 조건이다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
