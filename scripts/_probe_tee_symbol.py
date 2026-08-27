# -*- coding: utf-8 -*-
"""그 700~1000mm 틈 자리에 «부속 기호» 가 그려져 있나.

지금까지 좁혀진 것:
  · 끝점이 배관 «중간» 을 향해 700~1000mm 떨어져 있다 (동일선상 아님, T 접속)
  · R7(헤드틈)은 원리적으로 못 잡는다 — 사이에 헤드가 없다(153건 중 9건뿐)
  · `_split_tee_branches` 상한은 20mm 라 못 잡는다

남은 질문 하나: **그 접속 자리에 티 기호가 그려져 있는가?**

B1F 의 티 기호는 INSERT 도 CIRCLE 도 아니라 «배관 레이어에 직접 그린 ARC 쌍»
이다(BLOCKED §14 R8d 실측). A 에 그것을 찾는 `_find_fitting_symbols` 가 이미
있고, R8 이 관통 교차를 티로 볼 때 그 증거를 쓴다.

기호가 있으면 이 접속은 «추정» 이 아니라 도면이 그린 것이다 — R8d 와 같은
근거로 잡을 수 있다. 없으면 거리뿐이라 폐지된 추정 브리지가 된다.

    python scripts/_probe_tee_symbol.py [도면.dxf]
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
    ap.add_argument("--lo", type=float, default=300.0)
    ap.add_argument("--hi", type=float, default=1500.0)
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import remote30_prototype as A
    from remote30_graph import HeadRegion, _point_to_segment_dist

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    real = A._join_head_gap_endpoints

    def grab(graph, edge_len, head_pts, *args, **kw):
        n = real(graph, edge_len, head_pts, *args, **kw)
        GRAB["graph"] = {k: set(v) for k, v in graph.items()}
        GRAB["edge_len"] = dict(edge_len)
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

    # A 가 쓰는 그 탐지기 그대로 — 새 규칙을 짜지 않는다.
    fits = A._find_fitting_symbols(ents, cat)
    print(f"\n{dxf.name}")
    print(f"  배관 레이어 부속 기호 {len(fits):,}개 "
          f"(허용 반경 {A.CROSS_TEE_SYMBOL_TOL_MM:.0f}mm)")

    graph, edge_len = GRAB["graph"], GRAB["edge_len"]
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

    CELL = 4000.0
    sgrid = defaultdict(list)
    for (u, v) in edge_len:
        x0, x1 = sorted((u[0], v[0]))
        y0, y1 = sorted((u[1], v[1]))
        for gx in range(int(x0 // CELL), int(x1 // CELL) + 1):
            for gy in range(int(y0 // CELL), int(y1 // CELL) + 1):
                sgrid[(gx, gy)].append((u, v))
    fgrid = defaultdict(list)
    for f in fits:
        fgrid[(int(f[0] // CELL), int(f[1] // CELL))].append(f)

    tol = A.CROSS_TEE_SYMBOL_TOL_MM
    band = Counter()
    with_sym = Counter()
    near_u = Counter()
    for u in ends:
        gx, gy = int(u[0] // CELL), int(u[1] // CELL)
        cells = [(gx + dx, gy + dy) for dx in (-1, 0, 1) for dy in (-1, 0, 1)]
        ru = find(u)
        best, bd = None, float("inf")
        for (p, q) in (s for c in cells for s in sgrid.get(c, ())):
            if find(p) == ru:
                continue
            d = _point_to_segment_dist(u[0], u[1], p[0], p[1], q[0], q[1])
            if d < bd:
                bd, best = d, (p, q)
        if best is None or not (a.lo <= bd <= a.hi):
            continue
        key = int(bd // 200) * 200
        band[key] += 1
        p, q = best
        dx, dy = q[0] - p[0], q[1] - p[1]
        L2 = dx * dx + dy * dy or 1.0
        t = max(0.0, min(1.0, ((u[0] - p[0]) * dx + (u[1] - p[1]) * dy) / L2))
        fx, fy = p[0] + t * dx, p[1] + t * dy
        syms = [f for c in cells for f in fgrid.get(c, ())]
        # 수선발(접속점) 둘레에 기호가 있나
        if any(math.hypot(f[0] - fx, f[1] - fy) <= tol for f in syms):
            with_sym[key] += 1
        # 끝점 둘레에 기호가 있나 (가지관 쪽 표시)
        if any(math.hypot(f[0] - u[0], f[1] - u[1]) <= tol for f in syms):
            near_u[key] += 1

    tot = sum(band.values())
    print(f"\n  {a.lo:.0f}~{a.hi:.0f}mm 직교 틈 {tot:,}곳\n")
    print(f"  {'틈':>10} {'건수':>7} {'접속점에 기호':>14} {'끝점에 기호':>12}")
    print("  " + "-" * 48)
    for k in sorted(band):
        n = band[k]
        s = with_sym[k]
        e = near_u[k]
        print(f"  {k:>5}~{k + 200:<4} {n:>7,} "
              f"{s:>9,} ({s / n * 100:>3.0f}%) {e:>7,} ({e / n * 100:>3.0f}%)")
    ts, te = sum(with_sym.values()), sum(near_u.values())
    print(f"  {'합계':>10} {tot:>7,} {ts:>9,} ({ts / max(1, tot) * 100:>3.0f}%)"
          f" {te:>7,} ({te / max(1, tot) * 100:>3.0f}%)")
    print("\n  접속점에 기호가 높은 비율로 있으면 «도면이 그린 티» 다 —")
    print("  R8d 와 같은 근거로 잡을 수 있다. 낮으면 거리뿐이라 추정이 된다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
