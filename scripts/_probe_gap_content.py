# -*- coding: utf-8 -*-
"""그 950mm 틈에 «무엇이» 있는가 — 비었나, 버려진 선이 있나.

이것이 갈림길이다:

    틈에 선이 있다   → 레이어 분류 문제. R10 처럼 승격하면 되고 추정이 아니다.
    틈이 비었다      → 작도 규약 해석이 필요. BLOCKED §14 가 보류한 이유.

끝점이 배관 중간을 향해 직교로 떨어진 자리를 골라, 그 사각형 안에 놓인 DXF
도형을 레이어별로 센다. 배관으로 인정된 레이어는 애초에 그래프에 있으니
여기서 세지 않는다 — 궁금한 것은 «버려진» 쪽이다.

    python scripts/_probe_gap_content.py [도면.dxf] [--n 40]
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


def _seg_pts(e):
    """entity → 선분 목록 [(a, b), …].

    파서는 좌표를 `p` 에 **평평하게** 담는다: L 은 [x1,y1,x2,y2], PL 은
    [x1,y1,x2,y2,…]. 모양을 짐작하지 말고 그대로 읽는다.
    """
    t = str(e.get("t") or "")
    if t not in ("L", "PL"):
        return []
    p = e.get("p") or []
    if not p:
        return []
    # 두 모양이 섞여 있다 — L 은 평평한 [x1,y1,x2,y2], PL 은 [[x,y], …].
    if isinstance(p[0], (list, tuple)):
        pts = [(float(q[0]), float(q[1])) for q in p if len(q) >= 2]
    else:
        if len(p) < 4:
            return []
        pts = [(float(p[i]), float(p[i + 1]))
               for i in range(0, len(p) - 1, 2)]
    return [(pts[i], pts[i + 1]) for i in range(len(pts) - 1)]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--n", type=int, default=40)
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

    graph = GRAB["graph"]
    edge_len = GRAB["edge_len"]
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

    # 600~1500mm 직교 틈만 고른다 — 봉우리 구간.
    picks = []
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
        if best is not None and 600.0 <= bd <= 1500.0:
            picks.append((bd, u, best))
    picks.sort()
    print(f"\n600~1500mm 직교 틈 {len(picks):,}곳 — 앞 {min(a.n, len(picks))}곳을 본다\n")

    # 배관으로 «인정된» 레이어는 이미 그래프에 있다 — 궁금한 것은 버려진 쪽.
    pipe_layers = {n for n, c in cat.items() if c == "PIPE"}
    egrid = defaultdict(list)
    for e in ents:
        for (p, q) in _seg_pts(e):
            x0, x1 = sorted((p[0], q[0]))
            y0, y1 = sorted((p[1], q[1]))
            for gx in range(int(x0 // CELL), int(x1 // CELL) + 1):
                for gy in range(int(y0 // CELL), int(y1 // CELL) + 1):
                    egrid[(gx, gy)].append((e, p, q))

    empty = 0
    found = Counter()
    for bd, u, (p, q) in picks[:a.n]:
        # 틈의 반대편 = 수선발
        dx, dy = q[0] - p[0], q[1] - p[1]
        L2 = dx * dx + dy * dy or 1.0
        t = max(0.0, min(1.0, ((u[0] - p[0]) * dx + (u[1] - p[1]) * dy) / L2))
        fx, fy = p[0] + t * dx, p[1] + t * dy
        mx, my = (u[0] + fx) / 2.0, (u[1] + fy) / 2.0
        gx, gy = int(mx // CELL), int(my // CELL)
        hit = Counter()
        for (e, sa, sb) in (s for c in [(gx + i, gy + j)
                                        for i in (-1, 0, 1) for j in (-1, 0, 1)]
                            for s in egrid.get(c, ())):
            ly = str(e.get("l") or "0")
            if ly in pipe_layers:
                continue
            # 그 선분이 틈 구간을 실제로 지나는가 — 중점에서 가까우면 본다.
            d = _point_to_segment_dist(mx, my, sa[0], sa[1], sb[0], sb[1])
            if d <= bd * 0.6:
                hit[f"{ly} [{cat.get(ly, '?')}]"] += 1
        if not hit:
            empty += 1
        for k, v in hit.items():
            found[k] += v

    print(f"  틈이 «완전히» 빈 곳     {empty} / {min(a.n, len(picks))}")
    print(f"  틈에 버려진 도형이 있는 곳 {min(a.n, len(picks)) - empty}\n")
    if found:
        print("  틈에서 발견된 «배관 아닌» 레이어 (많은 순)")
        for k, v in found.most_common(14):
            print(f"    {v:>5}  {k}")
    else:
        print("  ★틈에 아무 도형도 없다 — 도면에 연결선이 그려져 있지 않다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
