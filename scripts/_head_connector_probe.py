# -*- coding: utf-8 -*-
"""미부착 헤드 곁 비-PIPE 짧은 선이 '그려진 연결관' 인가 — 반대쪽 끝이 배관에 닿는가.

헤드에서 한쪽 끝이 가깝고 반대쪽 끝이 PIPE 그래프에 닿는 선(=드롭/후렉시블 형태)을
찾는다. 닿으면 승격(실측), 안 닿으면 추정 드롭 외 방법 없음.
"""
from __future__ import annotations

import math
import sys
from collections import Counter
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = BASE / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf"

cap: dict = {}
_orig_final = R._finalize_selection


def _spy(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw):
    cap.update(graph=graph, heads=heads)
    return _orig_final(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw)


R._finalize_selection = _spy

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
R.select_worst30_heads(pe, lc, k=30)
graph, heads = cap["graph"], cap["heads"]
edges = [(u, v) for u, nbs in graph.items() for v in nbs if u < v]
miss = [h for h in heads if h.pos not in graph]


def _seg_dist(p, a, b):
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 <= 0:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2))
    return math.hypot(p[0] - (a[0] + t * vx), p[1] - (a[1] + t * vy))


NON_PIPE = {"HEAD", "TEXT", "ALARM", "ARCH", "EXCLUDE"}
cand = []   # (a, b, layer)  — 그래프에서 빠진 선분
for en in bundle.entities:
    ly = en.get("l", "")
    if lc.get(ly, "OTHER") in NON_PIPE or lc.get(ly, "OTHER") == "PIPE":
        continue
    t = en.get("t")
    if t == "L":
        p = en["p"]
        cand.append(((p[0], p[1]), (p[2], p[3]), ly))
    elif t == "PL":
        pts = en["p"]
        for a, b in zip(pts, pts[1:]):
            cand.append(((a[0], a[1]), (b[0], b[1]), ly))

CELL = 1000.0
grid = {}
for i, (a, b, ly) in enumerate(cand):
    for pt in (a, b):
        grid.setdefault((int(pt[0] // CELL), int(pt[1] // CELL)), []).append(i)


def _near_cand(p, r=400.0):
    cx, cy = int(p[0] // CELL), int(p[1] // CELL)
    seen = set()
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            seen.update(grid.get((cx + dx, cy + dy), ()))
    return [i for i in seen
            if min(math.hypot(p[0] - cand[i][k][0], p[1] - cand[i][k][1])
                   for k in (0, 1)) <= r]


def _pipe_gap(p):
    return min((_seg_dist(p, u, v) for u, v in edges), default=math.inf)


print(f"미부착 {len(miss)} · 비-PIPE 선분 {len(cand)}")
res = Counter()
gaps = []
for h in miss:
    p = h.pos
    best = None
    for i in _near_cand(p):
        a, b, ly = cand[i]
        da = math.hypot(p[0] - a[0], p[1] - a[1])
        db = math.hypot(p[0] - b[0], p[1] - b[1])
        near_end, far_end = (a, b) if da <= db else (b, a)
        if min(da, db) > 300:
            continue
        fg = _pipe_gap(far_end)
        if best is None or fg < best[0]:
            best = (fg, ly, math.hypot(b[0]-a[0], b[1]-a[1]), near_end, far_end)
    if best is None:
        res["곁에 300mm 내 끝점을 둔 비-PIPE 선 없음"] += 1
        continue
    fg = best[0]
    gaps.append(fg)
    bucket = ("먼쪽 끝이 배관에 닿음(≤50mm)" if fg <= 50 else
              "먼쪽 끝 ≤200mm" if fg <= 200 else
              "먼쪽 끝 ≤500mm" if fg <= 500 else "먼쪽 끝 >500mm")
    res[bucket] += 1

for k, n in res.most_common():
    print(f"  {n:>3}개  {k}")
if gaps:
    s = sorted(gaps)
    q = lambda f: s[min(len(s)-1, int(len(s)*f))]  # noqa: E731
    print(f"\n  먼쪽 끝 → PIPE edge: min {s[0]:.0f} · 중앙 {q(.5):.0f}"
          f" · p75 {q(.75):.0f} · max {s[-1]:.0f} mm")
