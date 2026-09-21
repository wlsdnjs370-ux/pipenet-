# -*- coding: utf-8 -*-
"""최원거리 미부착 헤드 주변 1.5 m 안의 모든 선분을 레이어·분류·길이와 함께 덤프."""
from __future__ import annotations

import math
import sys
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
    cap.update(graph=graph, heads=heads, src=src)
    return _orig_final(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw)


R._finalize_selection = _spy

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
R.select_worst30_heads(pe, lc, k=30)
graph, heads, src = cap["graph"], cap["heads"], cap["src"]
miss = [h for h in heads if h.pos not in graph]
miss.sort(key=lambda h: -math.hypot(h.pos[0] - src[0], h.pos[1] - src[1]))
ok = [h for h in heads if h.pos in graph]


def _seg_dist(p, a, b):
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 <= 0:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2))
    return math.hypot(p[0] - (a[0] + t * vx), p[1] - (a[1] + t * vy))


segs = []
for en in bundle.entities:
    t, ly = en.get("t"), en.get("l", "")
    if t == "L":
        p = en["p"]
        segs.append(((p[0], p[1]), (p[2], p[3]), ly))
    elif t == "PL":
        pts = en["p"]
        for a, b in zip(pts, pts[1:]):
            segs.append(((a[0], a[1]), (b[0], b[1]), ly))

gedges = {(min(u, v), max(u, v)) for u, nbs in graph.items() for v in nbs}


def dump(h, tag):
    p = h.pos
    print(f"\n=== {tag} 헤드 {tuple(round(v) for v in p)}"
          f" · source 유클리드 {math.hypot(p[0]-src[0], p[1]-src[1])/1000:.1f} m ===")
    near = []
    for (a, b, ly) in segs:
        d = _seg_dist(p, a, b)
        if d <= 1500:
            near.append((d, ly, math.hypot(b[0]-a[0], b[1]-a[1]), a, b))
    near.sort()
    print(f"  1.5m 내 선분 {len(near)}개 — 가까운 순 12개")
    for d, ly, L, a, b in near[:12]:
        print(f"    {d:>7.0f}mm  길이{L:>8.0f}  {lc.get(ly,'?'):<7} {ly[:34]:<34}"
              f" ({a[0]:.0f},{a[1]:.0f})→({b[0]:.0f},{b[1]:.0f})")
    gd = min((_seg_dist(p, a, b) for (a, b) in gedges), default=math.inf)
    print(f"  PIPE 그래프 edge 최근접: {gd:.0f} mm")


for i, h in enumerate(miss[:4], 1):
    dump(h, f"[미부착 최원 {i}]")
for i, h in enumerate(sorted(ok, key=lambda h: -math.hypot(h.pos[0]-src[0], h.pos[1]-src[1]))[:2], 1):
    dump(h, f"[정상부착 비교 {i}]")
