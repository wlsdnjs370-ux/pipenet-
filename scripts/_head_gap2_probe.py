# -*- coding: utf-8 -*-
"""승격 후에도 남은 미부착 헤드 — 곁의 연결관 후보가 배관에서 얼마나 모자라는가."""
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
_orig = R._finalize_selection


def _spy(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw):
    cap.update(graph=graph, heads=heads, src=src)
    return _orig(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw)


R._finalize_selection = _spy

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
print("승격 기록:", bundle.promoted_layers)
pe = R.filter_pipenet_only(bundle)
R.select_worst30_heads(pe, lc, k=30)
graph, heads, src = cap["graph"], cap["heads"], cap["src"]
miss = [h for h in heads if h.pos not in graph]
miss.sort(key=lambda h: -math.hypot(h.pos[0] - src[0], h.pos[1] - src[1]))

D = R._point_to_segment_dist
pipe_segs, cand = [], []
for en in bundle.entities:
    cat = lc.get(en["l"], "OTHER")
    pts = en.get("p") or []
    if en["t"] == "L":
        segs = [((pts[0], pts[1]), (pts[2], pts[3]))] if len(pts) >= 4 else []
    elif en["t"] == "PL":
        segs = [((a[0], a[1]), (b[0], b[1])) for a, b in zip(pts, pts[1:])]
    else:
        continue
    if cat == "PIPE":
        pipe_segs += segs
    elif cat not in R.NON_PIPE_GEOMETRY_CATS:
        cand += [(a, b, en["l"]) for a, b in segs]


def pgap(p):
    return min((D(p[0], p[1], a[0], a[1], b[0], b[1]) for a, b in pipe_segs),
               default=math.inf)


print(f"\n미부착 {len(miss)}개 — 유클리드 최원 8개")
for h in miss[:8]:
    p = h.pos
    best = None
    for a, b, ly in cand:
        da, db = math.hypot(p[0]-a[0], p[1]-a[1]), math.hypot(p[0]-b[0], p[1]-b[1])
        if min(da, db) > R.HEAD_DROP_MAX_MM:
            continue
        far = b if da <= db else a
        g = pgap(far)
        if best is None or g < best[0]:
            best = (g, ly, math.hypot(b[0]-a[0], b[1]-a[1]), min(da, db))
    eu = math.hypot(p[0]-src[0], p[1]-src[1]) / 1000
    if best is None:
        print(f"  {tuple(round(v) for v in p)} eu {eu:.1f}m · 배관 {pgap(p):.0f}mm"
              f" · 300mm 내 연결관 후보 없음")
    else:
        g, ly, L, near = best
        print(f"  {tuple(round(v) for v in p)} eu {eu:.1f}m · 배관 {pgap(p):.0f}mm"
              f" · 후보[{ly}] 길이{L:.0f} 헤드측{near:.0f} → 배관까지 {g:.0f}mm")
