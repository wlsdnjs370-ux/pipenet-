# -*- coding: utf-8 -*-
"""미부착 헤드 주변에 무엇이 있는가 — 배관이 필터에서 빠진 것인지, 정말 배관이 없는지.

각 미부착 헤드에서 가장 가까운 '원본 DXF 선분'을 레이어/분류와 함께 찾는다.
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
miss = [h for h in heads if h.pos not in graph]
print(f"헤드 후보 {len(heads)} · 미부착 {len(miss)}")


def _segs(ents):
    """엔티티 → (a, b, layer) 선분 리스트."""
    out = []
    for en in ents:
        t = en.get("t")
        ly = en.get("l", "")
        if t == "L":
            p = en["p"]
            out.append(((p[0], p[1]), (p[2], p[3]), ly))
        elif t == "PL":
            pts = en["p"]
            for a, b in zip(pts, pts[1:]):
                out.append(((a[0], a[1]), (b[0], b[1]), ly))
    return out


def _seg_dist(p, a, b):
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 <= 0:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2))
    return math.hypot(p[0] - (a[0] + t * vx), p[1] - (a[1] + t * vy))


all_segs = _segs(bundle.entities)
pipe_segs = _segs(pe)
print(f"원본 선분 {len(all_segs)} · 배관필터 통과 선분 {len(pipe_segs)}")

near_layer = Counter()
near_cat = Counter()
gap_all, gap_pipe = [], []
for h in miss:
    p = h.pos
    bd, bl = math.inf, ""
    for (a, b, ly) in all_segs:
        d = _seg_dist(p, a, b)
        if d < bd:
            bd, bl = d, ly
    gap_all.append(bd)
    near_layer[bl] += 1
    near_cat[lc.get(bl, "?")] += 1
    bp = min((_seg_dist(p, a, b) for (a, b, _) in pipe_segs), default=math.inf)
    gap_pipe.append(bp)


def _stat(name, arr):
    s = sorted(arr)
    q = lambda f: s[min(len(s) - 1, int(len(s) * f))]  # noqa: E731
    print(f"  {name}: min {s[0]:.0f} · 중앙 {q(.5):.0f} · p75 {q(.75):.0f} · max {s[-1]:.0f} mm")


print("\n[미부착 헤드 → 가장 가까운 선분까지 거리]")
_stat("원본 전체 레이어", gap_all)
_stat("배관필터 통과분 ", gap_pipe)
print("\n[가장 가까운 선분의 레이어]")
for ly, n in near_layer.most_common(12):
    print(f"  {n:>3}개  {lc.get(ly, '?'):<8} {ly}")
print("\n[가장 가까운 선분의 분류]")
for c, n in near_cat.most_common():
    print(f"  {n:>3}개  {c}")

print("\n[부착된 헤드의 정상 간격 — 비교군]")
ok = [h for h in heads if h.pos in graph]
gaps_ok = []
for h in ok:
    gaps_ok.append(min((_seg_dist(h.pos, a, b) for (a, b, _) in pipe_segs), default=math.inf))
_stat("배관필터 통과분 ", gaps_ok)

# 미부착 헤드 곁의 '배관필터 통과' 선분이 왜 그래프에 없나 — 레이어/분류/길이로 좁힌다.
print("\n[미부착 헤드에서 가장 가까운 '배관필터 통과' 선분의 정체]")
near_pipe = Counter()
short_cut = 0
for h in miss:
    p = h.pos
    bd, bseg = math.inf, None
    for (a, b, ly) in pipe_segs:
        d = _seg_dist(p, a, b)
        if d < bd:
            bd, bseg = d, (a, b, ly)
    if bseg is None or bd > 300:
        near_pipe[f"(300mm 밖 {bd:.0f}mm)"] += 1
        continue
    a, b, ly = bseg
    L = math.hypot(b[0] - a[0], b[1] - a[1])
    near_pipe[f"{lc.get(ly, '?'):<8} {ly}"] += 1
    if L < R.MIN_PIPE_EDGE_MM:
        short_cut += 1
for k, n in near_pipe.most_common(12):
    print(f"  {n:>3}개  {k}")
print(f"  그중 선분 길이 < MIN_PIPE_EDGE_MM({R.MIN_PIPE_EDGE_MM}mm) 라 잘렸을 것: {short_cut}개")
