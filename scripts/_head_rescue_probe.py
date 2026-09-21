# -*- coding: utf-8 -*-
"""미부착 헤드를 '가장 가까운 PIPE edge' 로 구제할 때의 안전 마진 측정.

각 헤드에 대해 (1) 최근접 PIPE edge 거리 (2) 그 edge 를 뺀 다음으로 가까운
'다른 가지' 까지의 거리 를 재서, 붙일 곳이 유일하게 가까운지(=오부착 위험)를 본다.
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

TARGETS = [
    ("대명동", BASE / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf"),
    ("B1F최소", BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"),
]

cap: dict = {}
_orig_final = R._finalize_selection


def _spy(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw):
    cap.update(graph=graph, heads=heads, src=src)
    return _orig_final(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw)


R._finalize_selection = _spy


def _seg_dist(p, a, b):
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 <= 0:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2))
    return math.hypot(p[0] - (a[0] + t * vx), p[1] - (a[1] + t * vy))


def _stat(name, arr):
    if not arr:
        print(f"  {name}: (없음)")
        return
    s = sorted(arr)
    q = lambda f: s[min(len(s) - 1, int(len(s) * f))]  # noqa: E731
    print(f"  {name}: min {s[0]:.0f} · p25 {q(.25):.0f} · 중앙 {q(.5):.0f}"
          f" · p75 {q(.75):.0f} · p90 {q(.9):.0f} · max {s[-1]:.0f} mm")


for label, dxf in TARGETS:
    if not dxf.exists():
        print(f"\n### {label}: 파일 없음 — 건너뜀")
        continue
    cap.clear()
    bundle = R.parse_dxf_bundle_cached(dxf)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pe = R.filter_pipenet_only(bundle)
    R.select_worst30_heads(pe, lc, k=30)
    graph, heads = cap["graph"], cap["heads"]
    edges = [(u, v) for u, nbs in graph.items() for v in nbs if u < v]
    miss = [h for h in heads if h.pos not in graph]
    print(f"\n### {label} — 헤드 {len(heads)} · 미부착 {len(miss)} · edge {len(edges)}")

    d1, ratio, far = [], [], Counter()
    for h in miss:
        p = h.pos
        ds = sorted((_seg_dist(p, u, v), u, v) for u, v in edges)
        if not ds:
            continue
        best = ds[0]
        d1.append(best[0])
        # 최근접 edge 와 노드를 공유하지 않는(=다른 가지) 첫 edge
        touch = {best[1], best[2]}
        alt = next((d for d, u, v in ds[1:] if u not in touch and v not in touch), math.inf)
        ratio.append(alt / best[0] if best[0] > 0 else math.inf)
        for cut in (300, 500, 800, 1000, 1500, 2000):
            if best[0] <= cut:
                far[cut] += 1
    _stat("미부착 헤드 → 최근접 PIPE edge", d1)
    _stat("(그 다음 다른 가지까지)/(최근접) 배수", ratio)
    print("  상한별 구제 가능 수: " + " · ".join(
        f"{c}mm→{far[c]}/{len(miss)}" for c in (300, 500, 800, 1000, 1500, 2000)))

    okd = []
    for h in heads:
        if h.pos in graph:
            okd.append(min((_seg_dist(h.pos, u, v) for u, v in edges), default=0.0))
    _stat("정상 부착 헤드 → 최근접 PIPE edge", okd)
