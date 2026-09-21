# -*- coding: utf-8 -*-
"""대명동 — "가장 먼 헤드"가 top-30 에서 빠지는 원인 실측.

헤드 후보 전수를 (a) 배관망 부착 실패 (b) 부착됐으나 source 도달 불가
(c) 도달 — 세 부류로 분류하고, 유클리드 최원거리 헤드가 어디로 갔는지 본다.
"""
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
    cap.update(graph=graph, edge_len=edge_len, src=src, heads=heads, k=k)
    return _orig_final(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw)


R._finalize_selection = _spy

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
res = R.select_worst30_heads(pe, lc, k=30)

graph, edge_len, src, heads = cap["graph"], cap["edge_len"], cap["src"], cap["heads"]
dist = R._dijkstra_from(graph, edge_len, src)

rows = []
for h in heads:
    attached = h.pos in graph
    d = dist.get(h.pos, float("inf")) if attached else float("inf")
    eu = math.hypot(h.pos[0] - src[0], h.pos[1] - src[1])
    kind = "도달" if math.isfinite(d) else ("미도달(부착됨)" if attached else "미부착")
    rows.append({"pos": h.pos, "eu": eu, "d": d, "kind": kind})

n_ok = sum(1 for r in rows if r["kind"] == "도달")
n_iso = sum(1 for r in rows if r["kind"] == "미도달(부착됨)")
n_det = sum(1 for r in rows if r["kind"] == "미부착")
print(f"DXF: {DXF.name}")
print(f"source={tuple(round(v, 1) for v in src)} · 헤드 후보 {len(heads)}개")
print(f"  도달 {n_ok} · 미도달(부착됐지만 source 와 단절) {n_iso} · 미부착(300mm 내 배관 없음) {n_det}")
print(f"  → 선정 {len(res.heads)}개 (최원 {max(res.distances):.0f}mm / 최근 {min(res.distances):.0f}mm)")

sel = {h.pos for h in res.heads}
print("\n[유클리드 최원거리 헤드 20개] — 실제로 멀리 있는 헤드가 선정됐는가")
print(f"  {'유클리드m':>9} {'배관경로m':>10}  {'상태':<16} 선정")
for r in sorted(rows, key=lambda r: -r["eu"])[:20]:
    dm = f"{r['d']/1000:.1f}" if math.isfinite(r["d"]) else "-"
    print(f"  {r['eu']/1000:>9.1f} {dm:>10}  {r['kind']:<16} {'O' if r['pos'] in sel else 'X'}")

print("\n[배관경로 최원거리 헤드 35개] — 정렬 자체는 맞는가")
for i, r in enumerate(sorted([r for r in rows if math.isfinite(r["d"])],
                            key=lambda r: -r["d"])[:35], 1):
    print(f"  {i:>3}. {r['d']/1000:>8.1f} m  {'O' if r['pos'] in sel else 'X'}")

def _seg_dist(p, a, b):
    """점 p 에서 선분 ab 까지의 거리와 수선발 파라미터 t(0~1)."""
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 <= 0:
        return math.hypot(p[0] - a[0], p[1] - a[1]), 0.0
    t = ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2
    t = max(0.0, min(1.0, t))
    return math.hypot(p[0] - (a[0] + t * vx), p[1] - (a[1] + t * vy)), t


if n_det:
    nodes = list(graph)
    edges = {(min(u, v), max(u, v)) for u, nbs in graph.items() for v in nbs}
    miss = [r for r in rows if r["kind"] == "미부착"]
    print(f"\n[미부착 헤드 {len(miss)}개] 노드(끝점) 거리 vs 선분(배관 몸통) 거리"
          f"  — 상한 {R.HEAD_DROP_MAX_MM}mm")
    dn, de, interior = [], [], 0
    for r in miss:
        p = r["pos"]
        dn.append(min((math.hypot(p[0] - n[0], p[1] - n[1]) for n in nodes), default=math.inf))
        best_d, best_t = math.inf, 0.0
        for (a, b) in edges:
            d, t = _seg_dist(p, a, b)
            if d < best_d:
                best_d, best_t = d, t
        de.append(best_d)
        if best_d <= R.HEAD_DROP_MAX_MM and 1e-6 < best_t < 1 - 1e-6:
            interior += 1
    for name, arr in (("끝점까지", dn), ("선분까지", de)):
        s = sorted(arr)
        q = lambda p: s[min(len(s) - 1, int(len(s) * p))]  # noqa: E731
        print(f"  {name}: min {s[0]:.0f} · p25 {q(.25):.0f} · 중앙 {q(.5):.0f}"
              f" · p75 {q(.75):.0f} · max {s[-1]:.0f} mm")
    n_rescue = sum(1 for d in de if d <= R.HEAD_DROP_MAX_MM)
    print(f"  → 선분 기준이면 상한 이내 {n_rescue}/{len(miss)}개 구제 가능"
          f" (그중 배관 몸통 중간에 붙는 것 {interior}개)")
