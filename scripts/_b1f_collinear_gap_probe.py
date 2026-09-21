# -*- coding: utf-8 -*-
"""B1F 헤드틈 접속 전후 — 조각/헤드 소속 분포."""
from __future__ import annotations

import math
import sys
from collections import Counter, defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
heads = R._find_head_candidates(pe, lc)
hpts = [h.pos for h in heads]

eps = R.auto_snap_eps(pe, lc)
graph, edge_len = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                                 layer_categories=lc)
R.collapse_parallel_ladders(graph, edge_len)
R._split_tee_branches(graph, edge_len)


def _report(tag: str) -> None:
    comps = R._connected_components(graph)
    comp_of = {n: i for i, c in enumerate(comps) for n in c}
    clen: dict = defaultdict(float)
    for (a, b), L in edge_len.items():
        clen[comp_of[a]] += L
    hc: Counter = Counter()
    far = 0
    for hp in hpts:
        nn = R._nearest_graph_node(graph, hp)
        if nn is None or math.hypot(nn[0] - hp[0], nn[1] - hp[1]) > R.HEAD_DROP_MAX_MM:
            far += 1
            continue
        hc[comp_of[nn]] += 1
    ends = sum(1 for n, nb in graph.items() if len(nb) == 1)
    print(f"[{tag}] 노드 {len(graph)} 간선 {len(edge_len)} 조각 {len(comps)} "
          f"끝점 {ends} 헤드미부착 {far}")
    top = hc.most_common(8)
    print("   헤드 보유 상위 조각 (헤드수, 노드, 연장m):",
          [(n, len(comps[c]), round(clen[c] / 1000, 1)) for c, n in top])
    print(f"   헤드 보유 조각 {len(hc)}개 / 최대 조각 {max(len(c) for c in comps)}노드 "
          f"{round(max(clen.values()) / 1000, 1)}m")


_report("접속 전")
n = R._join_head_gap_endpoints(graph, edge_len, hpts)
print(f"--- 헤드틈 접속 {n}건 ---")
_report("접속 후")
n = R._split_crossing_tees(graph, edge_len)
print(f"--- 관통 티 절단 {n}건 ---")
_report("관통 티 후")

# 남은 헤드 보유 조각이 주배관에서 얼마나 떨어져 있나
comps = R._connected_components(graph)
comp_of = {n_: i for i, c in enumerate(comps) for n_ in c}
clen: dict = defaultdict(float)
for (a, b), L in edge_len.items():
    clen[comp_of[a]] += L
main = max(clen, key=clen.get)
main_keys = [k for k in edge_len if comp_of[k[0]] == main]
hc2: Counter = Counter()
for hp in hpts:
    nn = R._nearest_graph_node(graph, hp)
    if nn is not None:
        hc2[comp_of[nn]] += 1
gaps = []
for cid, cnt in hc2.items():
    if cid == main:
        continue
    best = float("inf")
    for u in comps[cid]:
        for a, b in main_keys:
            abx, aby = b[0] - a[0], b[1] - a[1]
            L2 = abx * abx + aby * aby
            t = max(0.0, min(1.0, ((u[0]-a[0])*abx + (u[1]-a[1])*aby) / L2)) if L2 else 0.0
            d = math.hypot(u[0]-a[0]-t*abx, u[1]-a[1]-t*aby)
            best = min(best, d)
    gaps.append((round(best), cnt))
gaps.sort()
print(f"\n비주망 헤드조각 {len(gaps)}개 — 주망까지 거리(mm, 헤드수):", gaps[:20])
