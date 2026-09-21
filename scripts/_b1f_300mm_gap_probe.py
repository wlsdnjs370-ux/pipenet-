# -*- coding: utf-8 -*-
"""B1F 3차 단절 실측 — 헤드 조각과 주망 사이 300mm 틈의 정체."""
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
R._drop_covered_edges(graph, edge_len)
R._join_head_gap_endpoints(graph, edge_len, hpts)
R._split_crossing_tees(graph, edge_len)

comps = R._connected_components(graph)
comp_of = {n: i for i, c in enumerate(comps) for n in c}
clen: dict = defaultdict(float)
for (a, b), L in edge_len.items():
    clen[comp_of[a]] += L
main = max(clen, key=clen.get)
main_keys = [k for k in edge_len if comp_of[k[0]] == main]

hc: Counter = Counter()
for hp in hpts:
    nn = R._nearest_graph_node(graph, hp)
    if nn is not None:
        hc[comp_of[nn]] += 1


def _closest(u, a, b):
    abx, aby = b[0] - a[0], b[1] - a[1]
    L2 = abx * abx + aby * aby
    t = max(0.0, min(1.0, ((u[0]-a[0])*abx + (u[1]-a[1])*aby) / L2)) if L2 else 0.0
    q = (a[0] + t * abx, a[1] + t * aby)
    return math.hypot(u[0] - q[0], u[1] - q[1]), q, t


rows = []
for cid, cnt in hc.items():
    if cid == main:
        continue
    best = (float("inf"), None, None, None, None)
    for u in comps[cid]:
        for a, b in main_keys:
            d, q, t = _closest(u, a, b)
            if d < best[0]:
                best = (d, u, q, (a, b), t)
    d, u, q, key, t = best
    rows.append((round(d, 1), cnt, len(comps[cid]),
                 round(clen[cid] / 1000, 1), u, q, key, round(t, 3)))
rows.sort()

print(f"비주망 헤드조각 {len(rows)}개 (헤드 {sum(r[1] for r in rows)}개) "
      f"/ 주망 헤드 {hc[main]}개")
print()
for d, cnt, nn, Lm, u, q, key, t in rows:
    deg = len(graph[u])
    endish = "끝점" if t <= 0.001 or t >= 0.999 else f"중간 t={t}"
    gx, gy = q[0] - u[0], q[1] - u[1]
    ax = R._axis_index(u, q, 5.0)
    axn = {0: "세로", 1: "가로", -1: "사선"}[ax]
    # 조각 쪽 인접 간선 방향
    nbaxes = sorted({R._axis_index(u, w, 5.0) for w in graph[u]})
    print(f"  d={d:7.1f} 헤드{cnt} 노드{nn} {Lm:5.1f}m | 조각노드 deg={deg} "
          f"인접축={nbaxes} | 틈축={axn} ({gx:+.1f},{gy:+.1f}) | 주망측 {endish}")

print()
print("거리 분포:", Counter(r[0] for r in rows).most_common())
