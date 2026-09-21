# -*- coding: utf-8 -*-
"""앵커 방식 — 가장 먼 헤드에서 «배관 거리» 로 가까운 K개. 설계면적이 되나."""
from __future__ import annotations
import heapq, math, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT); sys.path.insert(0, ROOT)
import routes.module_f as mf
mf._boot(); sys.path.append(str(mf.EDITOR_ROOT))
from services.cad_import.edit.session import EditSession

es = EditSession.open("B1F 현장조사 소화설비 평면도", out_dir=None,
                      load_saved=True, use_cache=True)
b = es.board
scan = mf._autojoin_scan(b); mf._autojoin_apply(b, scan)
b._head_nodes()

adj = {}
for a, c in b.edges:
    adj.setdefault(a, []).append(c); adj.setdefault(c, []).append(a)


def dijkstra(seeds):
    dist, prev = {}, {}
    pq = []
    for s in seeds:
        dist[s] = 0.0; heapq.heappush(pq, (0.0, s))
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, 1e18): continue
        for v in adj.get(u, ()):
            nd = d + math.dist(b.pts[u], b.pts[v])
            if nd < dist.get(v, 1e18):
                dist[v] = nd; prev[v] = u; heapq.heappush(pq, (nd, v))
    return dist, prev

# 헤드별 급수원 최단거리(부착 노드)
src_dist, _ = dijkstra(list(b.sources))
head_node, head_far = {}, {}
for hi, nodes in enumerate(b.hnodes):
    r = [n for n in nodes if n in src_dist]
    if r:
        n = min(r, key=lambda x: src_dist[x])
        head_node[hi] = n; head_far[hi] = src_dist[n]
if not head_far:
    print("도달 헤드 없음"); sys.exit()

K = 30
# 앵커 = 급수원에서 가장 먼 헤드
anchor = max(head_far, key=head_far.get)
an_node = head_node[anchor]
print(f"앵커 헤드 #{anchor} · 급수원에서 {head_far[anchor]/1000:.1f}m")

# 앵커에서 배관거리로 가까운 K개 헤드
an_dist, _ = dijkstra([an_node])
cand = [(an_dist.get(head_node[hi], 1e18), hi) for hi in head_node]
cand.sort()
picked = [hi for _d, hi in cand[:K]]

hpos = [(b.disks[hi][0], b.disks[hi][1]) for hi in picked]
xs = [p[0] for p in hpos]; ys = [p[1] for p in hpos]
diag = math.hypot(max(xs)-min(xs), max(ys)-min(ys))
print(f"\n[앵커 방식 {K}개] 퍼짐 대각 {diag/1000:.1f}m  (먼순서: 95.9m)")
print(f"  앵커 기준 최원 헤드 배관거리 {cand[K-1][0]/1000:.1f}m")

# 이 K개의 corridor(급수원까지 최단경로 합집합)
_, prev = dijkstra(list(b.sources))
load = {}
for hi in picked:
    cur = head_node[hi]
    while cur in prev:
        nx = prev[cur]; key=(min(cur,nx),max(cur,nx))
        load[key]=load.get(key,0)+1; cur=nx
vals = sorted(load.values(), reverse=True)
tot = sum(math.dist(b.pts[a], b.pts[c]) for a,c in load)
print(f"  corridor 간선 {len(load)} · 총연장 {tot/1000:.1f}m · 최대 load {vals[0]}")
print(f"  → 급수원에서 앵커까지가 '최원 유하거리'(수리계산 기준압 지점)")
