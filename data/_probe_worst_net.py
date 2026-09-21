# -*- coding: utf-8 -*-
"""최불리 K 선정·표시가 실제 수리계산 대상과 얼마나 어긋나나 — B1F 실측.

재는 것:
  1) «먼 순서 K개» 가 도면상 흩어지나 (설계면적 개념 위반)
  2) 화면에 그리는 최단경로 트리와, 변환이 실제로 쓰는 배관망이 같나
  3) 각 간선이 담당하는 헤드 수(load)를 지금은 세는가 (관경 결정 근거)
"""
from __future__ import annotations
import math, os, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT); sys.path.insert(0, ROOT)
import routes.module_f as mf
mf._boot(); sys.path.append(str(mf.EDITOR_ROOT))
from services.cad_import.edit.session import EditSession

es = EditSession.open("B1F 현장조사 소화설비 평면도", out_dir=None,
                      load_saved=True, use_cache=True)
b = es.board
# 자동 이음 적용(실사용 흐름과 동일)
scan = mf._autojoin_scan(b); mf._autojoin_apply(b, scan)
b._head_nodes()
print(f"손질망 · 노드 {len(b.pts)} · 간선 {len(b.edges)} · 헤드 {len(b.disks)}")

st = b.water_state()
print(f"물 닿는 헤드 {len(st['wet_heads'])}/{len(b.disks)}")

K = 30
w = mf._worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=K)
print(f"\n[먼 순서 {K}개] 최원 {w['far_m']}m · 끝 {w['near_m']}m · 경로간선 {len(w['edges'])}")

# 1) 흩어짐 — 뽑힌 K개 헤드의 좌표 퍼짐(bounding box 대각선)
hpos = [(b.disks[hi][0], b.disks[hi][1]) for hi in w["heads"]]
xs = [p[0] for p in hpos]; ys = [p[1] for p in hpos]
diag = math.hypot(max(xs)-min(xs), max(ys)-min(ys))
# 헤드 간 평균 최근접(설계면적이면 촘촘해야 함)
nn = []
for i, p in enumerate(hpos):
    d = min((math.hypot(p[0]-q[0], p[1]-q[1])
             for j, q in enumerate(hpos) if j != i), default=0)
    nn.append(d)
print(f"  헤드 {K}개 퍼짐: 대각 {diag/1000:.1f}m · 최근접 중앙값 {sorted(nn)[len(nn)//2]/1000:.2f}m")
print(f"  → 설계면적이라면 대각이 수 m 수준이어야 한다. 지금은?")

# 2) load — 각 경로 간선이 담당하는 헤드 수를 세어본다(현재 F 엔 없는 셈)
adj = {}
for a, c in b.edges:
    adj.setdefault(a, []).append(c); adj.setdefault(c, []).append(a)
import heapq
dist, prev = {}, {}
pq = [(0.0, s) for s in b.sources]
for s in b.sources: dist[s] = 0.0
while pq:
    d, u = heapq.heappop(pq)
    if d > dist.get(u, 1e18): continue
    for v in adj.get(u, ()):
        nd = d + math.dist(b.pts[u], b.pts[v])
        if nd < dist.get(v, 1e18):
            dist[v] = nd; prev[v] = u; heapq.heappush(pq, (nd, v))
# 각 최불리 헤드 → 급수원 경로의 간선마다 +1
load = {}
for hi in w["heads"]:
    nodes = [n for n in b.hnodes[hi] if n in dist]
    if not nodes: continue
    cur = min(nodes, key=lambda n: dist[n])
    while cur in prev:
        nx = prev[cur]; key = (min(cur, nx), max(cur, nx))
        load[key] = load.get(key, 0) + 1; cur = nx
if load:
    vals = sorted(load.values(), reverse=True)
    print(f"\n[담당 헤드 수(load)] 간선 {len(load)}개 · 최대 {vals[0]} (주배관)"
          f" · load=1 가지 {sum(1 for v in vals if v==1)}개")
    print(f"  → 이 최대값이 NFPC 별표1 의 관경 결정값. 지금 F 는 이걸 안 센다.")
