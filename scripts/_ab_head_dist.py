# -*- coding: utf-8 -*-
"""A/B: force_spanning_tree 만 격리 비교 — SPTm(현행) vs plain-SPT(구).

같은 프로세스에서 head 선정(pos)+밸브거리 signature 를 두 알고리즘으로 각각 뽑아
diff. 수리계산(밸브->헤드 거리) 불변 여부를 실제 파이프라인으로 확인한다.
"""
import heapq
import json
import math
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
import remote30_prototype as R  # noqa: E402
from core.remote30_graph import _connected_components  # noqa: E402

_SPTM = R.force_spanning_tree  # 현행 (min-weight SPT)


def _plain_spt(graph, edge_len, source=None):
    """구 로직 복제 — relax 순서 부모(첫 tight)."""
    tree_edges = set()
    for comp in _connected_components(graph):
        if not comp:
            continue
        if source is not None and source in comp:
            root = source
        else:
            root = min(comp, key=lambda p: (p[0], p[1]))
        dist = {root: 0.0}
        parent = {root: None}
        pq = [(0.0, root)]
        while pq:
            d, u = heapq.heappop(pq)
            if d > dist[u]:
                continue
            for v in graph.get(u, ()):
                if v not in comp:
                    continue
                w = edge_len.get((min(u, v), max(u, v)))
                if w is None:
                    w = math.hypot(u[0] - v[0], u[1] - v[1])
                nd = d + w
                if nd < dist.get(v, float("inf")):
                    dist[v] = nd
                    parent[v] = u
                    heapq.heappush(pq, (nd, v))
        for n, p in parent.items():
            if p is not None:
                tree_edges.add((min(n, p), max(n, p)))
    all_edges = set()
    for u, nbs in list(graph.items()):
        for v in nbs:
            all_edges.add((min(u, v), max(u, v)))
    removed = all_edges - tree_edges
    for (a, b) in removed:
        graph[a].discard(b)
        graph[b].discard(a)
        edge_len.pop((a, b), None)
    return tree_edges, removed


def sig():
    dxf = BASE / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"
    bundle = R.parse_dxf_bundle_cached(dxf)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pe = R.filter_pipenet_only(bundle)
    res = R.select_worst30_heads(pe, lc, k=30)
    return [[round(h.pos[0], 3), round(h.pos[1], 3), round(float(d), 3)]
            for h, d in zip(res.heads, res.distances)]


R.force_spanning_tree = _SPTM
after = sig()
R.force_spanning_tree = _plain_spt
before = sig()

# head set (pos) 비교 — 순서무관
set_a = {(r[0], r[1]) for r in after}
set_b = {(r[0], r[1]) for r in before}
# pos->dist 매핑 비교
map_a = {(r[0], r[1]): r[2] for r in after}
map_b = {(r[0], r[1]): r[2] for r in before}
common = set_a & set_b
dist_diffs = [(k, map_a[k], map_b[k]) for k in common if abs(map_a[k] - map_b[k]) > 1e-3]
print(json.dumps({
    "n_after": len(after),
    "n_before": len(before),
    "heads_only_in_SPTm": sorted(set_a - set_b),
    "heads_only_in_plainSPT": sorted(set_b - set_a),
    "common_heads": len(common),
    "dist_diffs_gt_1um": dist_diffs,
}, ensure_ascii=False, indent=2))
