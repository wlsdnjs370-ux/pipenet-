# -*- coding: utf-8 -*-
"""프로토타입: 최근접-이웃 컴포넌트 브리지 (agglomerative single-linkage).

각 컴포넌트를 '가장 가까운 다른 컴포넌트'에 연결(union-find greedy)해
조각들이 공간 순서대로 사슬을 이루도록 한다 (뱀 방지, 로컬 토폴로지 복원).
그 뒤 밸브->헤드 그래프 최단거리를 재서, main-centric(587m) 대비 개선 확인.
remote30_prototype.py 는 건드리지 않는다.
ASCII-only.
"""
from __future__ import annotations

import math
import sys
import time
from collections import defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, _build_graph,
    collapse_parallel_ladders, _find_head_candidates, _nearest_graph_node,
    _dijkstra_from,
)
from core.remote30_graph import _connected_components  # noqa: E402

DXF = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"
SRC = (661506.0, 177357.0)


def nn_component_bridge(graph, edge_len, max_tol=10000.0):
    """최근접 이웃 컴포넌트 브리지. 각 pass 에서 모든 컴포넌트를 그리드로 인덱싱,
    각 컴포넌트의 노드에서 '다른 컴포넌트' 최근접 노드를 찾아 후보 edge 수집,
    union-find 로 병합(각 comp 이 자기 최근접 이웃과 이어짐). tol 이내만."""
    inv = 1.0 / max_tol
    tol2 = max_tol * max_tol
    added = 0
    while True:
        comps = _connected_components(graph)
        if len(comps) <= 1:
            break
        node2comp = {}
        for i, c in enumerate(comps):
            for nd in c:
                node2comp[nd] = i
        # 그리드: 셀 -> [(node, comp)]
        grid = defaultdict(list)
        for nd, ci in node2comp.items():
            grid[(int(math.floor(nd[0] * inv)), int(math.floor(nd[1] * inv)))].append(nd)
        # 각 comp 의 최근접 (다른 comp) 쌍
        best_pair = {}  # comp_i -> (d2, u, v)
        for u, ci in node2comp.items():
            ux, uy = u
            cgx = int(math.floor(ux * inv)); cgy = int(math.floor(uy * inv))
            for dgx in (-1, 0, 1):
                for dgy in (-1, 0, 1):
                    for v in grid.get((cgx + dgx, cgy + dgy), ()):
                        cj = node2comp[v]
                        if cj == ci:
                            continue
                        dx = ux - v[0]; dy = uy - v[1]; d2 = dx * dx + dy * dy
                        if d2 > tol2:
                            continue
                        cur = best_pair.get(ci)
                        if cur is None or d2 < cur[0]:
                            best_pair[ci] = (d2, u, v)
        if not best_pair:
            break
        # union-find greedy merge
        parent = list(range(len(comps)))

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]; x = parent[x]
            return x

        pass_added = 0
        # 가까운 쌍부터 처리
        for ci, (d2, u, v) in sorted(best_pair.items(), key=lambda kv: kv[1][0]):
            cj = node2comp[v]
            if find(ci) == find(cj):
                continue
            parent[find(ci)] = find(cj)
            d = math.sqrt(d2)
            graph[u].add(v); graph[v].add(u)
            edge_len[(min(u, v), max(u, v))] = d
            added += 1; pass_added += 1
        if pass_added == 0:
            break
    return added


def main():
    t = time.time()
    bundle = parse_dxf_bundle(DXF)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pe = filter_pipenet_only(bundle)
    g, el = _build_graph(pe, layer_categories=lc)
    collapse_parallel_ladders(g, el)
    print(f"build+collapse: {time.time()-t:.1f}s comps={len(_connected_components(g)):,}")

    t = time.time()
    n = nn_component_bridge(g, el, max_tol=10000.0)
    comps = _connected_components(g)
    print(f"nn_bridge: {time.time()-t:.1f}s added={n:,} comps={len(comps):,}")

    heads = _find_head_candidates(pe, lc)
    hpos = next(h.pos for h in heads
                if 727000 <= h.pos[0] <= 727700 and 185000 <= h.pos[1] <= 188000)
    src_n = _nearest_graph_node(g, SRC)
    h_n = _nearest_graph_node(g, hpos)
    dm = _dijkstra_from(g, el, src_n)
    eu = math.hypot(SRC[0] - hpos[0], SRC[1] - hpos[1])
    pd = dm.get(h_n, float("inf"))
    print(f"euclid={eu:,.0f}mm  path={pd:,.0f}mm  ratio={pd/eu:.2f}")
    print("DONE")


if __name__ == "__main__":
    main()
