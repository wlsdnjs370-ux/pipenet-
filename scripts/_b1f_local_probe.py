# -*- coding: utf-8 -*-
"""B1F 국소 연결성 진단 — 헤드 기둥(x~727350)이 로컬로 잘 이어져 있나?

- 같은 기둥의 인접 두 헤드 사이 그래프 최단거리 (정상이면 ~3m)
- 헤드 기둥 근처 고차수 노드(브랜치 피드) 위치
- src(자동) 가 헤드에서 얼마나 떨어졌나 + 그 사이 경로가 어디로 새는가(초반 홉 덤프)
ASCII-only.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, _build_graph,
    collapse_parallel_ladders, _bridge_components, _find_head_candidates,
    _find_source, _nearest_graph_node, _dijkstra_from, _shortest_path,
)
from core.remote30_graph import _connected_components  # noqa: E402

DXF = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"


def main():
    bundle = parse_dxf_bundle(DXF)
    layer_categories = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = filter_pipenet_only(bundle)
    graph, edge_len = _build_graph(pipe_ents, layer_categories=layer_categories)
    collapse_parallel_ladders(graph, edge_len)
    for tol in (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0):
        _bridge_components(graph, edge_len, max_bridge_mm=tol)

    heads = _find_head_candidates(pipe_ents, layer_categories)
    # 헤드 기둥: x in [727000,727500]
    col = sorted([h.pos for h in heads if 727000 <= h.pos[0] <= 727500], key=lambda p: p[1])
    print(f"head column x~727350: {len(col)} heads, y range "
          f"{col[0][1]:.0f}..{col[-1][1]:.0f}")

    comps = _connected_components(graph)
    node2comp = {}
    for i, c in enumerate(comps):
        for nd in c:
            node2comp[nd] = i

    def snap(p):
        return p if p in graph else _nearest_graph_node(graph, p)

    # 인접 두 헤드 간 그래프 거리
    if len(col) >= 2:
        a, b = snap(col[0]), snap(col[1])
        eu = math.hypot(col[0][0]-col[1][0], col[0][1]-col[1][1])
        if a and b:
            dm_a = _dijkstra_from(graph, edge_len, a)
            gd = dm_a.get(b, float("inf"))
            print(f"adjacent heads euclid={eu:.0f}mm graph={gd:,.0f}mm "
                  f"same_comp={node2comp.get(a)==node2comp.get(b)} "
                  f"(comp {node2comp.get(a)} vs {node2comp.get(b)})")

    # src 자동
    src_raw, src_kind = _find_source(pipe_ents, layer_categories)
    src_nearest = _nearest_graph_node(graph, src_raw) if src_raw else None
    if src_nearest is None:
        src = max(graph, key=lambda nd: len(graph[nd]))
        src_kind = "highest_degree"
    else:
        src = src_nearest
    print(f"src_kind={src_kind} src={tuple(round(c) for c in src)} "
          f"deg={len(graph.get(src,()))} comp={node2comp.get(src)}")

    # src 와 헤드기둥 대표 헤드의 comp 비교
    h0 = snap(col[len(col)//2])
    print(f"mid head snap={tuple(round(c) for c in h0)} comp={node2comp.get(h0)} "
          f"same_comp_as_src={node2comp.get(h0)==node2comp.get(src)}")

    # 경로 초반 20홉 덤프 — 어디서 새는지
    path = _shortest_path(graph, edge_len, src, h0)
    print(f"path hops={len(path)}  first 12 nodes:")
    for p in path[:12]:
        print(f"   ({p[0]:.0f},{p[1]:.0f})")
    print("   ...")
    for p in path[-4:]:
        print(f"   ({p[0]:.0f},{p[1]:.0f})")
    print("DONE")


if __name__ == "__main__":
    main()
