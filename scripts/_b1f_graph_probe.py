# -*- coding: utf-8 -*-
"""B1F 그래프 위상 진단 — 루프 유무 / 직선 대비 그래프 최단경로.

select_worst30_heads 의 전처리(build+collapse+bridge)를 재현한 뒤:
  - 노드/엣지/컴포넌트 수, cycle 수 (edges - nodes + comps)  == 0 이면 forest(트리)
  - 최원거리 헤드에 대해 full-graph Dijkstra 최단 vs euclid
그래서 686m 가 '루프 절단' 문제인지 '직결 누락' 문제인지 판별.
ASCII-only 출력.
"""
from __future__ import annotations

import math
import sys
import time
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
    print(f"pipe_ents={len(pipe_ents):,}")

    graph, edge_len = _build_graph(pipe_ents, layer_categories=layer_categories)
    collapse_parallel_ladders(graph, edge_len)
    print(f"after build+collapse: nodes={len(graph):,} "
          f"edges={sum(len(v) for v in graph.values())//2:,} "
          f"comps={len(_connected_components(graph))}")
    for tol in (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0):
        _bridge_components(graph, edge_len, max_bridge_mm=tol)
    n = len(graph)
    e = sum(len(v) for v in graph.values()) // 2
    comps = _connected_components(graph)
    cyc = e - n + len(comps)
    print(f"after bridge: nodes={n:,} edges={e:,} comps={len(comps)} "
          f"cycles={cyc}  ({'FOREST/TREE' if cyc == 0 else 'HAS LOOPS'})")
    biggest = max(comps, key=len)
    print(f"biggest comp: nodes={len(biggest):,}")

    # source + worst head on full graph
    heads = _find_head_candidates(pipe_ents, layer_categories)
    src_raw, src_kind = _find_source(pipe_ents, layer_categories)
    src_nearest = _nearest_graph_node(graph, src_raw) if src_raw else None
    if src_nearest is None:
        src = max(graph, key=lambda nd: len(graph[nd]))
        src_kind = "highest_degree"
    else:
        src = src_nearest
    print(f"src_kind={src_kind} src={tuple(round(c) for c in src)}")

    dm = _dijkstra_from(graph, edge_len, src)
    # farthest reachable head
    best = None
    for h in heads:
        node = h.pos if h.pos in graph else _nearest_graph_node(graph, h.pos)
        if node is None:
            continue
        d = dm.get(node, float("inf"))
        if math.isfinite(d) and (best is None or d > best[0]):
            best = (d, node, h.pos)
    if best is None:
        print("no reachable head")
        return
    d, node, hpos = best
    eu = math.hypot(hpos[0] - src[0], hpos[1] - src[1])
    path = _shortest_path(graph, edge_len, src, node)
    print(f"\nfarthest head: euclid={eu:,.0f}mm  full-graph shortest={d:,.0f}mm "
          f"ratio={d/eu:.1f}  path_hops={len(path)}")
    # bbox span of the path to see how far it wanders
    xs = [p[0] for p in path]; ys = [p[1] for p in path]
    print(f"path bbox: x[{min(xs):.0f},{max(xs):.0f}] y[{min(ys):.0f},{max(ys):.0f}]")
    print("=> If full-graph shortest is already ~686k, the graph itself lacks a")
    print("   direct connection (missing bridge), NOT a loop-cut problem.")
    print("DONE")


if __name__ == "__main__":
    main()
