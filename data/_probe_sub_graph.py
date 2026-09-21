# -*- coding: utf-8 -*-
"""계통도·기계실 경로 그래프가 «화면으로 보낼 만한 크기» 인가.

실시간으로 선이 따라가려면 마우스가 움직일 때마다 최단경로를 다시 풀어야 한다.
서버 왕복은 LAN·터널에서 눈에 띄게 밀리므로 그래프를 화면에 한 번 보내고
브라우저가 직접 푸는 편이 낫다 — 다만 «보낼 만한 크기» 일 때만 그렇다.

그래서 먼저 잰다: 노드·간선 수와 JSON 바이트, 그리고 Dijkstra 한 번의 값.
"""
from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parents[1]
os.chdir(ROOT)
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import remote30_prototype as A            # noqa: E402
from routes.module_f.subdrawing import parse_subdrawing   # noqa: E402

CANDS = [
    "data/uploads/1. 입력도면 대명동 단위세대 계통도.dxf",
    "data/uploads/계통도_LH_306.dxf",
    "data/계통도_LH_306.dxf",
    "data/uploads/1. 입력도면 대명동 단위세대 기계실.dxf",
    "data/201동 기계실.dxf",
]

for rel in CANDS:
    p = ROOT / rel
    if not p.is_file():
        print(f"—  없음: {rel}")
        continue
    t = time.perf_counter()
    ents, diag = parse_subdrawing(str(p))
    t_parse = time.perf_counter() - t

    t = time.perf_counter()
    graph, edge_len, stats = A.build_system_graph(
        ents, layer_filter=None, force_connect=True)
    t_graph = time.perf_counter() - t

    n_nodes = len(graph)
    n_edges = sum(len(v) for v in graph.values()) // 2
    # 화면으로 보낼 모양 — 노드 좌표(정수 mm) + 간선(노드 번호쌍)
    idx = {n: i for i, n in enumerate(graph)}
    nodes = [[int(n[0]), int(n[1])] for n in graph]
    edges = sorted({(min(idx[a], idx[b]), max(idx[a], idx[b]))
                    for a, vs in graph.items() for b in vs})
    blob = json.dumps({"nodes": nodes, "edges": edges}, separators=(",", ":"))

    t = time.perf_counter()
    ns = list(graph)
    if len(ns) >= 2:
        A._shortest_path(graph, edge_len, ns[0], ns[-1])
    t_path = time.perf_counter() - t

    print(f"\n{rel}")
    print(f"   파싱 {t_parse:5.2f}s · 그래프 {t_graph:5.2f}s "
          f"· 최단경로 1회 {t_path * 1000:6.1f}ms")
    print(f"   노드 {n_nodes:,} · 간선 {n_edges:,} "
          f"· 보낼 JSON {len(blob) / 1024:,.0f} KB")
    print(f"   레이어 {len({str(e.get('l') or '0') for e in ents})}종 "
          f"· 도형 {len(ents):,}")
