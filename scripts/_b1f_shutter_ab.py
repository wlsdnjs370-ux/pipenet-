# -*- coding: utf-8 -*-
"""제외 레이어를 PIPE 로 승격했을 때의 B1F 추출 효과 A-B."""
from __future__ import annotations

import math
import sys
from collections import defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = Path(sys.argv[1])
PROMOTE = sys.argv[2:] or ["현장조사#셔터"]

bundle = R.parse_dxf_bundle_cached(DXF)
ORIG = {ly["name"]: ly["auto_category"] for ly in bundle.layers}


def extract(promote: list):
    for ly in bundle.layers:
        ly["auto_category"] = "PIPE" if ly["name"] in promote else ORIG[ly["name"]]
    try:
        lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
        pe = R.filter_pipenet_only(bundle)
        heads = R._find_head_candidates(pe, lc)
        hpts = [h.pos for h in heads]
        eps = R.auto_snap_eps(pe, lc)
        graph, edge_len = R._build_graph(
            pe, node_index=R._NodeIndex(epsilon_mm=eps), layer_categories=lc)
        R.collapse_parallel_ladders(graph, edge_len)
        R._split_tee_branches(graph, edge_len)
        R._drop_covered_edges(graph, edge_len)
        R._join_head_gap_endpoints(graph, edge_len, hpts)
        R._split_crossing_tees(graph, edge_len)
        return pe, hpts, graph, edge_len
    finally:
        for ly in bundle.layers:
            ly["auto_category"] = ORIG[ly["name"]]


def summarize(promote: list):
    pe, hpts, graph, edge_len = extract(promote)
    comps = R._connected_components(graph)
    comp_of = {n: i for i, c in enumerate(comps) for n in c}
    clen: dict = defaultdict(float)
    for (a, b), L in edge_len.items():
        clen[comp_of[a]] += L
    main = max(clen, key=clen.get)
    hit = sum(1 for hp in hpts
              if (nn := R._nearest_graph_node(graph, hp)) is not None
              and comp_of[nn] == main)
    tag = "+".join(promote) if promote else "(기본)"
    print(f"  {tag:26s} 엔티티{len(pe):5d} 헤드{len(hpts):4d} 주망헤드{hit:4d} "
          f"조각{len(comps):4d} 간선{len(edge_len):5d} 주망{clen[main]/1000:7.1f}m")
    return graph, edge_len, comp_of, main, hpts


base = summarize([])
for name in PROMOTE:
    summarize([name])
if len(PROMOTE) > 1:
    summarize(PROMOTE)

graph, edge_len, comp_of, main, hpts = base
miss = [hp for hp in hpts
        if (nn := R._nearest_graph_node(graph, hp)) is None or comp_of[nn] != main]
print(f"\n미도달 헤드 {len(miss)}개")
for hp in sorted(miss)[:40]:
    d = min(math.hypot(n[0]-hp[0], n[1]-hp[1]) for n in graph)
    print(f"    ({hp[0]:10.0f}, {hp[1]:10.0f})  최근접노드 {d:8.1f}")
