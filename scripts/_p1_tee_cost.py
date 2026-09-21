# -*- coding: utf-8 -*-
"""P1 비용 스모크 — 대형 도면에서 _split_tee_branches 의 추가 시간/분할 수."""
from __future__ import annotations

import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import remote30_prototype as rp  # noqa: E402

for name in ("B1F 현장조사 소화설비 평면도.dxf", "LH 지하층배관도.dxf",
             "LH306동.dxf", "계통도_LH_306.dxf"):
    p = REPO / "samples/dxf" / name
    bundle = rp.parse_dxf_bundle(p)
    ents = rp.filter_pipenet_only(bundle)
    cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    graph, edge_len = rp._build_graph(ents, layer_categories=cats)
    rp.collapse_parallel_ladders(graph, edge_len)
    n_e, n_d = len(edge_len), sum(1 for _n, nb in graph.items() if len(nb) == 1)
    t0 = time.perf_counter()
    s = rp._split_tee_branches(graph, edge_len)
    print(f"{name:34s} edge {n_e:6d}  느슨한끝 {n_d:5d}  "
          f"split {s:5d}  {time.perf_counter() - t0:6.3f}s")
