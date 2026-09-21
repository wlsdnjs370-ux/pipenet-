# -*- coding: utf-8 -*-
"""R8c 헤드 증거 게이트 — 실제 _split_crossing_tees 로 B1F 재실측 (probe 대비 검증)."""
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

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "data/uploads/B1F_.dxf"

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
heads = R._find_head_candidates(pe, lc)
hpts = [h.pos for h in heads]


def run(head_pts):
    eps = R.auto_snap_eps(pe, lc)
    g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                           layer_categories=lc)
    R.collapse_parallel_ladders(g, el)
    R._split_tee_branches(g, el)
    R._drop_covered_edges(g, el)
    R._join_head_gap_endpoints(g, el, hpts)
    n = R._split_crossing_tees(g, el, head_pts=head_pts)
    comps = R._connected_components(g)
    comp_of = {node: i for i, c in enumerate(comps) for node in c}
    clen: dict = defaultdict(float)
    for (a, b), L in el.items():
        clen[comp_of[a]] += L
    main = max(clen, key=clen.get)
    hit = sum(1 for hp in hpts
              if (nn := R._nearest_graph_node(g, hp)) is not None
              and comp_of[nn] == main)
    tag = "게이트ON " if head_pts is not None else "게이트OFF"
    print(f"{tag} 절단 {n:3d}  주망헤드 {hit:4d}  주망 {clen[main]/1000:7.1f}m  조각 {len(comps)}")


run(None)
run(hpts)
