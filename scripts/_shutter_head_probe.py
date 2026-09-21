# -*- coding: utf-8 -*-
"""셔터 런 주변 헤드 · 승격으로 새로 붙는 헤드가 어디 있나."""
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

DXF = BASE / "data/uploads/B1F_.dxf"
LAYER = "현장조사#셔터"

bundle = R.parse_dxf_bundle_cached(DXF)
ORIG = {ly["name"]: ly["auto_category"] for ly in bundle.layers}


def extract(promote):
    for ly in bundle.layers:
        ly["auto_category"] = "PIPE" if ly["name"] in promote else ORIG[ly["name"]]
    try:
        lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
        pe = R.filter_pipenet_only(bundle)
        hpts = [h.pos for h in R._find_head_candidates(pe, lc)]
        eps = R.auto_snap_eps(pe, lc)
        g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                               layer_categories=lc)
        R.collapse_parallel_ladders(g, el)
        R._split_tee_branches(g, el)
        R._drop_covered_edges(g, el)
        R._join_head_gap_endpoints(g, el, hpts)
        R._split_crossing_tees(g, el)
        comps = R._connected_components(g)
        co = {n: i for i, c in enumerate(comps) for n in c}
        cl: dict = defaultdict(float)
        for (a, b), L in el.items():
            cl[co[a]] += L
        main = max(cl, key=cl.get)
        ok = set()
        for hp in hpts:
            nn = R._nearest_graph_node(g, hp)
            if nn is not None and co[nn] == main:
                ok.add(hp)
        return ok
    finally:
        for ly in bundle.layers:
            ly["auto_category"] = ORIG[ly["name"]]


before = extract([])
after = extract([LAYER])
gained = sorted(after - before)
lost = sorted(before - after)
print(f"승격 전 주망헤드 {len(before)} → 후 {len(after)} · 신규 {len(gained)} · 이탈 {len(lost)}")
for hp in gained:
    print(f"  + ({hp[0]:10.0f}, {hp[1]:10.0f})")
for hp in lost:
    print(f"  - ({hp[0]:10.0f}, {hp[1]:10.0f})")

print(f"\n{LAYER} 런 주변(±600mm) 헤드")
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
hpts = [h.pos for h in R._find_head_candidates(R.filter_pipenet_only(bundle), lc)]
near = sorted(hp for hp in hpts
              if 209500 <= hp[1] <= 211200 and 714000 <= hp[0] <= 762000)
for hp in near:
    print(f"    ({hp[0]:10.0f}, {hp[1]:10.0f})  Δy(210361)={hp[1]-210361:+8.1f}")
print(f"    총 {len(near)}개")
