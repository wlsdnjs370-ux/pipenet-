# -*- coding: utf-8 -*-
"""승격 레이어가 끌고 들어온 지오메트리 · 그로 인한 관통 교차 절단 실측."""
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
AX = R.CROSS_TEE_AXIS_TOL_MM


def segs(en):
    p = en.get("p") or []
    if en["t"] == "L":
        return [((p[0], p[1]), (p[2], p[3]))] if len(p) >= 4 else []
    if en["t"] == "PL":
        return list(zip(p, p[1:]))
    return []


bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
print(f"승격 기록: {bundle.promoted_layers}")

# ── (a) 승격 레이어 인벤토리 ──────────────────────────────────────────────
kinds: dict = defaultdict(int)
sl = []
for en in bundle.entities:
    if en.get("l") != LAYER:
        continue
    kinds[en["t"]] += 1
    sl.extend(segs(en))
lens = sorted(math.hypot(b[0]-a[0], b[1]-a[1]) for a, b in sl)
tot = sum(lens)
print(f"\n[{LAYER}] entity {dict(kinds)} · 세그 {len(sl)} · 총연장 {tot/1000:.1f}m")
if lens:
    q = [lens[int(len(lens)*f)] for f in (0.0, 0.25, 0.5, 0.75, 0.99)]
    print("  길이 분위 min/25/50/75/99 = " + " / ".join(f"{v:.0f}" for v in q)
          + f" / max {lens[-1]:.0f}")
ax_n = defaultdict(int)
for a, b in sl:
    ax_n[R._axis_index(a, b, AX)] += 1
print(f"  축분포 (0=세로 1=가로 -1=사선): {dict(ax_n)}")

# ── (b) R8 절단이 이 레이어를 몇 번 건드리나 ────────────────────────────
shut = set()
for a, b in sl:
    shut.add((round(a[0], 1), round(a[1], 1), round(b[0], 1), round(b[1], 1)))
    shut.add((round(b[0], 1), round(b[1], 1), round(a[0], 1), round(a[1], 1)))


def on_shutter(e) -> bool:
    """간선이 승격 레이어 세그먼트와 같은 선 위에 얹혀 있나 (포함 판정)."""
    (ax_, ay), (bx, by) = e
    for sa, sb in sl:
        if R._axis_index(sa, sb, AX) != R._axis_index((ax_, ay), (bx, by), AX):
            continue
        axis = R._axis_index(sa, sb, AX)
        if axis < 0:
            continue
        run = 1 if axis == 0 else 0
        if abs((sa[1-run]+sb[1-run])*0.5 - (ay if run == 0 else ax_)) > AX:
            continue
        lo, hi = min(sa[run], sb[run]), max(sa[run], sb[run])
        e0, e1 = min(e[0][run], e[1][run]), max(e[0][run], e[1][run])
        if e0 >= lo - AX and e1 <= hi + AX:
            return True
    return False


pe = R.filter_pipenet_only(bundle)
heads = [h.pos for h in R._find_head_candidates(pe, lc)]
eps = R.auto_snap_eps(pe, lc)
g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps), layer_categories=lc)
R.collapse_parallel_ladders(g, el)
R._split_tee_branches(g, el)
R._drop_covered_edges(g, el)
R._join_head_gap_endpoints(g, el, heads)
cuts: list = []
R._split_crossing_tees(g, el, cuts_out=cuts)
print(f"\nR8 관통 교차 절단 {len(cuts)}건")
hit = [c for c in cuts
       if on_shutter((tuple(c["h"][0]), tuple(c["h"][1])))
       or on_shutter((tuple(c["v"][0]), tuple(c["v"][1])))]
print(f"  그중 승격 레이어가 한쪽인 절단 {len(hit)}건")
for c in hit[:25]:
    hl = math.dist(c["h"][0], c["h"][1])
    vl = math.dist(c["v"][0], c["v"][1])
    print(f"    교차 ({c['p'][0]:9.0f},{c['p'][1]:9.0f})  가로 {hl:8.0f}mm  세로 {vl:8.0f}mm")
