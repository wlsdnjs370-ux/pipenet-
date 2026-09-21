# -*- coding: utf-8 -*-
"""역할(헤드 다는 가지관 vs 주관) × ARC 기호 조합 게이트 A/B."""
from __future__ import annotations

import bisect
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
hpts = [h.pos for h in R._find_head_candidates(pe, lc)]

HC = R.HEAD_DROP_MAX_MM
hgrid: dict = defaultdict(list)
for hx, hy in hpts:
    hgrid[(int(hx // HC), int(hy // HC))].append((hx, hy))

AC = 500.0
agrid: dict = defaultdict(list)
narc = 0
for en in bundle.entities:
    if en["t"] == "A" and en.get("c") and lc.get(en["l"]) == "PIPE":
        narc += 1
        agrid[(int(en["c"][0] // AC), int(en["c"][1] // AC))].append(tuple(en["c"]))
print(f"배관 레이어 ARC {narc} · 헤드 {len(hpts)}")


ARC_TOL = 5.0


def has_arc(p, tol=None):
    tol = ARC_TOL if tol is None else tol
    span = int(tol // AC) + 1
    cx, cy = int(p[0] // AC), int(p[1] // AC)
    for dx in range(-span, span + 1):
        for dy in range(-span, span + 1):
            for ax, ay in agrid.get((cx + dx, cy + dy), ()):
                if math.hypot(ax - p[0], ay - p[1]) <= tol:
                    return True
    return False


def edge_has_head(key):
    (x1, y1), (x2, y2) = key
    dx, dy = x2 - x1, y2 - y1
    L2 = dx * dx + dy * dy
    lo = (int(min(x1, x2) // HC), int(min(y1, y2) // HC))
    hi = (int(max(x1, x2) // HC), int(max(y1, y2) // HC))
    for cx in range(lo[0] - 1, hi[0] + 2):
        for cy in range(lo[1] - 1, hi[1] + 2):
            for hx, hy in hgrid.get((cx, cy), ()):
                t = 0.0 if L2 == 0 else max(0.0, min(1.0, ((hx-x1)*dx + (hy-y1)*dy) / L2))
                if math.hypot(hx - (x1 + t*dx), hy - (y1 + t*dy)) <= HC:
                    return True
    return False


def ends_at(key, p, tol):
    a, b = key
    return min(math.hypot(a[0]-p[0], a[1]-p[1]), math.hypot(b[0]-p[0], b[1]-p[1])) <= tol


def fresh():
    eps = R.auto_snap_eps(pe, lc)
    g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                           layer_categories=lc)
    R.collapse_parallel_ladders(g, el)
    R._split_tee_branches(g, el)
    R._drop_covered_edges(g, el)
    R._join_head_gap_endpoints(g, el, hpts)
    return g, el


def run(rule):
    g, el = fresh()
    hs, vs = [], []
    for key in el:
        a, b = key
        ax = R._axis_index(a, b, R.CROSS_TEE_AXIS_TOL_MM)
        if ax == 1:
            hs.append((min(a[0], b[0]), max(a[0], b[0]), (a[1] + b[1]) * 0.5, key))
        elif ax == 0:
            vs.append((min(a[1], b[1]), max(a[1], b[1]), (a[0] + b[0]) * 0.5, key))
    vs.sort(key=lambda t: t[2])
    vxs = [t[2] for t in vs]
    m = R.MIN_PIPE_EDGE_MM
    found = []
    for x0, x1, y, hk in hs:
        for y0, y1, x, vk in vs[bisect.bisect_left(vxs, x0 + m):
                                bisect.bisect_right(vxs, x1 - m)]:
            if y0 + m <= y <= y1 - m:
                found.append(((x, y), hk, vk))
    comps = R._connected_components(g)
    comp_of = {n: i for i, c in enumerate(comps) for n in c}
    parent: dict = {}

    def find(x):
        while parent.get(x, x) != x:
            parent[x] = parent.get(parent[x], parent[x])
            x = parent[x]
        return x

    for a, b in el:
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    def is_mark(key):
        return (len(comps[comp_of[key[0]]]) < 3
                and el[key] <= R.HEAD_GAP_JOIN_MAX_MM)

    hh: dict = {}
    cuts: dict = defaultdict(set)
    n = 0
    for p, hk, vk in sorted(found):
        if is_mark(hk) or is_mark(vk):
            continue
        ra, rb = find(hk[0]), find(vk[0])
        if ra == rb:
            continue
        if rule is not None:
            for k in (hk, vk):
                if k not in hh:
                    hh[k] = edge_has_head(k)
            if not rule(hh[hk], hh[vk], has_arc(p),
                        ends_at(hk, p, END_TOL) or ends_at(vk, p, END_TOL)):
                continue
        parent[ra] = rb
        cuts[hk].add(p)
        cuts[vk].add(p)
        n += 1
    for key, ps in cuts.items():
        a, b = key
        g[a].discard(b)
        g[b].discard(a)
        del el[key]
        chain = [a] + sorted(ps, key=lambda q: (q[0]-a[0])**2 + (q[1]-a[1])**2) + [b]
        for u, v in zip(chain, chain[1:]):
            g.setdefault(u, set()).add(v)
            g.setdefault(v, set()).add(u)
            el[(min(u, v), max(u, v))] = math.hypot(v[0]-u[0], v[1]-u[1])
    comps = R._connected_components(g)
    comp_of = {nd: i for i, c in enumerate(comps) for nd in c}
    clen: dict = defaultdict(float)
    for (a, b), L in el.items():
        clen[comp_of[a]] += L
    main = max(clen, key=clen.get)
    hit = sum(1 for hp in hpts
              if (nn := R._nearest_graph_node(g, hp)) is not None
              and comp_of[nn] == main)
    return n, hit, clen[main] / 1000, len(comps)


END_TOL = 400.0
RULES = {
    "무게이트 (현행 R8)": None,
    "C 가지×가지 금지": lambda a, b, arc, end: not (a and b),
    "E 기호 or 끝점 or 주관끼리": lambda a, b, arc, end: arc or end or not (a or b),
    "F E + 가지×가지는 기호만": lambda a, b, arc, end: (
        arc if (a and b) else (arc or end or not (a or b))),
    "G 가지×가지 금지 + 기호/끝점": lambda a, b, arc, end: (
        False if (a and b) else (arc or end or not (a or b))),
}
for etol in (200.0, 400.0, 900.0):
    END_TOL = etol
    print(f"\n── 끝점 인정 반경 {etol:.0f}mm (ARC 300mm) ──")
    ARC_TOL = 300.0
    for tag, rule in RULES.items():
        if rule is None and etol != 200.0:
            continue
        n, hit, km, nc = run(rule)
        print(f"  {tag:26s} 절단 {n:3d}  주망헤드 {hit:4d}  주망 {km:7.1f}m  조각 {nc}")
