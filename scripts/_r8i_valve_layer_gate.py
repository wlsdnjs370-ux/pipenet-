# -*- coding: utf-8 -*-
"""부속 기호 = `6-소화-밸브` 레이어 (LINE 포함 전 종류) 기준 교차 게이트 A/B."""
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
SYM_LAYERS = {"6-소화-밸브"}

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
hpts = [h.pos for h in R._find_head_candidates(pe, lc)]


def ent_pts(en):
    t = en["t"]
    if t == "L":
        q = en["p"]
        return [(q[0], q[1]), (q[2], q[3]), ((q[0]+q[2])*0.5, (q[1]+q[3])*0.5)]
    if t in ("PL", "S", "H"):
        return [(v[0], v[1]) for v in en.get("p", [])]
    if t in ("C", "A"):
        c = en.get("c") or [0, 0]
        return [(c[0], c[1])]
    q = en.get("p") or []
    return [(q[0], q[1])] if len(q) >= 2 else []


SC = 500.0
sgrid: dict = defaultdict(list)
nsym = 0
for en in bundle.entities:
    if en["l"] not in SYM_LAYERS:
        continue
    nsym += 1
    for x, y in ent_pts(en):
        sgrid[(int(x // SC), int(y // SC))].append((x, y))
print(f"기호 엔티티 {nsym} (레이어 {sorted(SYM_LAYERS)}) · 헤드 {len(hpts)}")


def nearest_sym(p, rmax=3000.0):
    span = int(rmax // SC) + 1
    cx, cy = int(p[0] // SC), int(p[1] // SC)
    best = 1e9
    for dx in range(-span, span + 1):
        for dy in range(-span, span + 1):
            for sx, sy in sgrid.get((cx + dx, cy + dy), ()):
                d = math.hypot(sx - p[0], sy - p[1])
                if d < best:
                    best = d
    return best


def fresh():
    eps = R.auto_snap_eps(pe, lc)
    g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                           layer_categories=lc)
    R.collapse_parallel_ladders(g, el)
    R._split_tee_branches(g, el)
    R._drop_covered_edges(g, el)
    R._join_head_gap_endpoints(g, el, hpts)
    return g, el


def split_gated(g, el, gate_mm, invert=False):
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

    cuts: dict = defaultdict(set)
    n = 0
    dists = []
    for p, hk, vk in sorted(found):
        if is_mark(hk) or is_mark(vk):
            continue
        ra, rb = find(hk[0]), find(vk[0])
        if ra == rb:
            continue
        d = nearest_sym(p)
        dists.append(d)
        if gate_mm is not None and ((d <= gate_mm) == invert):
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
    return n, dists


def main_net(g, el):
    comps = R._connected_components(g)
    comp_of = {n: i for i, c in enumerate(comps) for n in c}
    clen: dict = defaultdict(float)
    for (a, b), L in el.items():
        clen[comp_of[a]] += L
    main = max(clen, key=clen.get)
    hit = sum(1 for hp in hpts
              if (nn := R._nearest_graph_node(g, hp)) is not None
              and comp_of[nn] == main)
    return hit, clen[main] / 1000, len(comps)


g, el = fresh()
n0, dists = split_gated(g, el, None)
hit, km, nc = main_net(g, el)
print(f"\n무게이트  절단 {n0:3d}  주망헤드 {hit:4d}  주망 {km:7.1f}m  조각 {nc}")
print("절단 후보 → 밸브 레이어 최근접 거리:")
for lim in (1, 30, 100, 200, 400, 800, 1500, 3000):
    print(f"    ≤{lim:5d}mm  {sum(1 for d in dists if d <= lim):4d} / {len(dists)}")

for gate in (100.0, 200.0, 400.0, 800.0):
    for inv in (False, True):
        g, el = fresh()
        n, _ = split_gated(g, el, gate, invert=inv)
        hit, km, nc = main_net(g, el)
        tag = f"{'기호없음 >' if inv else '기호있음 ≤'}{gate:4.0f}mm"
        print(f"{tag} 만 절단 {n:3d}  주망헤드 {hit:4d}  주망 {km:7.1f}m  조각 {nc}")
