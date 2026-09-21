# -*- coding: utf-8 -*-
"""부속 기호 = ARC(끊긴 원) 라는 발견을 반영한 교차 게이트 A/B.

기호 후보를 ARC 중심 + CIRCLE 중심으로 잡고, 교차 절단 후보마다 최근접
기호 거리를 재서 반경별 절단 수 · 주망 헤드 · 주망 연장을 비교한다.
"""
from __future__ import annotations

import bisect
import math
import sys
from collections import Counter, defaultdict
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

# ── 기호 후보: ARC / CIRCLE 중심 ────────────────────────────────────────────
syms: list[tuple[float, float, float, str]] = []   # x, y, r, tag
radii = Counter()
for en in bundle.entities:
    if en["t"] not in ("A", "C"):
        continue
    c = en.get("c")
    r = float(en.get("r", 0) or 0)
    if not c or r <= 0:
        continue
    syms.append((c[0], c[1], r, f"{en['t']}:{en['l']}"))
    radii[(en["t"], round(r))] += 1
print(f"ARC/CIRCLE {len(syms)} · 헤드 {len(hpts)}")
print("  반경 상위:", ", ".join(f"{t}r{r}×{n}" for (t, r), n in radii.most_common(12)))

cell = 500.0
grid: dict = defaultdict(list)
for sx, sy, sr, tag in syms:
    grid[(int(sx // cell), int(sy // cell))].append((sx, sy, sr, tag))


def nearest_sym(p, rmax=2000.0):
    span = int(rmax // cell) + 1
    cx, cy = int(p[0] // cell), int(p[1] // cell)
    best, btag = None, ""
    for dx in range(-span, span + 1):
        for dy in range(-span, span + 1):
            for sx, sy, sr, tag in grid.get((cx + dx, cy + dy), ()):
                d = math.hypot(sx - p[0], sy - p[1])
                if best is None or d < best:
                    best, btag = d, tag
    return best, btag


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
    """_split_crossing_tees 선별 루프 + 기호 게이트 (동일 순서)."""
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
        d, _tag = nearest_sym(p)
        dists.append(d if d is not None else 1e9)
        if gate_mm is not None:
            has = d is not None and d <= gate_mm
            if has == invert:
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
print(f"\n무게이트   절단 {n0:3d}  주망헤드 {hit:4d}  주망 {km:7.1f}m  조각 {nc}")
print("절단 후보 최근접 기호 거리 분포:")
for lim in (1, 10, 30, 60, 120, 250, 500, 1000, 2000):
    print(f"    ≤{lim:5d}mm  {sum(1 for d in dists if d <= lim):4d} / {len(dists)}")
print(f"    없음     {sum(1 for d in dists if d >= 1e8):4d}")

for gate in (1.0, 10.0, 30.0, 60.0, 120.0, 250.0):
    g, el = fresh()
    n, _ = split_gated(g, el, gate)
    hit, km, nc = main_net(g, el)
    print(f"기호있음 ≤{gate:5.0f}mm 만 절단 {n:3d}  주망헤드 {hit:4d}  주망 {km:7.1f}m  조각 {nc}")

for gate in (1.0, 60.0, 250.0):
    g, el = fresh()
    n, _ = split_gated(g, el, gate, invert=True)
    hit, km, nc = main_net(g, el)
    print(f"기호없음 >{gate:5.0f}mm 만 절단 {n:3d}  주망헤드 {hit:4d}  주망 {km:7.1f}m  조각 {nc}")
