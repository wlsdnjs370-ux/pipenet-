# -*- coding: utf-8 -*-
"""기호 없는 교차 접속이 정말 구조적으로 필요한가 — 영향도 순 추적 + 현장 덤프."""
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

syms = []
for en in bundle.entities:
    if en["t"] in ("A", "C") and en.get("c") and float(en.get("r", 0) or 0) > 0:
        syms.append((en["c"][0], en["c"][1], en["t"], en["l"], float(en["r"])))
SC = 500.0
sgrid: dict = defaultdict(list)
for s in syms:
    sgrid[(int(s[0] // SC), int(s[1] // SC))].append(s)


def nearest_sym(p, rmax=2000.0):
    span = int(rmax // SC) + 1
    cx, cy = int(p[0] // SC), int(p[1] // SC)
    best, btag = 1e9, "없음"
    for dx in range(-span, span + 1):
        for dy in range(-span, span + 1):
            for sx, sy, t, ly, r in sgrid.get((cx + dx, cy + dy), ()):
                d = math.hypot(sx - p[0], sy - p[1])
                if d < best:
                    best, btag = d, f"{t}r{r:.0f}@{ly}"
    return best, btag


eps = R.auto_snap_eps(pe, lc)
g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps), layer_categories=lc)
R.collapse_parallel_ladders(g, el)
R._split_tee_branches(g, el)
R._drop_covered_edges(g, el)
R._join_head_gap_endpoints(g, el, hpts)

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

rlen: dict = defaultdict(float)
for (a, b), L in el.items():
    rlen[find(a)] += L
rheads: dict = defaultdict(int)
for hp in hpts:
    nn = R._nearest_graph_node(g, hp)
    if nn is not None and math.hypot(nn[0]-hp[0], nn[1]-hp[1]) <= R.HEAD_DROP_MAX_MM:
        rheads[find(nn)] += 1


def is_mark(key):
    return len(comps[comp_of[key[0]]]) < 3 and el[key] <= R.HEAD_GAP_JOIN_MAX_MM


rows = []
for p, hk, vk in sorted(found):
    if is_mark(hk) or is_mark(vk):
        continue
    ra, rb = find(hk[0]), find(vk[0])
    if ra == rb:
        continue
    d, tag = nearest_sym(p)
    la, lb = rlen.get(ra, 0.0), rlen.get(rb, 0.0)
    ha, hb = rheads.get(ra, 0), rheads.get(rb, 0)
    rows.append((min(la, lb), p, d, tag, (la, lb), (ha, hb)))
    parent[ra] = rb
    nr = find(rb)
    rlen[nr] = la + lb
    rheads[nr] = ha + hb

rows.sort(reverse=True)
noly = [r for r in rows if r[2] > 1.0]
print(f"접속 {len(rows)}건 · 기호 1mm 초과 {len(noly)}건")
print("\n=== 영향도 상위 20 (작은 쪽 조각 길이 기준) ===")
for small, p, d, tag, (la, lb), (ha, hb) in rows[:20]:
    flag = "  " if d <= 1.0 else "★"
    print(f" {flag} ({p[0]:9.0f},{p[1]:9.0f}) 기호 {d:8.1f} {tag:22s} "
          f"조각 {la/1000:7.1f}m/{lb/1000:6.1f}m 헤드 {ha:4d}/{hb:3d}")

print("\n=== 기호 없는 접속 중 영향도 상위 6곳 현장 덤프 (400mm) ===")


def ent_pts(en):
    t = en["t"]
    if t == "L":
        q = en["p"]
        return [(q[0], q[1]), (q[2], q[3])]
    if t in ("PL", "S", "H"):
        return [(v[0], v[1]) for v in en.get("p", [])]
    if t in ("C", "A"):
        c = en.get("c") or [0, 0]
        return [(c[0], c[1])]
    q = en.get("p") or []
    return [(q[0], q[1])] if len(q) >= 2 else []


EC = 1000.0
egrid: dict = defaultdict(list)
for i, en in enumerate(bundle.entities):
    for x, y in ent_pts(en):
        egrid[(int(x // EC), int(y // EC))].append(i)


def desc(en):
    t = en["t"]
    if t == "C":
        return f"C r={en['r']:.0f}"
    if t == "A":
        a = en.get("a") or [0, 0]
        return f"A r={en['r']:.0f} {a[0]:.0f}~{a[1]:.0f}"
    if t == "I":
        return f"I {en.get('n','')}"
    if t == "L":
        q = en["p"]
        return f"L len={math.hypot(q[2]-q[0], q[3]-q[1]):.0f}"
    if t == "T":
        return f"T {en.get('v','')[:16]!r}"
    q = en.get("p", [])
    if q:
        xs = [v[0] for v in q]; ys = [v[1] for v in q]
        return f"{t} n={len(q)} bbox={max(xs)-min(xs):.0f}x{max(ys)-min(ys):.0f}"
    return t


for small, p, d, tag, _l, _h in noly[:6]:
    print(f"\n  ({p[0]:.0f}, {p[1]:.0f})  최근접기호 {d:.0f}mm {tag}  작은조각 {small/1000:.1f}m")
    span = 1
    cx, cy = int(p[0] // EC), int(p[1] // EC)
    seen = set()
    rows2 = []
    for dx in range(-span, span + 1):
        for dy in range(-span, span + 1):
            for i in egrid.get((cx + dx, cy + dy), ()):
                if i in seen:
                    continue
                seen.add(i)
                en = bundle.entities[i]
                dd = min(math.hypot(x - p[0], y - p[1]) for x, y in ent_pts(en))
                if dd <= 400:
                    rows2.append((dd, en))
    for dd, en in sorted(rows2, key=lambda t: t[0])[:12]:
        print(f"      {dd:7.1f}  {desc(en):30s} {en['l']}[{lc.get(en['l'],'?')}]")
