# -*- coding: utf-8 -*-
"""교차 후보를 (수평 배관 레이어 × 수직 배관 레이어 × ARC기호 유무) 로 분류."""
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
hpts = [h.pos for h in R._find_head_candidates(pe, lc)]

# ── 배관 선분 → 레이어 색인 ────────────────────────────────────────────────
segs: list[tuple[float, float, float, float, str]] = []
for en in pe:
    if lc.get(en["l"]) != "PIPE":
        continue
    if en["t"] == "L":
        q = en["p"]
        segs.append((q[0], q[1], q[2], q[3], en["l"]))
    elif en["t"] == "PL":
        pts = en.get("p") or []
        for a, b in zip(pts, pts[1:]):
            segs.append((a[0], a[1], b[0], b[1], en["l"]))
GC = 2000.0
sgrid: dict = defaultdict(list)
for i, (x1, y1, x2, y2, ly) in enumerate(segs):
    for cx in range(int(min(x1, x2) // GC), int(max(x1, x2) // GC) + 1):
        for cy in range(int(min(y1, y2) // GC), int(max(y1, y2) // GC) + 1):
            sgrid[(cx, cy)].append(i)


def _pt_seg(px, py, x1, y1, x2, y2):
    dx, dy = x2 - x1, y2 - y1
    L2 = dx * dx + dy * dy
    t = 0.0 if L2 == 0 else max(0.0, min(1.0, ((px-x1)*dx + (py-y1)*dy) / L2))
    return math.hypot(px - (x1 + t*dx), py - (y1 + t*dy))


def seg_layer(p):
    cx, cy = int(p[0] // GC), int(p[1] // GC)
    best, bly = 1e9, "?"
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for i in sgrid.get((cx + dx, cy + dy), ()):
                x1, y1, x2, y2, ly = segs[i]
                d = _pt_seg(p[0], p[1], x1, y1, x2, y2)
                if d < best:
                    best, bly = d, ly
    return bly


# ── ARC r180 색인 ──────────────────────────────────────────────────────────
AC = 500.0
agrid: dict = defaultdict(list)
for en in bundle.entities:
    if en["t"] == "A" and abs(float(en.get("r", 0) or 0) - 180) < 1 and en.get("c"):
        agrid[(int(en["c"][0] // AC), int(en["c"][1] // AC))].append(tuple(en["c"]))


def has_arc(p, tol=5.0):
    cx, cy = int(p[0] // AC), int(p[1] // AC)
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for ax, ay in agrid.get((cx + dx, cy + dy), ()):
                if math.hypot(ax - p[0], ay - p[1]) <= tol:
                    return True
    return False


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


def is_mark(key):
    return len(comps[comp_of[key[0]]]) < 3 and el[key] <= R.HEAD_GAP_JOIN_MAX_MM


def mid(key):
    a, b = key
    return ((a[0]+b[0])*0.5, (a[1]+b[1])*0.5)


tab = Counter()
samples: dict = defaultdict(list)
for p, hk, vk in sorted(found):
    if is_mark(hk) or is_mark(vk):
        continue
    ra, rb = find(hk[0]), find(vk[0])
    if ra == rb:
        continue
    parent[ra] = rb
    lh, lv = seg_layer(mid(hk)), seg_layer(mid(vk))
    pair = " × ".join(sorted((lh, lv)))
    key = (pair, "ARC" if has_arc(p) else "무기호")
    tab[key] += 1
    if len(samples[key]) < 3:
        samples[key].append(p)

print(f"교차 접속 후보 {sum(tab.values())}건\n")
print(f"{'레이어 조합':52s} {'ARC':>5s} {'무기호':>6s}")
pairs = sorted({k[0] for k in tab}, key=lambda s: -(tab[(s, 'ARC')] + tab[(s, '무기호')]))
for pr in pairs:
    print(f"  {pr:50s} {tab[(pr,'ARC')]:5d} {tab[(pr,'무기호')]:6d}")

print("\n=== 무기호 조합 샘플 좌표 ===")
for pr in pairs:
    if tab[(pr, "무기호")]:
        pts = samples[(pr, "무기호")]
        print(f"  {pr:50s} " + "  ".join(f"({q[0]:.0f},{q[1]:.0f})" for q in pts))
