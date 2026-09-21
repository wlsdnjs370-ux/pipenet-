# -*- coding: utf-8 -*-
"""R8 관통 교차 절단에 부속기호 게이트를 걸면 어떻게 되나 — 반경별 A/B.

교차 절단 후보(합류+비표시획)마다 최근접 기호 거리를 재고,
게이트 반경별로 절단 수 · 주망 헤드 · 주망 연장을 비교한다.
"""
from __future__ import annotations

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
SYM_CIRCLE_MAX_R = 300.0

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
heads = R._find_head_candidates(R.filter_pipenet_only(bundle), lc)
hpts = [h.pos for h in heads]

# ── 부속기호 후보: 비헤드 INSERT + 작은 원(비 HEAD 레이어) ─────────────────
syms: list[tuple[float, float, str]] = []
for en in bundle.entities:
    cat = lc.get(en.get("l", ""), "OTHER")
    if en["t"] == "I":
        p = en.get("p") or []
        name = en.get("n", "") or ""
        if len(p) >= 2 and name not in R.KNOWN_HEAD_BLOCKS and cat != "HEAD":
            syms.append((p[0], p[1], f"I:{name}"))
    elif en["t"] == "C" and cat != "HEAD":
        r = float(en.get("r", 0) or 0)
        c = en.get("c")
        if c and 0 < r <= SYM_CIRCLE_MAX_R:
            syms.append((c[0], c[1], f"C:r{r:.0f}@{en.get('l','')}"))
print(f"기호후보 {len(syms)} (INSERT {sum(1 for s in syms if s[2].startswith('I'))}"
      f" / 원 {sum(1 for s in syms if s[2].startswith('C'))}) · 헤드 {len(hpts)}")

cell = 1000.0
grid: dict = defaultdict(list)
for sx, sy, nm in syms:
    grid[(int(sx // cell), int(sy // cell))].append((sx, sy, nm))


def nearest_sym(p, rmax=3000.0):
    best, bnm = None, ""
    span = int(rmax // cell) + 1
    cx, cy = int(p[0] // cell), int(p[1] // cell)
    for dx in range(-span, span + 1):
        for dy in range(-span, span + 1):
            for sx, sy, nm in grid.get((cx + dx, cy + dy), ()):
                d = math.hypot(sx - p[0], sy - p[1])
                if best is None or d < best:
                    best, bnm = d, nm
    return best, bnm


def fresh():
    eps = R.auto_snap_eps(pe, lc)
    g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                           layer_categories=lc)
    R.collapse_parallel_ladders(g, el)
    R._split_tee_branches(g, el)
    R._drop_covered_edges(g, el)
    R._join_head_gap_endpoints(g, el, hpts)
    return g, el


def head_nodes(g):
    """헤드 → 최근접 노드 (HEAD_DROP_MAX_MM 이내만) 카운트."""
    cnt: dict = defaultdict(int)
    for hp in hpts:
        nn = R._nearest_graph_node(g, hp)
        if nn is not None and math.hypot(nn[0]-hp[0], nn[1]-hp[1]) <= R.HEAD_DROP_MAX_MM:
            cnt[nn] += 1
    return cnt


def split_gated(g, el, gate_mm, head_gate=False):
    """_split_crossing_tees 의 선별 루프 + 기호 게이트 (동일 순서)."""
    hs, vs = [], []
    for key in el:
        a, b = key
        ax = R._axis_index(a, b, R.CROSS_TEE_AXIS_TOL_MM)
        if ax == 1:
            hs.append((min(a[0], b[0]), max(a[0], b[0]), (a[1] + b[1]) * 0.5, key))
        elif ax == 0:
            vs.append((min(a[1], b[1]), max(a[1], b[1]), (a[0] + b[0]) * 0.5, key))
    import bisect
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

    hn = head_nodes(g) if head_gate else {}
    root_heads: dict = defaultdict(int)
    if head_gate:
        for node, c in hn.items():
            root_heads[find(node)] += c

    def is_mark(key):
        return (len(comps[comp_of[key[0]]]) < 3
                and el[key] <= R.HEAD_GAP_JOIN_MAX_MM)

    cuts: dict = defaultdict(set)
    n = 0
    dists = []
    side_stats = Counter()
    for p, hk, vk in sorted(found):
        if is_mark(hk) or is_mark(vk):
            continue
        ra, rb = find(hk[0]), find(vk[0])
        if ra == rb:
            continue
        d, _nm = nearest_sym(p)
        dists.append(d if d is not None else 1e9)
        if head_gate:
            ha, hb = root_heads.get(ra, 0), root_heads.get(rb, 0)
            side_stats["양쪽" if (ha and hb) else "한쪽" if (ha or hb) else "0쪽"] += 1
            if head_gate == "both" and not (ha and hb):
                continue
            if head_gate == "any" and not (ha or hb):
                continue
        elif gate_mm is not None and (d is None or d > gate_mm):
            continue
        ra, rb = find(hk[0]), find(vk[0])
        parent[ra] = rb
        if head_gate:
            root_heads[find(rb)] = root_heads.pop(ra, 0) + root_heads.get(find(rb), 0)
        cuts[hk].add(p)
        cuts[vk].add(p)
        n += 1
    for key, pts in cuts.items():
        a, b = key
        g[a].discard(b)
        g[b].discard(a)
        del el[key]
        chain = [a] + sorted(pts, key=lambda q: (q[0]-a[0])**2 + (q[1]-a[1])**2) + [b]
        for u, v in zip(chain, chain[1:]):
            g.setdefault(u, set()).add(v)
            g.setdefault(v, set()).add(u)
            el[(min(u, v), max(u, v))] = math.hypot(v[0]-u[0], v[1]-u[1])
    return n, dists, side_stats


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

for hg in (False, "both", "any"):
    g, el = fresh()
    n, dists, ss = split_gated(g, el, None, head_gate=hg)
    hit, km, nc = main_net(g, el)
    tag = {False: "무게이트", "both": "양쪽헤드", "any": "한쪽헤드"}[hg]
    print(f"게이트 {tag:8s} 절단 {n:3d}  주망헤드 {hit:4d}  주망 {km:7.1f}m  조각 {nc}"
          + (f"  후보내역 {dict(ss)}" if hg else ""))
