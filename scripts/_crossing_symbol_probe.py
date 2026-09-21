# -*- coding: utf-8 -*-
"""관통 교차(끝점 없는 X 교차)에 접속 기호가 있는가 — B1F vs 대명동 실측."""
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

DXFS = [
    ("B1F", BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"),
    ("대명동", BASE / "samples/dxf/대명동201동 단위세대_layer정리.dxf"),
]
EPS = 1.0
MERGE_ONLY = "--merge-only" in sys.argv
MINLEN = next((float(a.split("=")[1]) for a in sys.argv if a.startswith("--minlen=")), 0.0)
NOLONE = "--no-lone" in sys.argv


def _crossings(edge_len: dict, graph: dict):
    """축평행 edge 간 내부×내부 직교 교차 — [(pt, hkey, vkey)]."""
    hs, vs = [], []
    for (a, b) in edge_len:
        if abs(a[1] - b[1]) <= EPS and abs(a[0] - b[0]) > EPS:
            hs.append(((min(a[0], b[0]), max(a[0], b[0])), a[1], (a, b)))
        elif abs(a[0] - b[0]) <= EPS and abs(a[1] - b[1]) > EPS:
            vs.append(((min(a[1], b[1]), max(a[1], b[1])), a[0], (a, b)))
    out = []
    for (xr, y, hk) in hs:
        for (yr, x, vk) in vs:
            if not (xr[0] + EPS < x < xr[1] - EPS):
                continue
            if not (yr[0] + EPS < y < yr[1] - EPS):
                continue
            out.append(((x, y), hk, vk))
    return out


for tag, path in DXFS:
    bundle = R.parse_dxf_bundle_cached(path)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pe = R.filter_pipenet_only(bundle)
    heads = R._find_head_candidates(pe, lc)
    hpts = [h.pos for h in heads]
    eps = R.auto_snap_eps(pe, lc)
    graph, edge_len = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                                     layer_categories=lc)
    R.collapse_parallel_ladders(graph, edge_len)
    R._split_tee_branches(graph, edge_len)
    R._join_head_gap_endpoints(graph, edge_len, hpts)

    def _stat(tag2: str) -> None:
        comps_ = R._connected_components(graph)
        cmap = {n: i for i, c in enumerate(comps_) for n in c}
        clen: dict = defaultdict(float)
        for (a_, b_), L in edge_len.items():
            clen[cmap[a_]] += L
        hc_: Counter = Counter()
        for hp in hpts:
            nn = R._nearest_graph_node(graph, hp)
            if nn is not None and math.hypot(nn[0] - hp[0], nn[1] - hp[1]) <= R.HEAD_DROP_MAX_MM:
                hc_[cmap[nn]] += 1
        best = hc_.most_common(1)
        cid = best[0][0] if best else None
        cyc = len(edge_len) - len(graph) + len(comps_)
        print(f"    [{tag2}] 노드 {len(graph)} 간선 {len(edge_len)} 조각 {len(comps_)} "
              f"루프 {cyc} — 최대 헤드보유 조각 {best[0][1] if best else 0}헤드/"
              f"{round(clen.get(cid, 0.0) / 1000, 1)}m / 전체 헤드 {len(hpts)}")

    xs = _crossings(edge_len, graph)
    comps = R._connected_components(graph)
    comp_of = {n: i for i, c in enumerate(comps) for n in c}
    cross_comp = sum(1 for p, hk, vk in xs if comp_of[hk[0]] != comp_of[vk[0]])

    syms = [(en["p"][0], en["p"][1], en.get("n") or "", en.get("l") or "")
            for en in pe if en["t"] == "I"]
    print(f"\n=== {tag} — 노드 {len(graph)} 간선 {len(edge_len)} 조각 {len(comps)} "
          f"기호(INSERT) {len(syms)}")
    print(f"    내부×내부 직교 교차 {len(xs)}곳 (그중 서로 다른 조각을 잇는 것 {cross_comp}곳)")
    dists = []
    for p, hk, vk in xs:
        d = min((math.hypot(p[0] - sx, p[1] - sy) for sx, sy, _n, _l in syms),
                default=None)
        dh = min((math.hypot(p[0] - hx, p[1] - hy) for hx, hy in hpts), default=None)
        dists.append((d, dh))
    if dists:
        print("    교차점→최근접 기호 거리 분포:",
              Counter(("없음" if d is None else
                       ("≤100mm" if d <= 100 else "≤300mm" if d <= 300 else
                        "≤1000mm" if d <= 1000 else ">1000mm"))
                      for d, _ in dists).most_common())
        print("    교차점→최근접 헤드 거리 분포:",
              Counter(("없음" if d is None else
                       ("≤100mm" if d <= 100 else "≤300mm" if d <= 300 else
                        "≤1000mm" if d <= 1000 else ">1000mm"))
                      for _, d in dists).most_common())

    _stat("절단 전")

    # 합류하는 교차만 골라낸다 — 절단해도 루프가 안 생기는 것(union-find).
    parent: dict = {}

    def _find(x):
        while parent.get(x, x) != x:
            parent[x] = parent.get(parent[x], parent[x])
            x = parent[x]
        return x

    for a_, b_ in edge_len:
        ra, rb = _find(a_), _find(b_)
        if ra != rb:
            parent[ra] = rb
    merging = []
    for p, hk, vk in sorted(xs):
        if MINLEN and (edge_len.get(hk, 0.0) < MINLEN or edge_len.get(vk, 0.0) < MINLEN):
            continue
        if NOLONE and (len(comps[comp_of[hk[0]]]) < 3 or len(comps[comp_of[vk[0]]]) < 3):
            continue
        ra, rb = _find(hk[0]), _find(vk[0])
        if ra == rb:
            continue
        parent[ra] = rb
        merging.append((p, hk, vk))
    print(f"    합류 교차(루프 안 만드는 절단) {len(merging)}곳 / 전체 {len(xs)}곳")

    # 전부 절단해 붙이면 어떻게 되나 — 상한 없는 최대 효과 측정
    # 한 간선이 여러 번 교차될 수 있으므로 간선별로 교차점을 모아 한꺼번에 분할한다.
    cuts: dict = defaultdict(set)
    for p, hk, vk in (merging if MERGE_ONLY else xs):
        cuts[hk].add(p)
        cuts[vk].add(p)
    for key, pts in cuts.items():
        a, b = key
        graph[a].discard(b)
        graph[b].discard(a)
        edge_len.pop(key, None)
        chain = [a] + sorted(pts, key=lambda q: (q[0] - a[0]) ** 2 + (q[1] - a[1]) ** 2) + [b]
        for u, v in zip(chain, chain[1:]):
            if u == v:
                continue
            graph.setdefault(u, set()).add(v)
            graph.setdefault(v, set()).add(u)
            edge_len[(min(u, v), max(u, v))] = math.hypot(v[0] - u[0], v[1] - u[1])
    _stat("합류만 절단 후" if MERGE_ONLY else "전부 절단 후")
