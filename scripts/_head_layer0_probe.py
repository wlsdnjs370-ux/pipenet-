# -*- coding: utf-8 -*-
"""미부착 헤드 곁의 레이어 '0' 선분이 진짜 가지배관 런인가, 헤드 기호 도형인가.

가장 가까운 '0' 선분이 속한 연결성분(끝점 공유)을 추적해 규모·연장·형태를 재고,
그 성분이 기존 PIPE 그래프와 닿는지(끝점 근접) 확인한다.
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

DXF = BASE / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf"

cap: dict = {}
_orig_final = R._finalize_selection


def _spy(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw):
    cap.update(graph=graph, heads=heads)
    return _orig_final(graph, edge_len, src, src_kind, heads, k, _pcb, *a, **kw)


R._finalize_selection = _spy

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
R.select_worst30_heads(pe, lc, k=30)
graph, heads = cap["graph"], cap["heads"]
miss = [h for h in heads if h.pos not in graph]
print(f"헤드 후보 {len(heads)} · 미부착 {len(miss)} · PIPE 그래프 노드 {len(graph)}")


def _seg_dist(p, a, b):
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 <= 0:
        return math.hypot(p[0] - a[0], p[1] - a[1])
    t = max(0.0, min(1.0, ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2))
    return math.hypot(p[0] - (a[0] + t * vx), p[1] - (a[1] + t * vy))


# 비-PIPE(=그래프에서 빠진) entity 를 성분 단위로 묶는다.
NON_PIPE = {"HEAD", "TEXT", "ALARM", "ARCH", "EXCLUDE"}
idx = R._NodeIndex()
ent_segs: dict[int, list] = {}
ent_layer: dict[int, str] = {}
at_node: dict[tuple, set[int]] = defaultdict(set)
for ei, en in enumerate(bundle.entities):
    ly = en.get("l", "")
    cat = lc.get(ly, "OTHER")
    if cat == "PIPE" or cat in NON_PIPE:
        continue
    t = en.get("t")
    if t == "L":
        p = en["p"]
        segs = [((p[0], p[1]), (p[2], p[3]))]
    elif t == "PL":
        pts = en["p"]
        segs = [((a[0], a[1]), (b[0], b[1])) for a, b in zip(pts, pts[1:])]
    else:
        continue
    if not segs:
        continue
    ent_segs[ei] = segs
    ent_layer[ei] = ly
    for a, b in segs:
        at_node[idx.canonical(*a)].add(ei)
        at_node[idx.canonical(*b)].add(ei)

print(f"비-PIPE 후보 entity {len(ent_segs)}개 (레이어 "
      f"{len(set(ent_layer.values()))}종)")

comp_of: dict[int, int] = {}
comps: list[list[int]] = []
for ei in ent_segs:
    if ei in comp_of:
        continue
    cid = len(comps)
    members = [ei]
    comp_of[ei] = cid
    stack = [ei]
    while stack:
        for a, b in ent_segs[stack.pop()]:
            for pt in (a, b):
                for nb in at_node[idx.canonical(*pt)]:
                    if nb not in comp_of:
                        comp_of[nb] = cid
                        members.append(nb)
                        stack.append(nb)
    comps.append(members)
print(f"연결성분 {len(comps)}개")


def _comp_stat(cid):
    ents = comps[cid]
    total = 0.0
    xs, ys = [], []
    for ei in ents:
        for a, b in ent_segs[ei]:
            total += math.hypot(b[0] - a[0], b[1] - a[1])
            xs += [a[0], b[0]]; ys += [a[1], b[1]]
    return {
        "n": len(ents),
        "len": total,
        "bbox": (max(xs) - min(xs), max(ys) - min(ys)),
        "layers": Counter(ent_layer[e] for e in ents),
    }


gnodes = list(graph)
gcells: dict[tuple, list] = defaultdict(list)
CELL = 500.0
for n in gnodes:
    gcells[(int(n[0] // CELL), int(n[1] // CELL))].append(n)


def _near_graph(pt, r=200.0):
    cx, cy = int(pt[0] // CELL), int(pt[1] // CELL)
    best = math.inf
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for n in gcells.get((cx + dx, cy + dy), ()):
                d = math.hypot(pt[0] - n[0], pt[1] - n[1])
                if d < best:
                    best = d
    return best


print("\n[미부착 헤드 → 최근접 비-PIPE 선분이 속한 성분]")
hit_comps = Counter()
rows = []
for h in miss:
    p = h.pos
    bd, bei = math.inf, None
    for ei, segs in ent_segs.items():
        for a, b in segs:
            d = _seg_dist(p, a, b)
            if d < bd:
                bd, bei = d, ei
    if bei is None:
        continue
    cid = comp_of[bei]
    hit_comps[cid] += 1
    rows.append((h, bd, cid))

print(f"  미부착 헤드가 가리키는 서로 다른 성분 {len(hit_comps)}개")
for cid, n in hit_comps.most_common(10):
    s = _comp_stat(cid)
    # 이 성분이 PIPE 그래프와 닿는가
    touch = math.inf
    for ei in comps[cid]:
        for a, b in ent_segs[ei]:
            for pt in (a, b):
                touch = min(touch, _near_graph(pt))
    ly = ", ".join(f"{k}×{v}" for k, v in s["layers"].most_common(3))
    print(f"  헤드 {n:>2}개 ← 성분#{cid}: entity {s['n']:>4} · 총연장 {s['len']/1000:>7.1f} m"
          f" · bbox {s['bbox'][0]/1000:.1f}×{s['bbox'][1]/1000:.1f} m"
          f" · PIPE망 최근접 {touch:.0f} mm · [{ly}]")
