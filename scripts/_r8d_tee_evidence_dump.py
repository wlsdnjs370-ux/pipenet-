# -*- coding: utf-8 -*-
"""실제 교차 절단 지점 주변에 무엇이 그려져 있나 — 원본 엔티티 통째 덤프.

번들(폭발된 엔티티)과 ezdxf 원본(중첩 INSERT 포함) 양쪽을 본다.
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
RAD = float(sys.argv[2]) if len(sys.argv) > 2 else 400.0

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
heads = R._find_head_candidates(pe, lc)
hpts = [h.pos for h in heads]

eps = R.auto_snap_eps(pe, lc)
g, el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps), layer_categories=lc)
R.collapse_parallel_ladders(g, el)
R._split_tee_branches(g, el)
R._drop_covered_edges(g, el)
R._join_head_gap_endpoints(g, el, hpts)
cuts: list = []
n = R._split_crossing_tees(g, el, head_pts=hpts, cuts_out=cuts)
pts = [tuple(c["p"]) for c in cuts]
print(f"절단 {n}건 · 번들 엔티티 {len(bundle.entities)}")

# ── 번들 엔티티 공간색인 ────────────────────────────────────────────────────
cell = 1000.0
grid: dict = defaultdict(list)


def ent_pts(en):
    t = en["t"]
    if t == "L":
        p = en["p"]
        return [(p[0], p[1]), (p[2], p[3])]
    if t in ("PL", "S", "H"):
        return [(q[0], q[1]) for q in en.get("p", [])]
    if t in ("C", "A"):
        c = en.get("c") or [0, 0]
        return [(c[0], c[1])]
    p = en.get("p") or []
    return [(p[0], p[1])] if len(p) >= 2 else []


for i, en in enumerate(bundle.entities):
    for x, y in ent_pts(en):
        grid[(int(x // cell), int(y // cell))].append(i)


def near(p, rad):
    span = int(rad // cell) + 1
    cx, cy = int(p[0] // cell), int(p[1] // cell)
    out = set()
    for dx in range(-span, span + 1):
        for dy in range(-span, span + 1):
            for i in grid.get((cx + dx, cy + dy), ()):
                for x, y in ent_pts(bundle.entities[i]):
                    if math.hypot(x - p[0], y - p[1]) <= rad:
                        out.add(i)
                        break
    return sorted(out)


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
        p = en["p"]
        return f"L len={math.hypot(p[2]-p[0], p[3]-p[1]):.0f}"
    if t == "T":
        return f"T {en.get('v','')[:20]!r}"
    if t in ("PL", "S", "H"):
        q = en.get("p", [])
        xs = [v[0] for v in q]; ys = [v[1] for v in q]
        return (f"{t} n={len(q)} bbox={max(xs)-min(xs):.0f}x{max(ys)-min(ys):.0f}"
                if q else f"{t} 빈")
    return t


# ── 절단점 주변 종합 통계 ──────────────────────────────────────────────────
kinds = Counter()
layers = Counter()
for p in pts:
    seen = set()
    for i in near(p, RAD):
        en = bundle.entities[i]
        key = (en["t"], en.get("n", ""))
        if key not in seen:
            seen.add(key)
            kinds[desc(en).split()[0] + (":" + en.get("n", "") if en["t"] == "I" else "")] += 1
        layers[f"{en['l']}[{lc.get(en['l'],'?')}]"] += 1
print(f"\n=== 절단점 {RAD:.0f}mm 이내 엔티티 종류 (절단점 수 기준) ===")
for k, v in kinds.most_common(20):
    print(f"  {k:24s} {v:5d} / {len(pts)}")
print(f"\n=== 레이어 (엔티티 수) ===")
for k, v in layers.most_common(15):
    print(f"  {k:44s} {v:6d}")

# ── 샘플 절단점 상세 ───────────────────────────────────────────────────────
print(f"\n=== 샘플 절단점 상세 (앞 6개) ===")
for p in pts[:6]:
    print(f"\n  절단점 ({p[0]:.0f}, {p[1]:.0f})")
    rows = []
    for i in near(p, RAD):
        en = bundle.entities[i]
        d = min(math.hypot(x - p[0], y - p[1]) for x, y in ent_pts(en))
        rows.append((d, en))
    for d, en in sorted(rows, key=lambda t: t[0])[:14]:
        print(f"      {d:7.1f}  {desc(en):32s} {en['l']}[{lc.get(en['l'],'?')}]")
