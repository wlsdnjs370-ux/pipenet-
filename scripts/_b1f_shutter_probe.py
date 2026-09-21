# -*- coding: utf-8 -*-
"""제외 레이어 정체 — 선분 좌표 · 길이 · 현재망 최근접 거리."""
from __future__ import annotations

import math
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"
TARGETS = set(sys.argv[2:]) or {"현장조사#셔터"}

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
heads = R._find_head_candidates(pe, lc)
eps = R.auto_snap_eps(pe, lc)
graph, _el = R._build_graph(pe, node_index=R._NodeIndex(epsilon_mm=eps),
                            layer_categories=lc)
nodes = list(graph)
gx = [n[0] for n in nodes]
gy = [n[1] for n in nodes]
print(f"{DXF.name}\n망 노드 {len(nodes)} eps={eps:.1f} 헤드 {len(heads)} "
      f"bbox [{min(gx):.0f},{min(gy):.0f} ~ {max(gx):.0f},{max(gy):.0f}]\n")


def segs(en):
    p = en.get("p") or []
    t = en.get("t")
    if not p:
        return []
    if isinstance(p[0], (int, float)):
        if t in ("L",) and len(p) >= 4:
            return [((p[0], p[1]), (p[2], p[3]))]
        return []
    q = [(v[0], v[1]) for v in p]
    return list(zip(q, q[1:]))


for name in sorted(TARGETS):
    ents = [e for e in bundle.entities if e.get("l") == name]
    print(f"── {name} · {len(ents)}개 · cat={lc.get(name)}")
    tot = 0.0
    for en in ents:
        ss = segs(en)
        if not ss:
            print(f"    {en.get('t')} n={en.get('n')} r={en.get('r')} p={en.get('p')}")
            continue
        for a, b in ss:
            L = math.hypot(b[0]-a[0], b[1]-a[1])
            tot += L
            if L < 1.0:
                continue
            da = min(math.hypot(n[0]-a[0], n[1]-a[1]) for n in nodes)
            db = min(math.hypot(n[0]-b[0], n[1]-b[1]) for n in nodes)
            print(f"    {en.get('t')} ({a[0]:10.0f},{a[1]:10.0f})-"
                  f"({b[0]:10.0f},{b[1]:10.0f}) L={L:9.1f} | 망 a={da:8.1f} b={db:8.1f}")
    print(f"    합계 연장 {tot/1000:.1f}m\n")
