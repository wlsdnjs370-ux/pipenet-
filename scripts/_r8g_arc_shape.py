# -*- coding: utf-8 -*-
"""r=180 ARC 쌍이 무슨 모양인가 — 교차점 주변 정확 좌표 덤프."""
from __future__ import annotations

import math
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = BASE / "data/uploads/B1F_.dxf"
TARGETS = [(615364.0, 117123.0), (662814.0, 175148.0)]
RAD = 700.0

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}


def near_entities(p, rad):
    out = []
    for en in bundle.entities:
        t = en["t"]
        if t == "L":
            q = en["p"]
            pts = [(q[0], q[1]), (q[2], q[3])]
        elif t in ("PL", "S", "H"):
            pts = [(v[0], v[1]) for v in en.get("p", [])]
        elif t in ("C", "A"):
            c = en.get("c") or [0, 0]
            pts = [(c[0], c[1])]
        else:
            q = en.get("p") or []
            pts = [(q[0], q[1])] if len(q) >= 2 else []
        if not pts:
            continue
        d = min(math.hypot(x - p[0], y - p[1]) for x, y in pts)
        if d <= rad:
            out.append((d, en))
    return sorted(out, key=lambda t: t[0])


for tgt in TARGETS:
    print(f"\n════ 교차점 ({tgt[0]:.0f}, {tgt[1]:.0f}) · {RAD:.0f}mm ════")
    for d, en in near_entities(tgt, RAD)[:26]:
        rel = lambda x, y: f"({x-tgt[0]:+8.1f},{y-tgt[1]:+8.1f})"  # noqa: E731
        t = en["t"]
        if t == "L":
            q = en["p"]
            geo = f"L {rel(q[0], q[1])}→{rel(q[2], q[3])}"
        elif t == "A":
            c = en["c"]; r = en["r"]; a0, a1 = en["a"]
            e0 = (c[0] + r*math.cos(math.radians(a0)), c[1] + r*math.sin(math.radians(a0)))
            e1 = (c[0] + r*math.cos(math.radians(a1)), c[1] + r*math.sin(math.radians(a1)))
            geo = (f"A c{rel(c[0], c[1])} r{r:.0f} {a0:.0f}~{a1:.0f} "
                   f"끝{rel(*e0)}→{rel(*e1)}")
        elif t == "C":
            geo = f"C c{rel(en['c'][0], en['c'][1])} r{en['r']:.0f}"
        elif t == "I":
            q = en["p"]
            geo = f"I {en.get('n','')} {rel(q[0], q[1])}"
        elif t in ("PL", "S", "H"):
            q = en.get("p", [])
            geo = f"{t} " + " ".join(rel(v[0], v[1]) for v in q[:4])
        else:
            geo = t
        print(f"  {d:7.1f} {geo:74s} {en['l']}[{lc.get(en['l'],'?')}]")
