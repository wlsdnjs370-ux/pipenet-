# -*- coding: utf-8 -*-
"""승격 레이어 세그먼트 원좌표 덤프 — 무엇이 배관이고 무엇이 아닌가."""
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
LAYER = sys.argv[1] if len(sys.argv) > 1 else "현장조사#셔터"

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
heads = [h.pos for h in R._find_head_candidates(R.filter_pipenet_only(bundle), lc)]

rows = []
for ei, en in enumerate(bundle.entities):
    if en.get("l") != LAYER:
        continue
    p = en.get("p") or []
    if en["t"] == "I":
        print(f"  INSERT  n={en.get('n')!r}  ({p[0]:9.0f},{p[1]:9.0f})")
        continue
    sl = [((p[0], p[1]), (p[2], p[3]))] if en["t"] == "L" and len(p) >= 4 else \
         list(zip(p, p[1:])) if en["t"] == "PL" else []
    for a, b in sl:
        rows.append((a, b, math.hypot(b[0]-a[0], b[1]-a[1]), en["t"], ei))

print(f"\n[{LAYER}] 세그 {len(rows)}개 — y 오름차순")
for a, b, L, t, ei in sorted(rows, key=lambda r: ((r[0][1]+r[1][1])*0.5, r[0][0])):
    ax = R._axis_index(a, b, R.CROSS_TEE_AXIS_TOL_MM)
    tag = {0: "세로", 1: "가로"}.get(ax, "사선")
    near = sum(1 for h in heads
               if min(a[0], b[0])-200 <= h[0] <= max(a[0], b[0])+200
               and min(a[1], b[1])-200 <= h[1] <= max(a[1], b[1])+200)
    print(f"  {tag} {L:8.0f}mm  ({a[0]:9.0f},{a[1]:9.0f})→({b[0]:9.0f},{b[1]:9.0f})"
          f"  {t}#{ei}  런위헤드 {near}")
