# -*- coding: utf-8 -*-
"""B1F 알람밸브 결합 실측 — 소스 결합 완화 전후."""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"
SRC = (156776.0, 177970.0)

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
pe = R.filter_pipenet_only(bundle)
heads = R._find_head_candidates(pe, lc)
xs = [h.pos[0] for h in heads]
ys = [h.pos[1] for h in heads]
zone = (min(xs) - 2000, min(ys) - 2000, max(xs) + 2000, max(ys) + 2000)

audit: dict = {}
res = R.select_worst30_heads_anchored(
    pe, lc, alarm_xy=SRC, head_region=R.HeadRegion(rects=[zone]),
    k=115, audit_out=audit)
total = sum(L for _a, _b, L in res.edges)
print(f"[anchored] 헤드 {len(res.heads)}  간선 {len(res.edges)}  연장 {total:.0f}mm  "
      f"src={tuple(round(v) for v in res.source_pos)}")
print("  head_gap_joins:", len(audit.get("head_gap_joins") or []))
print("  crossing_tees :", len(audit.get("crossing_tees") or []))
print("  source_attach:", audit.get("source_attach"))
print("  fragments    :", audit.get("fragments"))
h = dict(audit.get("heads") or {})
h["unreachable"] = len(h.get("unreachable") or [])
print("  heads        :", h)
