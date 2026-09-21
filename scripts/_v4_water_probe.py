# -*- coding: utf-8 -*-
"""V4 급수 감사 실측 — 대명동 서측 세대 anchored 추출의 water audit 수치·비용 보고."""
from __future__ import annotations

import json
import sys
import time
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import remote30_prototype as rp  # noqa: E402

DXF = REPO / "samples/dxf/대명동201동 단위세대_layer정리.dxf"
WEST_UNIT_POLY = [(244500.0, -243500.0), (253500.0, -243500.0),
                  (253500.0, -221500.0), (244500.0, -221500.0)]

bundle = rp.parse_dxf_bundle(DXF)
ents = rp.filter_pipenet_only(bundle)
cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
region = rp.HeadRegion.from_polygon(WEST_UNIT_POLY)
gated = rp.detect_heads(ents, cats, region=region)
cx = sum(h.pos[0] for h in gated) / len(gated)
cy = sum(h.pos[1] for h in gated) / len(gated)


def run():
    t0 = time.perf_counter()
    r = rp.select_worst30_heads_anchored(ents, cats, alarm_xy=(cx, cy),
                                         head_region=region)
    return r, time.perf_counter() - t0


res, t_with = run()
j = res.audit.to_json_dict()
print(f"헤드(region 승인): {j['heads']['detected_in_region']}  "
      f"부착 {j['heads']['attached']}  미도달 {len(j['heads']['unreachable'])}")
print(f"최종 배출 edge: {len(res.edges)}")
print("water:", json.dumps(j["water"], ensure_ascii=False, indent=2))

_real = rp.water_load_audit
rp.water_load_audit = lambda *a, **k: {}
_, t_without = run()
rp.water_load_audit = _real
print(f"\n소요: 감사 포함 {t_with:.3f}s / 미포함 {t_without:.3f}s "
      f"(증분 {t_with - t_without:+.3f}s)")
