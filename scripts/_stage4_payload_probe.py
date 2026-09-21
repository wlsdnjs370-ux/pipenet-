# -*- coding: utf-8 -*-
"""실제 대명동 추출 결과로 build_stage4_entities 가 m/dm 을 제대로 싣는지 확인."""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as rp  # noqa: E402

FIXTURE = BASE / "samples/dxf/대명동201동 단위세대_layer정리.dxf"
WEST = [(244500.0, -243500.0), (253500.0, -243500.0),
        (253500.0, -221500.0), (244500.0, -221500.0)]

bundle = rp.parse_dxf_bundle(FIXTURE)
layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
ents_in = rp.filter_pipenet_only(bundle)
region = rp.HeadRegion.from_polygon(WEST)
gated = rp.detect_heads(ents_in, layer_cat, region=region)
cx = sum(h.pos[0] for h in gated) / len(gated)
cy = sum(h.pos[1] for h in gated) / len(gated)
sel = rp.select_worst30_heads_anchored(
    ents_in, layer_cat, alarm_xy=(cx, cy), head_region=region)

ents = rp.build_stage4_entities(sel)
edges = [e for e in ents if e["l"] == "_subgraph"]
heads = [e for e in ents if e["l"] == "_subgraph_head"]
av = [e for e in ents if e["l"] == "_alarm_valve"]

print(f"edge {len(edges)} · head {len(heads)} · AV {len(av)}")
print(f"m 누락 edge   : {sum(1 for e in edges if 'm' not in e)}")
print(f"dm 누락 head  : {sum(1 for e in heads if 'dm' not in e)}")
print(f"edge 실길이 합 : {sum(e['m'] for e in edges) / 1000:.1f} m")
print(f"헤드 AV거리   : 최소 {min(h['dm'] for h in heads)/1000:.1f} m"
      f" · 최대 {max(h['dm'] for h in heads)/1000:.1f} m")
print(f"selection 최대 : {max(sel.distances)/1000:.1f} m  (일치해야 함)")

# 프론트가 하는 누적계산을 그대로 재현 — 트리 밖 말고는 전부 이어져야 한다.
def key(x, y):
    return (round(x * 10), round(y * 10))

at = {key(av[0]["c"][0], av[0]["c"][1]): 0.0}
gap = 0
for e in edges:
    p = e["p"]
    up = at.get(key(p[0], p[1]))
    if up is None:
        gap += 1
        continue
    d = up + e["m"]
    at.setdefault(key(p[2], p[3]), d)
print(f"누적 못 잇는 edge: {gap} / {len(edges)}"
      f"  (트리 밖 {sum(1 for e in edges if e.get('x'))}건 이내여야 함)")

ok = (not any("m" not in e for e in edges) and not any("dm" not in h for h in heads)
      and gap <= sum(1 for e in edges if e.get("x")))
print("\n" + ("PASS" if ok else "FAIL"))
sys.exit(0 if ok else 1)
