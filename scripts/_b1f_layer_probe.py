# -*- coding: utf-8 -*-
"""B1F 레이어 인벤토리 — 이름 · 자동분류 · 엔티티 타입 분포."""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXF = BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"

bundle = R.parse_dxf_bundle_cached(DXF)
lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}

by_layer: dict = defaultdict(Counter)
for en in bundle.entities:
    by_layer[en.get("l")][en.get("t")] += 1

print(f"레이어 {len(bundle.layers)}개 · 엔티티 {len(bundle.entities)}개")
print(f"PIPENET_CATEGORIES = {R.PIPENET_CATEGORIES}")
print(f"KEEP_BASE_LAYERS   = {R.KEEP_BASE_LAYERS}\n")
print(f"{'레이어':38s} {'분류':10s} {'수':>5s}  타입분포")
for name in sorted(by_layer, key=lambda n: -sum(by_layer[n].values())):
    cat = lc.get(name, "(레이어정의없음)")
    tot = sum(by_layer[name].values())
    keep = "○" if (cat in R.PIPENET_CATEGORIES or name in R.KEEP_BASE_LAYERS) else "·"
    print(f"{keep} {str(name):36s} {str(cat):10s} {tot:5d}  {dict(by_layer[name])}")
