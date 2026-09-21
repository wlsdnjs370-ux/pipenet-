# -*- coding: utf-8 -*-
"""B1F 제외 레이어 실측 — bbox · 원 반지름 · 현재망과의 끝점 일치."""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

CANDS = [
    ("최소", BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"),
    ("samples", BASE / "samples/dxf/B1F 현장조사 소화설비 평면도.dxf"),
    ("upload1", BASE / "data/uploads/B1F_.dxf"),
    ("upload2", BASE / "data/uploads/B1F__1.dxf"),
    ("upload3", BASE / "data/uploads/B1F___.dxf"),
]


def pts(en):
    p = en.get("p") or []
    if not p:
        return []
    if isinstance(p[0], (int, float)):
        return [(p[0], p[1])]
    return [(q[0], q[1]) for q in p]


for tag, path in CANDS:
    if not path.exists():
        print(f"[{tag}] 파일 없음: {path.name}")
        continue
    b = R.parse_dxf_bundle_cached(path)
    lc = {ly["name"]: ly["auto_category"] for ly in b.layers}
    by: dict = defaultdict(list)
    for en in b.entities:
        by[en.get("l")].append(en)
    print(f"\n=== [{tag}] {path.name} · 레이어 {len(b.layers)} · 엔티티 {len(b.entities)}")
    for name in sorted(by, key=lambda n: -len(by[n])):
        cat = lc.get(name, "(정의없음)")
        keep = cat in R.PIPENET_CATEGORIES or name in R.KEEP_BASE_LAYERS
        ents = by[name]
        tc = Counter(e.get("t") for e in ents)
        xs = [q[0] for e in ents for q in pts(e)]
        ys = [q[1] for e in ents for q in pts(e)]
        bb = (f"[{min(xs):.0f},{min(ys):.0f} ~ {max(xs):.0f},{max(ys):.0f}]"
              if xs else "[-]")
        radii = sorted({round(e.get("r") or 0, 1)
                        for e in ents if e.get("t") == "C"})
        print(f"  {'○' if keep else '·'} {str(name):34s} {str(cat):6s} "
              f"{len(ents):4d} {dict(tc)} {bb}" + (f" r={radii}" if radii else ""))
