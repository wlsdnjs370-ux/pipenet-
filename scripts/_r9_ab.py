# -*- coding: utf-8 -*-
"""R9(중복 선분 제거) on/off A-B — 4도면 최종 지오메트리 대조."""
from __future__ import annotations

import hashlib
import json
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXFS = [
    ("B1F", BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"),
    ("대명동", BASE / "samples/dxf/대명동201동 단위세대_layer정리.dxf"),
    ("LH306", BASE / "samples/dxf/LH306동_배관망.dxf"),
    ("LH지하", BASE / "samples/dxf/LH 지하층배관도_배관망.dxf"),
]
REAL = R._drop_covered_edges


def run(path: Path, on: bool) -> tuple:
    R._drop_covered_edges = REAL if on else (lambda *a, **k: 0)
    try:
        bundle = R.parse_dxf_bundle_cached(path)
        lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
        pe = R.filter_pipenet_only(bundle)
        heads = R._find_head_candidates(pe, lc)
        if not heads:
            return (0, 0, 0.0, "no-heads")
        xs = [h.pos[0] for h in heads]
        ys = [h.pos[1] for h in heads]
        zone = (min(xs) - 2000, min(ys) - 2000, max(xs) + 2000, max(ys) + 2000)
        res = R.select_worst30_heads_anchored(
            pe, lc, alarm_xy=(sum(xs) / len(xs), sum(ys) / len(ys)),
            head_region=R.HeadRegion(rects=[zone]), k=len(heads))
        blob = json.dumps(sorted((round(a[0], 3), round(a[1], 3),
                                  round(b[0], 3), round(b[1], 3), round(L, 3))
                                 for a, b, L in res.edges))
        return (len(res.heads), len(res.edges),
                round(sum(L for _a, _b, L in res.edges) / 1000, 1),
                hashlib.sha256(blob.encode()).hexdigest()[:12])
    finally:
        R._drop_covered_edges = REAL


for tag, path in DXFS:
    if not path.exists():
        print(f"[{tag}] 파일 없음")
        continue
    off = run(path, False)
    on = run(path, True)
    mark = "동일" if off == on else "변경"
    print(f"[{tag}] {mark}\n    off 헤드{off[0]:4d} 간선{off[1]:4d} {off[2]:7.1f}m {off[3]}"
          f"\n    on  헤드{on[0]:4d} 간선{on[1]:4d} {on[2]:7.1f}m {on[3]}")
