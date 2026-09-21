# -*- coding: utf-8 -*-
"""레이어별 '헤드 틈 지문' 계수 — 이름 아닌 지오메트리로 배관 레이어 판별 가능한가."""
from __future__ import annotations

import math
import sys
from collections import defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXFS = [
    ("B1F_upload", BASE / "data/uploads/B1F_.dxf"),
    ("B1F_최소", BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"),
    ("대명동", BASE / "samples/dxf/대명동201동 단위세대_layer정리.dxf"),
    ("LH306", BASE / "samples/dxf/LH306동_배관망.dxf"),
    ("LH지하", BASE / "samples/dxf/LH 지하층배관도_배관망.dxf"),
]
AX_TOL = R.CROSS_TEE_AXIS_TOL_MM
GAP_MAX = R.HEAD_GAP_JOIN_MAX_MM
GAP_TOL = R.HEAD_GAP_JOIN_TOL_MM
MIN_SEG = R.MIN_PIPE_EDGE_MM


def segs(en):
    p = en.get("p") or []
    if not p:
        return []
    if isinstance(p[0], (int, float)):
        return [((p[0], p[1]), (p[2], p[3]))] if len(p) >= 4 else []
    q = [(v[0], v[1]) for v in p]
    return list(zip(q, q[1:]))


def signature(seglist, heads) -> int:
    """같은 선 위 두 조각이 헤드 하나를 사이에 두고 벌어진 횟수."""
    lanes: dict = defaultdict(list)
    for a, b in seglist:
        if math.hypot(b[0]-a[0], b[1]-a[1]) < MIN_SEG:
            continue
        ax = R._axis_index(a, b, AX_TOL)
        if ax < 0:
            continue
        run = 1 if ax == 0 else 0
        off = (a[1-run] + b[1-run]) * 0.5
        lanes[(ax, round(off / AX_TOL))].append(
            (min(a[run], b[run]), max(a[run], b[run]), off))
    hlane: dict = defaultdict(list)
    for hp in heads:
        for ax in (0, 1):
            run = 1 if ax == 0 else 0
            hlane[(ax, round(hp[1-run] / GAP_TOL))].append(hp[run])
    n = 0
    for (ax, r), items in lanes.items():
        items.sort()
        run = 1 if ax == 0 else 0
        o = items[0][2]
        hr = round(o / GAP_TOL)
        hs = sorted(h for dr in (-1, 0, 1)
                    for h in hlane.get((ax, hr + dr), []))
        for (lo1, hi1, _o1), (lo2, hi2, _o2) in zip(items, items[1:]):
            gap = lo2 - hi1
            if not (AX_TOL < gap <= GAP_MAX):
                continue
            mid = (hi1 + lo2) * 0.5
            if any(abs(h - mid) <= GAP_TOL for h in hs):
                n += 1
    return n


for tag, path in DXFS:
    if not path.exists():
        print(f"[{tag}] 파일 없음")
        continue
    bundle = R.parse_dxf_bundle_cached(path)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    heads = [h.pos for h in R._find_head_candidates(R.filter_pipenet_only(bundle), lc)]
    by: dict = defaultdict(list)
    for en in bundle.entities:
        by[en.get("l")].extend(segs(en))
    print(f"\n=== [{tag}] 헤드 {len(heads)}개")
    rows = []
    for name, sl in by.items():
        cat = lc.get(name, "OTHER")
        n = signature(sl, heads)
        if n:
            rows.append((n, name, cat, len(sl)))
    for n, name, cat, ns in sorted(rows, reverse=True):
        keep = cat in R.PIPENET_CATEGORIES or name in R.KEEP_BASE_LAYERS
        print(f"  {'○' if keep else '·'} {str(name):34s} {str(cat):8s} "
              f"지문{n:5d} / 세그{ns:6d}")
