# -*- coding: utf-8 -*-
"""B1F 원본 — 켠/뺀 PIPE 레이어의 기하 관계 진단.

참조(정답)가 유지한 레이어 vs 삭제한 PIPE 레이어가 공간적으로 어떤 관계인지
(중복? 별개 영역? 별 계통?) 파악해 일반화 규칙 후보를 찾는다.

각 레이어: 엔티티수, bbox, 라인 좌표 샘플, 참조망 bbox 와의 겹침.
ASCII-only.
"""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
import remote30_prototype as R  # noqa: E402

ORIG = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"
REF = BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"

FOCUS_LAYERS = [
    "6-소화-가지관", "6-소화-가지관(측벽)",
    "6-소화-SP-메인", "6-소화-SP메인(B1F)", "6-소화-SP-메인(1차)",
    "6-소화-H-메인", "6-소화-H메인(B1F)",
    "6-소화-밸브", "1-PIT-소화",
    "6-소화-헤드-상향식(72)", "6-소화-헤드-하향식(DRY)", "6-소화-헤드-측벽형",
]


def _pair(v):
    """v 가 (x,y) 스칼라쌍이면 (float,float), 아니면 None."""
    try:
        if len(v) >= 2 and not hasattr(v[0], "__len__") and not hasattr(v[1], "__len__"):
            return (float(v[0]), float(v[1]))
    except (TypeError, ValueError):
        pass
    return None


def coords_of(ent):
    t = ent.get("t")
    if t == "L":
        p = ent.get("p")
        if p and len(p) >= 4:
            return [(float(p[0]), float(p[1])), (float(p[2]), float(p[3]))]
        return []
    if t in ("PL", "S"):
        out = []
        for pt in ent.get("p", []) or []:
            pr = _pair(pt)
            if pr:
                out.append(pr)
        return out
    if t in ("A", "C"):
        pr = _pair(ent.get("c", []))
        return [pr] if pr else []
    if t in ("I", "T", "H"):
        pr = _pair(ent.get("p", []))
        return [pr] if pr else []
    return []


def bbox_of(pts):
    if not pts:
        return None
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def layer_pts(bundle, layer):
    pts = []
    for e in bundle.entities:
        if e.get("l") == layer:
            pts.extend(coords_of(e))
    return pts


def bbox_overlap(a, b):
    if not a or not b:
        return 0.0
    ox = max(0.0, min(a[2], b[2]) - max(a[0], b[0]))
    oy = max(0.0, min(a[3], b[3]) - max(a[1], b[1]))
    area = ox * oy
    aa = (a[2] - a[0]) * (a[3] - a[1])
    return area / aa if aa > 0 else 0.0


def main():
    print("Parsing ORIG...", flush=True)
    bo = R.parse_dxf_bundle_cached(ORIG)
    print("Parsing REF...", flush=True)
    br = R.parse_dxf_bundle_cached(REF)

    ref_pts = []
    for e in br.entities:
        if e.get("l", "").startswith("6-소화") or e.get("l") == "1-PIT-소화":
            ref_pts.extend(coords_of(e))
    ref_bb = bbox_of(ref_pts)
    print(f"\nREF network bbox: {tuple(round(v) for v in ref_bb) if ref_bb else None}", flush=True)
    print(f"REF bbox size: {round(ref_bb[2]-ref_bb[0])} x {round(ref_bb[3]-ref_bb[1])} mm\n", flush=True)

    print(f"{'layer':32s} {'n':>6s} {'bbox(x0,y0,x1,y1)':>40s} {'size':>20s} {'ovl_ref':>8s}", flush=True)
    for ly in FOCUS_LAYERS:
        pts = layer_pts(bo, ly)
        bb = bbox_of(pts)
        if not bb:
            print(f"{ly:32s} {'0':>6s}  (absent)", flush=True)
            continue
        sz = f"{round(bb[2]-bb[0])}x{round(bb[3]-bb[1])}"
        ov = bbox_overlap(bb, ref_bb)
        bbs = f"({round(bb[0])},{round(bb[1])},{round(bb[2])},{round(bb[3])})"
        print(f"{ly:32s} {len(pts):>6d} {bbs:>40s} {sz:>20s} {ov:>7.2f}", flush=True)

    # 참조가 유지한 가지관 vs 원본 가지관: 참조 가지관 좌표가 원본 어디에 대응?
    print("\n-- REF 가지관 좌표 범위 (worst-head 영역 위치) --", flush=True)
    rg = layer_pts(br, "6-소화-가지관")
    rgb = bbox_of(rg)
    print(f"REF 가지관 bbox: ({round(rgb[0])},{round(rgb[1])},{round(rgb[2])},{round(rgb[3])}) size {round(rgb[2]-rgb[0])}x{round(rgb[3]-rgb[1])}", flush=True)
    og = layer_pts(bo, "6-소화-가지관")
    ogb = bbox_of(og)
    print(f"ORIG 가지관 bbox: ({round(ogb[0])},{round(ogb[1])},{round(ogb[2])},{round(ogb[3])}) size {round(ogb[2]-ogb[0])}x{round(ogb[3]-ogb[1])}", flush=True)
    print(f"REF 가지관 / ORIG 가지관 면적비: {((rgb[2]-rgb[0])*(rgb[3]-rgb[1])) / max(1e-9,(ogb[2]-ogb[0])*(ogb[3]-ogb[1])):.4f}", flush=True)
    print("\nDONE", flush=True)


if __name__ == "__main__":
    main()
