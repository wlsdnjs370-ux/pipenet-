# -*- coding: utf-8 -*-
"""그 15개 헤드는 «어느 선» 에 달렸나 — 가지관인가 셔터선인가.

`현장조사#셔터` 는 두 덩이다:
    (배관 승격)  18선분 · 43.4 m · 헤드 15   ← R10b 가 올린 것
    (그대로)      3선분 · 42.2 m · 헤드 15   ← OTHER 로 남은 것

둘 다 «헤드 15» 로 나오는데, 나란히 263 mm 떨어져 지나가기 때문이다(BLOCKED
§14 R10b). 헤드 결합 상한이 300 mm 라 같은 헤드가 양쪽에 다 걸린다.

그래서 «가까운 쪽» 으로 가른다 — 헤드마다 두 선까지 거리를 재서 어느 쪽이
주인인지 본다. 그리고 그 자리를 그림으로 남긴다.

    python scripts/_probe_shutter_which.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def _segs(e):
    t = str(e.get("t") or "")
    if t not in ("L", "PL"):
        return []
    p = e.get("p") or []
    if not p:
        return []
    if isinstance(p[0], (list, tuple)):
        pts = [(float(q[0]), float(q[1])) for q in p if len(q) >= 2]
    else:
        if len(p) < 4:
            return []
        pts = [(float(p[i]), float(p[i + 1])) for i in range(0, len(p) - 1, 2)]
    return [(pts[i], pts[i + 1]) for i in range(len(pts) - 1)]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--out", default=str(ROOT / "data" / "_shutter_zoom.png"))
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.collections import LineCollection

    import remote30_prototype as A
    from remote30_graph import _point_to_segment_dist

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    cat = {}
    for nm in {str(e.get("l") or "0") for e in ents}:
        try:
            cat[nm] = A._categorize_layer(nm)
        except Exception:  # noqa: BLE001
            cat[nm] = "OTHER"
    heads = A.detect_heads(ents, cat)

    promoted, raw = [], []
    for e in ents:
        ly = str(e.get("l") or "0")
        if "셔터" not in ly:
            continue
        (promoted if "승격" in ly else raw).extend(_segs(e))
    print(f"승격분 {len(promoted)}선분 · 미승격 {len(raw)}선분\n")
    if not promoted or not raw:
        print("두 덩이가 다 있어야 비교가 된다.")
        return 1

    def dmin(hp, segs):
        return min(_point_to_segment_dist(hp[0], hp[1], p[0], p[1], q[0], q[1])
                   for p, q in segs)

    # 두 덩이 어느 쪽이든 300mm 안에 드는 헤드만 본다.
    DROP = A.HEAD_DROP_MAX_MM
    near = []
    for h in heads:
        hp = (h.pos[0], h.pos[1])
        dp, dr = dmin(hp, promoted), dmin(hp, raw)
        if min(dp, dr) <= DROP:
            near.append((hp, dp, dr))
    print(f"두 선 중 하나에 {DROP:.0f}mm 안으로 붙는 헤드 {len(near)}개\n")
    print(f"  {'헤드 좌표':>26} {'승격분까지':>10} {'셔터선까지':>10}  주인")
    print("  " + "-" * 62)
    win_p = win_r = 0
    for hp, dp, dr in sorted(near, key=lambda t: t[0][0])[:20]:
        who = "가지관(승격분)" if dp < dr else "셔터선"
        if dp < dr:
            win_p += 1
        else:
            win_r += 1
        print(f"  ({hp[0]:>10.0f},{hp[1]:>9.0f}) {dp:>10.0f} {dr:>10.0f}  {who}")
    for hp, dp, dr in sorted(near, key=lambda t: t[0][0])[20:]:
        if dp < dr:
            win_p += 1
        else:
            win_r += 1
    print(f"\n  가지관(승격분)이 더 가까운 헤드 {win_p}개")
    print(f"  셔터선이 더 가까운 헤드      {win_r}개")
    if win_r == 0:
        print("\n  ★헤드는 전부 승격분(가지관) 쪽에 달렸다 —")
        print("    셔터선은 그 옆을 나란히 지나갈 뿐이다.")

    # 그림 — 눈으로도 확인한다.
    xs = [c for p, q in promoted + raw for c in (p[0], q[0])]
    ys = [c for p, q in promoted + raw for c in (p[1], q[1])]
    fig, ax = plt.subplots(figsize=(17, 5.5), dpi=110)
    ax.set_facecolor("#0a0d12")
    fig.patch.set_facecolor("#0a0d12")
    ax.add_collection(LineCollection(
        [[(p[0], p[1]), (q[0], q[1])] for p, q in promoted],
        colors="#22d3ee", linewidths=2.2))
    ax.add_collection(LineCollection(
        [[(p[0], p[1]), (q[0], q[1])] for p, q in raw],
        colors="#e879f9", linewidths=2.2))
    ax.scatter([hp[0] for hp, _, _ in near], [hp[1] for hp, _, _ in near],
               s=26, c="#ff3b30", zorder=5, linewidths=0)
    pad = 1500
    ax.set_xlim(min(xs) - pad, max(xs) + pad)
    ax.set_ylim(min(ys) - pad, max(ys) + pad)
    ax.set_aspect("equal")
    ax.axis("off")
    ax.set_title("청록=배관 승격분  분홍=셔터선(미승격)  빨강=헤드",
                 color="#cbd5e1", fontsize=11)
    fig.tight_layout(pad=0.3)
    fig.savefig(a.out, facecolor=fig.get_facecolor())
    print("\n그림:", a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
