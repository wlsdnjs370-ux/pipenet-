# -*- coding: utf-8 -*-
"""자동 추출이 «무엇을 뽑았는지» 그림으로 본다 — 말이 아니라 눈으로.

수치만으로는 「가지관이 빠졌다」를 확정할 수 없다. 뽑힌 망을 배경(전체 배관
그래프) 위에 겹쳐 그려 보면 한눈에 갈린다.

    회색   전체 배관 그래프 (도면에서 읽은 것 전부)
    빨강   검출한 헤드
    청록   뽑힌 최불리 배관망
    파랑   알람밸브

    python scripts/_probe_render_auto.py [도면.dxf] [--k 30] [--out x.png]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
GRAB: dict = {}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--out", default=str(ROOT / "data" / "_auto_render.png"))
    ap.add_argument("--zoom", action="store_true",
                    help="뽑힌 망 둘레만 확대")
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    import remote30_prototype as A
    from remote30_graph import HeadRegion

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    # 전체 배관 그래프를 파이프라인 한복판에서 붙잡는다.
    real = A._join_head_gap_endpoints

    def grab(graph, edge_len, head_pts, *args, **kw):
        n = real(graph, edge_len, head_pts, *args, **kw)
        GRAB["edges"] = list(edge_len.keys())
        return n

    A._join_head_gap_endpoints = grab
    try:
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = bundle.entities
        cat = {}
        for nm in {str(e.get("l") or "0") for e in ents}:
            try:
                cat[nm] = A._categorize_layer(nm)
            except Exception:  # noqa: BLE001
                cat[nm] = "OTHER"
        heads = A.detect_heads(ents, cat)
        pts = [(h.pos[0], h.pos[1]) for h in heads]
        sheet = A.sheet_frame_at(pts)
        inside = pts
        if sheet is not None:
            x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
            inside = [q for q in pts if x0 <= q[0] <= x1 and y0 <= q[1] <= y1]
        cx = sum(q[0] for q in inside) / len(inside)
        cy = sum(q[1] for q in inside) / len(inside)
        alarm = min(inside, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
        zones = A.head_bbox_for_region(pts, alarm)
        sel = A.select_worst30_heads_anchored(
            pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
            head_region=HeadRegion.from_rects(zones), zones=zones, k=a.k)
    finally:
        A._join_head_gap_endpoints = real

    bg = GRAB.get("edges") or []
    sel_ed = list(getattr(sel, "edges", None) or ())
    sel_hs = getattr(sel, "heads", None) or ()
    print(f"배경 간선 {len(bg):,} · 뽑힌 간선 {len(sel_ed):,} · "
          f"선정 헤드 {len(sel_hs)}")

    fig, ax = plt.subplots(figsize=(16, 10), dpi=110)
    ax.set_facecolor("#0a0d12")
    fig.patch.set_facecolor("#0a0d12")

    from matplotlib.collections import LineCollection
    if bg:
        ax.add_collection(LineCollection(
            [[(u[0], u[1]), (v[0], v[1])] for u, v in bg],
            colors="#4a5568", linewidths=0.35, zorder=1))
    ax.scatter([q[0] for q in inside], [q[1] for q in inside],
               s=1.2, c="#ff3b30", zorder=2, linewidths=0)
    if sel_ed:
        ax.add_collection(LineCollection(
            [[(e[0][0], e[0][1]), (e[1][0], e[1][1])] for e in sel_ed],
            colors="#22d3ee", linewidths=1.8, zorder=3))
    if sel_hs:
        ax.scatter([h.pos[0] for h in sel_hs], [h.pos[1] for h in sel_hs],
                   s=22, facecolors="none", edgecolors="#22d3ee",
                   linewidths=1.0, zorder=4)
    ax.plot([alarm[0]], [alarm[1]], marker="o", ms=11, mfc="none",
            mec="#3b82f6", mew=2.0, zorder=5)

    if a.zoom and sel_ed:
        xs = [c for e in sel_ed for c in (e[0][0], e[1][0])]
        ys = [c for e in sel_ed for c in (e[0][1], e[1][1])]
        pad = max(max(xs) - min(xs), max(ys) - min(ys)) * 0.12 + 2000
        ax.set_xlim(min(xs) - pad, max(xs) + pad)
        ax.set_ylim(min(ys) - pad, max(ys) + pad)
    else:
        ax.autoscale_view()
    ax.set_aspect("equal")
    ax.axis("off")
    fig.tight_layout(pad=0.2)
    fig.savefig(a.out, facecolor=fig.get_facecolor())
    print("그림:", a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
