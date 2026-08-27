# -*- coding: utf-8 -*-
"""배관 entity 가 그래프까지 오면서 «어디서» 사라지는가.

B1F 는 배관 레이어 entity 가 16,500개인데 그래프 간선은 4,090개다. 4분의 1도
안 남는다. 「가지관을 못 잡는다」의 실체가 여기일 수 있다 — 틈이 아니라 손실.

단계별로 센다:
    entity      레이어 분류가 PIPE 인 것
    segment     그 entity 를 선분으로 편 것
    window      앵커 작업창(W) 안에 든 것        ← anchored 경로만 있는 관문
    graph       실제로 그래프에 남은 간선

    python scripts/_probe_pipe_loss.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import math
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
GRAB: dict = {}


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
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import remote30_prototype as A
    from remote30_graph import HeadRegion

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

    # ── ① 레이어 분류
    by_cat = Counter(cat.get(str(e.get("l") or "0"), "OTHER") for e in ents)
    print(f"{dxf.name}\n")
    print("■ entity 를 레이어 분류로 나누면")
    for k, v in by_cat.most_common():
        print(f"    {k:<10} {v:>8,}")
    ex = sorted(n for n, c in cat.items() if c == "EXCLUDE")
    if ex:
        print(f"  EXCLUDE 레이어: {', '.join(ex)}")

    # ── ② 배관 entity → 선분
    pipe_ents = [e for e in ents
                 if cat.get(str(e.get("l") or "0")) == "PIPE"]
    seg_by_layer = Counter()
    segs = []
    for e in pipe_ents:
        s = _segs(e)
        seg_by_layer[str(e.get("l") or "0")] += len(s)
        segs.extend(s)
    print(f"\n■ 배관 entity {len(pipe_ents):,} → 선분 {len(segs):,}")
    for k, v in seg_by_layer.most_common():
        print(f"    {v:>8,}  {k}")

    # 길이 분포 — 짧은 선분 컷(MIN_PIPE_EDGE_MM)이 얼마나 먹는지 본다.
    ls = sorted(math.hypot(q[0] - p[0], q[1] - p[1]) for p, q in segs)
    if ls:
        print(f"\n  선분 길이 — 중앙값 {ls[len(ls) // 2]:.0f}mm · "
              f"최소 {ls[0]:.1f}mm · 최대 {ls[-1]:.0f}mm")
        for lim in (1, 10, 30, 50, A.MIN_PIPE_EDGE_MM, 100, 300):
            n = sum(1 for v in ls if v < lim)
            print(f"    {lim:>7.0f}mm 미만  {n:>8,}  ({n / len(ls) * 100:.1f}%)")

    # ── ③ 앵커 작업창 안에 몇 개나 드나
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
    zx0, zy0, zx1, zy1 = zones[0]
    m = A.ANCHOR_W_MARGIN_MM
    n_in = sum(1 for p, q in segs
               if (zx0 - m <= p[0] <= zx1 + m and zy0 - m <= p[1] <= zy1 + m
                   and zx0 - m <= q[0] <= zx1 + m and zy0 - m <= q[1] <= zy1 + m))
    print(f"\n■ 앵커 작업창 (영역 + 여유 {m:.0f}mm)")
    print(f"    창 안에 양끝이 다 든 선분 {n_in:,} / {len(segs):,} "
          f"({n_in / max(1, len(segs)) * 100:.1f}%)")

    # ── ④ 실제 그래프
    real = A._join_head_gap_endpoints

    def grab(graph, edge_len, head_pts, *args, **kw):
        n = real(graph, edge_len, head_pts, *args, **kw)
        GRAB["nodes"] = len(graph)
        GRAB["edges"] = len(edge_len)
        GRAB["len_mm"] = sum(edge_len.values())
        return n

    A._join_head_gap_endpoints = grab
    try:
        A.select_worst30_heads_anchored(
            pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
            head_region=HeadRegion.from_rects(zones), zones=zones, k=30)
    finally:
        A._join_head_gap_endpoints = real

    print(f"\n■ 그래프 — 절점 {GRAB.get('nodes', 0):,} · "
          f"간선 {GRAB.get('edges', 0):,} · "
          f"연장 {GRAB.get('len_mm', 0) / 1000:,.0f} m")
    tot = sum(ls) / 1000.0
    print(f"    배관 선분 총연장 {tot:,.0f} m → 그래프 "
          f"{GRAB.get('len_mm', 0) / 1000:,.0f} m "
          f"({GRAB.get('len_mm', 0) / 1000 / max(1e-9, tot) * 100:.1f}%)")
    print("\n  선분 대비 간선이 크게 줄었다면 «틈» 이 아니라 «손실» 이 문제다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
