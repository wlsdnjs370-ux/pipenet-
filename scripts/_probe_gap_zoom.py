# -*- coding: utf-8 -*-
"""그 틈을 «확대해서» 본다 — 도면 원본 도형을 분류별 색으로.

수치로는 「틈이 비었다」까지 왔다. 그런데 그 판정은 내가 짠 셈이므로, 마지막은
눈으로 확인한다. 틈 몇 곳을 골라 그 둘레를 확대하고, DXF 도형을 레이어 분류
색으로 전부 그린다.

    청록  PIPE      빨강  HEAD     주황  ALARM
    회색  ARCH      보라  OTHER    노랑  EXCLUDE
    흰 점선 = 끊긴 끝점에서 주관까지 (재려는 그 틈)

    python scripts/_probe_gap_zoom.py [도면.dxf] [--n 6]
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
GRAB: dict = {}
COLOR = {"PIPE": "#22d3ee", "HEAD": "#ff3b30", "ALARM": "#f97316",
         "ARCH": "#3a4250", "OTHER": "#a855f7", "EXCLUDE": "#eab308",
         "TEXT": "#64748b"}


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
    ap.add_argument("--n", type=int, default=6)
    ap.add_argument("--out", default=str(ROOT / "data" / "_gap_zoom.png"))
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.collections import LineCollection

    import remote30_prototype as A
    from remote30_graph import HeadRegion, _point_to_segment_dist

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    real = A._join_head_gap_endpoints

    def grab(graph, edge_len, head_pts, *args, **kw):
        n = real(graph, edge_len, head_pts, *args, **kw)
        GRAB["graph"] = {k: set(v) for k, v in graph.items()}
        GRAB["edge_len"] = dict(edge_len)
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
        A.select_worst30_heads_anchored(
            pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
            head_region=HeadRegion.from_rects(zones), zones=zones, k=30)
    finally:
        A._join_head_gap_endpoints = real

    graph, edge_len = GRAB["graph"], GRAB["edge_len"]
    ends = [n for n, nb in graph.items() if len(nb) == 1]
    comp: dict = {}

    def find(x):
        r = x
        while comp.get(r, r) != r:
            r = comp[r]
        while comp.get(x, x) != x:
            comp[x], x = r, comp[x]
        return r

    for u, v in edge_len:
        ru, rv = find(u), find(v)
        if ru != rv:
            comp[ru] = rv

    CELL = 4000.0
    sgrid = defaultdict(list)
    for (u, v) in edge_len:
        x0, x1 = sorted((u[0], v[0]))
        y0, y1 = sorted((u[1], v[1]))
        for gx in range(int(x0 // CELL), int(x1 // CELL) + 1):
            for gy in range(int(y0 // CELL), int(y1 // CELL) + 1):
                sgrid[(gx, gy)].append((u, v))

    picks = []
    for u in ends:
        gx, gy = int(u[0] // CELL), int(u[1] // CELL)
        cells = [(gx + dx, gy + dy) for dx in (-1, 0, 1) for dy in (-1, 0, 1)]
        ru = find(u)
        best, bd = None, float("inf")
        for (p, q) in (s for c in cells for s in sgrid.get(c, ())):
            if find(p) == ru:
                continue
            d = _point_to_segment_dist(u[0], u[1], p[0], p[1], q[0], q[1])
            if d < bd:
                bd, best = d, (p, q)
        if best is not None and 700.0 <= bd <= 1200.0:
            picks.append((bd, u, best))
    picks.sort()
    n = min(a.n, len(picks))
    print(f"700~1200mm 틈 {len(picks)}곳 — 앞 {n}곳을 확대")
    if not n:
        return 1

    # 도형을 셀 격자에 담아 창별로 빠르게 뽑는다.
    egrid = defaultdict(list)
    for e in ents:
        c = cat.get(str(e.get("l") or "0"), "OTHER")
        for (p, q) in _segs(e):
            gx = int(((p[0] + q[0]) / 2) // CELL)
            gy = int(((p[1] + q[1]) / 2) // CELL)
            egrid[(gx, gy)].append((c, p, q))
    hgrid = defaultdict(list)
    for hp in pts:
        hgrid[(int(hp[0] // CELL), int(hp[1] // CELL))].append(hp)

    cols = 3
    rows = (n + cols - 1) // cols
    fig, axes = plt.subplots(rows, cols, figsize=(6 * cols, 5.2 * rows), dpi=100)
    axes = [axes] if n == 1 else list(axes.flat)
    fig.patch.set_facecolor("#0a0d12")

    for i, (bd, u, (p, q)) in enumerate(picks[:n]):
        ax = axes[i]
        ax.set_facecolor("#0a0d12")
        R = max(2200.0, bd * 2.4)
        gx, gy = int(u[0] // CELL), int(u[1] // CELL)
        cells = [(gx + dx, gy + dy) for dx in (-2, -1, 0, 1, 2)
                 for dy in (-2, -1, 0, 1, 2)]
        by_cat = defaultdict(list)
        for (c, sa, sb) in (s for cc in cells for s in egrid.get(cc, ())):
            if (abs(sa[0] - u[0]) > R and abs(sb[0] - u[0]) > R) or \
               (abs(sa[1] - u[1]) > R and abs(sb[1] - u[1]) > R):
                continue
            by_cat[c].append([(sa[0], sa[1]), (sb[0], sb[1])])
        for c in ("ARCH", "OTHER", "EXCLUDE", "TEXT", "HEAD", "ALARM", "PIPE"):
            if by_cat.get(c):
                ax.add_collection(LineCollection(
                    by_cat[c], colors=COLOR[c],
                    linewidths=1.9 if c == "PIPE" else 0.8,
                    zorder=3 if c == "PIPE" else 1))
        hh = [h for cc in cells for h in hgrid.get(cc, ())
              if abs(h[0] - u[0]) <= R and abs(h[1] - u[1]) <= R]
        if hh:
            ax.scatter([h[0] for h in hh], [h[1] for h in hh], s=14,
                       c="#ff3b30", zorder=4, linewidths=0)
        # 재려는 틈 — 끝점에서 주관까지 수선
        dx, dy = q[0] - p[0], q[1] - p[1]
        L2 = dx * dx + dy * dy or 1.0
        t = max(0.0, min(1.0, ((u[0] - p[0]) * dx + (u[1] - p[1]) * dy) / L2))
        fx, fy = p[0] + t * dx, p[1] + t * dy
        ax.plot([u[0], fx], [u[1], fy], ls="--", lw=1.6, color="#ffffff",
                zorder=5)
        ax.plot([u[0]], [u[1]], "o", ms=7, mfc="none", mec="#ffffff", mew=1.6,
                zorder=6)
        ax.set_xlim(u[0] - R, u[0] + R)
        ax.set_ylim(u[1] - R, u[1] + R)
        ax.set_aspect("equal")
        ax.set_title(f"틈 {bd:.0f}mm", color="#cbd5e1", fontsize=11)
        ax.axis("off")
    for j in range(n, len(axes)):
        axes[j].axis("off")
    fig.tight_layout(pad=0.6)
    fig.savefig(a.out, facecolor=fig.get_facecolor())
    print("그림:", a.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
