# -*- coding: utf-8 -*-
"""셔터선 구역을 «바짝» 확대해 원본 도형을 그대로 본다.

앞선 측정은 헤드 결합 상한(300 mm)으로 «어느 선에 붙었나» 를 갈랐는데,
사용자가 「분홍선 사이사이에도 헤드가 정렬돼 있다」고 한다. 그렇다면 내가
못 본 것이 있다 — 요약 수치 말고 도형을 그대로 봐야 한다.

    · 레이어×색을 색으로 갈라 전부 그린다 (배관/비배관 가리지 않고)
    · 헤드 기호 자체(원·삼각·십자 획)도 함께 그린다
    · 몇 미터짜리 창으로 잘라 여러 칸에 나눠 본다

    python scripts/_probe_shutter_close.py [도면.dxf] [--span 6000]
"""
from __future__ import annotations

import argparse
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"

PALETTE = ["#22d3ee", "#e879f9", "#f59e0b", "#4ade80", "#60a5fa",
           "#f87171", "#a78bfa", "#fbbf24"]


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
    ap.add_argument("--span", type=float, default=6000.0)
    ap.add_argument("--rows", type=int, default=3)
    ap.add_argument("--out", default=str(ROOT / "data" / "_shutter_close.png"))
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib.collections import LineCollection

    import remote30_prototype as A

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

    # 셔터선(미승격)의 y 대역을 창의 기준으로 삼는다.
    raw = []
    for e in ents:
        ly = str(e.get("l") or "0")
        if "셔터" in ly and "승격" not in ly:
            raw.extend(_segs(e))
    if not raw:
        print("셔터선(미승격)이 없다.")
        return 1
    ys = [c for p, q in raw for c in (p[1], q[1])]
    xs = [c for p, q in raw for c in (p[0], q[0])]
    ymid = (min(ys) + max(ys)) / 2.0
    x0all, x1all = min(xs), max(xs)
    print(f"셔터선 x {x0all:.0f}~{x1all:.0f} · y {min(ys):.0f}~{max(ys):.0f}")

    # 창 둘레의 모든 도형을 레이어×색으로 모은다.
    CELL = 4000.0
    grid = defaultdict(list)
    for e in ents:
        ly = str(e.get("l") or "0")
        key = f"{ly}×{e.get('c')}"
        for (p, q) in _segs(e):
            gx = int(((p[0] + q[0]) / 2) // CELL)
            gy = int(((p[1] + q[1]) / 2) // CELL)
            grid[(gx, gy)].append((key, cat.get(ly, "OTHER"), p, q))
    # 원은 entity 목록 안에 'C' 로 들어 있다 — 별도 배열이 아니다.
    cgrid = defaultdict(list)
    for e in ents:
        if str(e.get("t") or "") != "C":
            continue
        p = e.get("p") or []
        r = float(e.get("r") or 0)
        if len(p) < 2 or r <= 0:
            continue
        cx, cy = float(p[0]), float(p[1])
        cgrid[(int(cx // CELL), int(cy // CELL))].append((cx, cy, r))
    hg = defaultdict(list)
    for h in heads:
        hg[(int(h.pos[0] // CELL), int(h.pos[1] // CELL))].append(h.pos)

    span = a.span
    rows = a.rows
    step = (x1all - x0all) / rows
    fig, axes = plt.subplots(rows, 1, figsize=(19, 3.4 * rows), dpi=115)
    axes = [axes] if rows == 1 else list(axes)
    fig.patch.set_facecolor("#0a0d12")
    colors: dict = {}

    for i, ax in enumerate(axes):
        cx0 = x0all + step * i + step / 2
        lo, hi = cx0 - span / 2, cx0 + span / 2
        ax.set_facecolor("#0a0d12")
        cells = set()
        for gx in range(int(lo // CELL) - 1, int(hi // CELL) + 2):
            for gy in range(int((ymid - 2500) // CELL) - 1,
                            int((ymid + 2500) // CELL) + 2):
                cells.add((gx, gy))
        by_key = defaultdict(list)
        for c in cells:
            for (key, ct, p, q) in grid.get(c, ()):
                if max(p[0], q[0]) < lo or min(p[0], q[0]) > hi:
                    continue
                if abs((p[1] + q[1]) / 2 - ymid) > 2200:
                    continue
                by_key[key].append([(p[0], p[1]), (q[0], q[1])])
        for key, segs in sorted(by_key.items(), key=lambda kv: -len(kv[1])):
            if key not in colors:
                colors[key] = PALETTE[len(colors) % len(PALETTE)]
            ax.add_collection(LineCollection(
                segs, colors=colors[key], linewidths=1.8,
                label=key if i == 0 else None))
        cc = [(x, y, r) for c in cells for (x, y, r) in cgrid.get(c, ())
              if lo <= x <= hi and abs(y - ymid) <= 2200]
        for (x, y, r) in cc:
            ax.add_patch(plt.Circle((x, y), r, fill=False,
                                    edgecolor="#ffffff", linewidth=0.9))
        hp = [p for c in cells for p in hg.get(c, ())
              if lo <= p[0] <= hi and abs(p[1] - ymid) <= 2200]
        if hp:
            ax.scatter([p[0] for p in hp], [p[1] for p in hp], s=30,
                       c="#ff3b30", zorder=6, linewidths=0)
        ax.set_xlim(lo, hi)
        ax.set_ylim(ymid - 900, ymid + 900)
        ax.set_aspect("equal")
        ax.axis("off")

    if colors:
        handles = [plt.Line2D([], [], color=v, lw=2.4, label=k)
                   for k, v in colors.items()]
        handles.append(plt.Line2D([], [], color="#ff3b30", marker="o", lw=0,
                                  label="detected head"))
        handles.append(plt.Line2D([], [], color="#ffffff", lw=1.0,
                                  label="circle (head symbol)"))
        fig.legend(handles=handles, loc="upper center", ncol=4,
                   facecolor="#0a0d12", edgecolor="#334155",
                   labelcolor="#cbd5e1", fontsize=8)
    fig.tight_layout(rect=(0, 0, 1, 0.93))
    fig.savefig(a.out, facecolor=fig.get_facecolor())
    print("그림:", a.out)
    for k, v in colors.items():
        print(f"  {v}  {k}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
