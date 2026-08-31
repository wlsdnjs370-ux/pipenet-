# -*- coding: utf-8 -*-
"""계통도·기계실을 «눈으로» 본다 — 레이어별 위치와 겹침.

수치만으로는 안 갈리는 것이 있었다. 계통도에서 레이어 「4」를 배관으로 치면
큰 덩이가 51 → 331 노드로 커지는데, 그 덩이 안에 원래 배관(SP·LSP) 노드가
**0개** 들어온다. 둘이 아예 다른 자리에 있다는 뜻이다 — 한 장에 여러 그림이
올라가 있거나(범례·상세도), 같은 그림의 다른 계통이거나.

그림 한 장이면 바로 갈린다. 레이어별 경계상자와 실제 선을 그린다.

    python scripts/_probe_riser_mr_view.py [도면.dxf ...]
    → data/_riser_mr/<이름>.png
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt          # noqa: E402

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "data" / "_riser_mr"
DEF = [
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 계통도.dxf",
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 기계실.dxf",
]
# 색은 «무엇인지» 로 고른다 — 이름이 아니라 역할이 눈에 들어와야 한다.
PALETTE = ["#e11d48", "#2563eb", "#16a34a", "#f59e0b", "#a855f7",
           "#0891b2", "#db2777", "#65a30d", "#dc2626", "#4f46e5"]


def segs_of(en) -> list:
    """entity 하나 → 선분 목록. 그리기용이라 호는 현으로 근사한다."""
    t = str(en.get("t") or "")
    if t == "L":
        x1, y1, x2, y2 = en["p"]
        return [((x1, y1), (x2, y2))]
    if t == "PL":
        pts = en["p"]
        return [((a[0], a[1]), (b[0], b[1])) for a, b in zip(pts, pts[1:])]
    if t == "A":
        import math
        cx, cy = en["c"]
        r = float(en.get("r", 0.0) or 0.0)
        if r <= 0:
            return []
        sa, ea = en.get("a", [0.0, 0.0])
        n = 8
        out = []
        for i in range(n):
            t0 = math.radians(sa + (ea - sa) * i / n)
            t1 = math.radians(sa + (ea - sa) * (i + 1) / n)
            out.append(((cx + r * math.cos(t0), cy + r * math.sin(t0)),
                        (cx + r * math.cos(t1), cy + r * math.sin(t1))))
        return out
    if t == "C":
        import math
        cx, cy = en["c"]
        r = float(en.get("r", 0.0) or 0.0)
        if r <= 0:
            return []
        n = 12
        return [((cx + r * math.cos(2 * math.pi * i / n),
                  cy + r * math.sin(2 * math.pi * i / n)),
                 (cx + r * math.cos(2 * math.pi * (i + 1) / n),
                  cy + r * math.sin(2 * math.pi * (i + 1) / n)))
                for i in range(n)]
    return []


def _split_draw(dxf, big, layers, heads, zoom, tag) -> None:
    """레이어 한 장에 한 칸 — «이 선이 무엇인가» 를 짐작하지 않게."""
    n = len(big)
    cols = 3
    rows = (n + cols - 1) // cols
    fig, axes = plt.subplots(rows, cols, figsize=(16, 4.2 * rows), dpi=110)
    axes = axes.ravel() if hasattr(axes, "ravel") else [axes]
    for k, (nm, ee) in enumerate(big):
        ax = axes[k]
        lines = []
        for en in ee:
            lines += segs_of(en)
        for (a, b) in lines[:9000]:
            ax.plot([a[0], b[0]], [a[1], b[1]], color="#111827",
                    lw=0.6, alpha=0.9)
        # 헤드 후보는 «어느 칸에서든» 같은 자리에 찍어 대조가 되게 한다.
        if heads:
            ax.scatter([h.pos[0] for h in heads], [h.pos[1] for h in heads],
                       s=18, facecolors="none", edgecolors="#e11d48",
                       linewidths=0.9)
        ax.set_aspect("equal")
        if zoom:
            ax.set_xlim(zoom[0], zoom[2])
            ax.set_ylim(zoom[1], zoom[3])
        ax.set_title(f"{nm}  [{layers.get(nm, 'OTHER')}]  {len(ee):,}",
                     fontsize=9)
        ax.tick_params(labelsize=6)
    for k in range(n, len(axes)):
        axes[k].axis("off")
    out = OUT / (dxf.stem + tag + ".png")
    fig.tight_layout()
    fig.savefig(out, bbox_inches="tight")
    plt.close(fig)
    print(f"   → {out}   (붉은 동그라미 = 헤드 후보, 모든 칸에 같이 찍음)")


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    OUT.mkdir(parents=True, exist_ok=True)
    args = list(sys.argv[1:])
    # --zoom x0,y0,x1,y1 — 한 층만 크게 본다. 전체 그림은 층이 스물다섯 겹이라
    #   무엇이 어느 레이어인지 안 보인다.
    zoom = None
    tag = ""
    if "--zoom" in args:
        i = args.index("--zoom")
        zoom = [float(v) for v in args[i + 1].split(",")]
        tag = "_zoom"
        del args[i:i + 2]
    # --split — 레이어를 한 칸씩 따로 그린다. 겹쳐 그리면 «이 선이 어느
    #   레이어인지» 를 색으로 짐작하게 되는데, 그 짐작이 틀리면 진단 전체가
    #   틀린다. 따로 그리면 짐작할 것이 없다.
    split = "--split" in args
    if split:
        args.remove("--split")
        tag += "_split"
    for dxf in [Path(x) for x in args] or DEF:
        if not dxf.is_file():
            print(f"■ {dxf.name} — 파일 없음")
            continue
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = bundle.entities
        layers = {ly.get("name"): (ly.get("auto_category") or "OTHER")
                  for ly in (bundle.layers or [])}
        heads = A.detect_heads(ents, layers)

        by_layer = defaultdict(list)
        for en in ents:
            by_layer[str(en.get("l") or "0")].append(en)
        big = sorted(by_layer.items(), key=lambda kv: -len(kv[1]))[:9]

        print(f"\n■ {dxf.name}")
        print(f"   {'레이어':<14}{'분류':<7}{'entity':>8}"
              f"{'x범위':>26}{'y범위':>26}")
        if split:
            _split_draw(dxf, big, layers, heads, zoom, tag)
            continue
        fig, ax = plt.subplots(figsize=(15, 10), dpi=110)
        for i, (nm, ee) in enumerate(big):
            xs, ys, lines = [], [], []
            for en in ee:
                for (a, b) in segs_of(en):
                    lines.append((a, b))
                    xs += [a[0], b[0]]
                    ys += [a[1], b[1]]
            if not xs:
                print(f"   {nm:<14}{layers.get(nm, 'OTHER'):<7}"
                      f"{len(ee):>8,}    (그릴 선 없음)")
                continue
            c = PALETTE[i % len(PALETTE)]
            for (a, b) in lines[:9000]:      # 너무 많으면 잘라 그린다
                ax.plot([a[0], b[0]], [a[1], b[1]], color=c,
                        lw=0.5, alpha=0.75)
            ax.plot([], [], color=c, lw=2,
                    label=f"{nm} [{layers.get(nm, 'OTHER')}] {len(ee):,}")
            print(f"   {nm:<14}{layers.get(nm, 'OTHER'):<7}{len(ee):>8,}"
                  f"{min(xs):>12,.0f}~{max(xs):<13,.0f}"
                  f"{min(ys):>12,.0f}~{max(ys):<13,.0f}")
        if heads:
            ax.scatter([h.pos[0] for h in heads], [h.pos[1] for h in heads],
                       s=14, facecolors="none", edgecolors="#111827",
                       linewidths=0.8,
                       label=f"헤드 후보 {len(heads):,}")
        ax.set_aspect("equal")
        if zoom:
            ax.set_xlim(zoom[0], zoom[2])
            ax.set_ylim(zoom[1], zoom[3])
        ax.legend(loc="upper right", fontsize=8)
        ax.set_title(dxf.name, fontsize=10)
        out = OUT / (dxf.stem + tag + ".png")
        fig.savefig(out, bbox_inches="tight")
        plt.close(fig)
        print(f"   → {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
