# -*- coding: utf-8 -*-
"""범례 격자를 여러 장에서 모은다 — 몇 줄·몇 칸이고, 노드 범례 색은 무엇인가."""
from __future__ import annotations

import collections
import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent

import fitz  # noqa: E402

LIB = ROOT / "data" / "reference_library"


def swatch_grid(page):
    """범례 칸만 골라 (아래에서 잰 y, x, 색) 로 돌려준다."""
    H = page.rect.height
    out = []
    for d in page.get_drawings():
        r = d["rect"]
        if d["fill"] is None or not (7 < r.width < 10 and 7 < r.height < 10):
            continue
        c = tuple(round(v, 4) for v in d["fill"])
        if c == (1.0, 1.0, 1.0) or 1 - r.y1 / H > 0.12:
            continue
        out.append((round(H - r.y1, 1), round(r.x0, 1), c))
    return out


def main(pattern: str, limit: int) -> None:
    shapes = collections.Counter()
    ramps = collections.Counter()
    pitches_row, pitches_col, gaps = [], [], []
    files = 0
    for pdf in sorted(LIB.rglob(pattern)):
        doc = fitz.open(pdf)
        page = doc[0]
        if page.rect.width > page.rect.height:
            doc.close()
            continue
        grid = swatch_grid(page)
        doc.close()
        if not grid:
            continue
        files += 1
        ys = sorted({y for y, _, _ in grid}, reverse=True)
        xs = sorted({x for _, x, _ in grid})
        shapes[(len(ys), len(xs))] += 1
        if len(ys) >= 2:
            pitches_row.append(ys[0] - ys[1])
        if len(xs) >= 2:
            pitches_col.append(xs[1] - xs[0])
        if len(ys) >= 4:
            gaps.append(ys[1] - ys[2])
        # 맨 위 두 줄 = 첫 범례. 줄 우선으로 읽는다.
        top = [c for _, _, c in sorted(g for g in grid if g[0] in ys[:2])]
        top = [c for y in ys[:2] for _, _, c in sorted(g for g in grid if g[0] == y)]
        ramps[tuple("#%02x%02x%02x" % tuple(round(v * 255) for v in c) for c in top)] += 1
        if files >= limit:
            break

    print(f"{pattern} 세로 도면 {files}장")
    print(f"  격자 (줄, 칸): {shapes.most_common(4)}")
    for name, vals in (("줄 간격", pitches_row), ("칸 간격", pitches_col),
                       ("범례 사이 간격", gaps)):
        if vals:
            print(f"  {name} 중앙 {statistics.median(vals):.2f} pt "
                  f"(최소 {min(vals):.2f} 최대 {max(vals):.2f})")
    print("  맨 위 범례 색:")
    for ramp, n in ramps.most_common(3):
        print(f"    {n:3d}장  {list(ramp)}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "*3.유량.pdf",
         int(sys.argv[2]) if len(sys.argv) > 2 else 40)
