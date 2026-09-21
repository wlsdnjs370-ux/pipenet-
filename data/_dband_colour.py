# -*- coding: utf-8 -*-
"""PIPENET 범례의 띠 색을 실제 색값으로 뽑는다 — 몇 단계이고 어떤 순서인가."""
from __future__ import annotations

import collections
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent

import fitz  # noqa: E402

LIB = ROOT / "data" / "reference_library"


def main(limit: int) -> None:
    seqs = collections.Counter()
    used_by_pipes = collections.Counter()
    files = 0
    for pdf in sorted(LIB.rglob("*3.유량.pdf")):
        doc = fitz.open(pdf)
        page = doc[0]
        h = page.rect.height
        swatches, strokes = [], collections.Counter()
        for d in page.get_drawings():
            if d["fill"] is not None and "".join(i[0] for i in d["items"]) in ("re", "l"):
                r = d["rect"]
                if r.y0 > h * 0.8 and 3 < r.width < 30 and 3 < r.height < 30:
                    swatches.append((round(r.y0, 1), round(r.x0, 1),
                                     tuple(round(v, 3) for v in d["fill"])))
            if d["color"] is not None and set(i[0] for i in d["items"]) == {"l"}:
                if d["rect"].y1 < h * 0.8:
                    strokes[tuple(round(v, 3) for v in d["color"])] += len(d["items"])
        doc.close()
        if not swatches:
            continue
        files += 1
        seqs[tuple(c for _, _, c in sorted(swatches))] += 1
        used_by_pipes.update(strokes)
        if files >= limit:
            break

    print(f"범례를 읽은 도면 {files}장")
    for seq, n in seqs.most_common(4):
        print(f"  {n:3d}장  {len(seq)}단계  {[('#%02x%02x%02x' % tuple(round(v*255) for v in c)) for c in seq]}")
    print("  관로 획에 실제로 쓰인 색 상위:")
    for c, n in used_by_pipes.most_common(10):
        print(f"    {'#%02x%02x%02x' % tuple(round(v*255) for v in c)}  도막 {n}")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
