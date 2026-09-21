# -*- coding: utf-8 -*-
"""PIPENET 도면의 판면을 잰다 — 범례와 표제란이 어디에 있고, 망은 종이를 얼마나 채우나."""
from __future__ import annotations

import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent

import fitz  # noqa: E402

LIB = ROOT / "data" / "reference_library"


def main(limit: int) -> None:
    page_sizes, legend_boxes, net_boxes, text_below = [], [], [], []
    files = 0
    for pdf in sorted(LIB.rglob("*3.유량.pdf")):
        doc = fitz.open(pdf)
        page = doc[0]
        W, H = page.rect.width, page.rect.height
        swatches, net = [], None
        for d in page.get_drawings():
            r = d["rect"]
            if d["fill"] is not None and 3 < r.width < 30 and 3 < r.height < 30 \
                    and "".join(i[0] for i in d["items"]) in ("re", "l"):
                if tuple(round(v, 3) for v in d["fill"]) != (1.0, 1.0, 1.0):
                    swatches.append(r)
        if not swatches:
            doc.close()
            continue
        sw = swatches[0]
        for a in swatches[1:]:
            sw |= a
        # 망: 범례가 아닌 획들의 외접 상자
        for d in page.get_drawings():
            if d["color"] is None:
                continue
            r = d["rect"]
            if r.intersects(sw) or (r.width < 1 and r.height < 1):
                continue
            net = r if net is None else (net | r)
        below = [(b[0] / W, b[2] / W, 1 - b[3] / H, 1 - b[1] / H, b[4].strip()[:40])
                 for b in page.get_text("blocks") if 1 - b[3] / H < 0.14]
        doc.close()
        if net is None:
            continue
        files += 1
        page_sizes.append((round(W), round(H)))
        legend_boxes.append((sw.x0 / W, sw.x1 / W, 1 - sw.y1 / H, 1 - sw.y0 / H))
        net_boxes.append((net.x0 / W, net.x1 / W, 1 - net.y1 / H, 1 - net.y0 / H))
        text_below.append((W > H, below))
        if files >= limit:
            break

    def col(rows, i):
        return statistics.median(r[i] for r in rows)

    print(f"판면을 읽은 도면 {files}장")
    print(f"  종이 크기: {statistics.mode(page_sizes)} pt (상위 {sorted(set(page_sizes))[:3]})")
    for name, rows in (("범례", legend_boxes), ("망", net_boxes)):
        print(f"  {name} 가로 {col(rows,0):.3f}~{col(rows,1):.3f}  "
              f"세로(아래=0) {col(rows,2):.3f}~{col(rows,3):.3f}")
    print("  아래쪽 14% 안의 글줄 상자 (첫 3장):")
    for land, blocks in text_below[:3]:
        print(f"    [{'가로' if land else '세로'}]")
        for x0, x1, y0, y1, t in sorted(blocks, key=lambda b: -b[3])[:8]:
            print(f"      x {x0:.3f}~{x1:.3f}  y {y0:.3f}~{y1:.3f}  {t!r}")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
