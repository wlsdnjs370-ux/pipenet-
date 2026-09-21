# -*- coding: utf-8 -*-
"""세로 도면 한 장의 범례를 있는 그대로 뜯어본다 — 칸 하나하나의 자리와 크기."""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent

import fitz  # noqa: E402

LIB = ROOT / "data" / "reference_library"

for pdf in sorted(LIB.rglob("*2.압력.pdf")):
    doc = fitz.open(pdf)
    page = doc[0]
    W, H = page.rect.width, page.rect.height
    if W > H:
        doc.close()
        continue
    boxes = [(d["rect"], d["fill"]) for d in page.get_drawings()
             if d["fill"] is not None and 1 - d["rect"].y1 / H < 0.12
             and d["rect"].width < 60 and d["rect"].height < 30]
    if len(boxes) < 6:
        doc.close()
        continue
    print(pdf.relative_to(LIB))
    print(f"  종이 {W:.0f}x{H:.0f} pt")
    for r, f in sorted(boxes, key=lambda b: b[0].x0):
        print(f"    x {r.x0:7.2f}~{r.x1:7.2f} ({r.width:5.2f})  "
              f"y아래 {H-r.y1:6.2f}~{H-r.y0:6.2f} ({r.height:5.2f})  "
              f"#{''.join('%02x' % round(v*255) for v in f)}")
    for b in page.get_text("blocks"):
        if 1 - b[3] / H < 0.16:
            print(f"    글 x {b[0]:7.2f}~{b[2]:7.2f}  y아래 {H-b[3]:6.2f}~{H-b[1]:6.2f}  "
                  f"{b[4].strip()!r}")
    doc.close()
    break

sys.exit(0)
