# -*- coding: utf-8 -*-
"""`flow-arrows` 가 화살표 스위치인지 원본 PDF 로 반증한다."""
from __future__ import annotations

import math
import pathlib
import re

import fitz

LIB = pathlib.Path(__file__).resolve().parent.parent / "data" / "reference_library"
PURE = {(1.0, 0.0, 0.0), (1.0, 0.676, 0.0), (0.0, 1.0, 0.0),
        (0.0, 1.0, 1.0), (0.0, 0.0, 1.0), (1.0, 0.0, 1.0)}


def flag(text: str, tag: str, key: str) -> str:
    m = re.search(rf"<{tag}([^>]*)/>", text)
    if not m:
        return "NA"
    k = re.search(rf'{key}="([^"]*)"', m.group(1))
    return k.group(1) if k else "-"


def chevrons(pdf: pathlib.Path) -> int:
    doc = fitz.open(pdf)
    n = 0
    for d in doc[0].get_drawings():
        col = None if d["color"] is None else tuple(round(v, 3) for v in d["color"])
        s = [((it[1].x, it[1].y), (it[2].x, it[2].y))
             for it in d["items"] if it[0] == "l"]
        if col not in PURE or len(s) != 2:
            continue
        a, b, c = s[0][0], s[0][1], s[1][1]
        l1, l2 = math.dist(a, b), math.dist(b, c)
        ang = math.degrees(abs(math.atan2(a[1] - b[1], a[0] - b[0])
                               - math.atan2(c[1] - b[1], c[0] - b[0])))
        ang = min(ang, 360 - ang)
        if 40 < ang < 70 and abs(l1 - l2) < 0.25 * max(l1, l2) and max(l1, l2) < 3:
            n += 1
    doc.close()
    return n


pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
rows = []
for sdf in sorted(LIB.rglob("*.sdf")):
    pdf = pdfs.get(sdf.stem)
    if pdf is None:
        continue
    t = sdf.read_text(encoding="utf-8", errors="replace")
    rows.append((flag(t, "Results-display", "flow-arrows"),
                 flag(t, "Label-display", "arrows"), chevrons(pdf), sdf.stem))

print(f"짝지어진 세트 {len(rows)}")
for want in ("0", "1"):
    sub = [r for r in rows if r[0] == want]
    if not sub:
        continue
    drawn = [r for r in sub if r[2] > 0]
    print(f"  flow-arrows={want}: {len(sub)}세트 중 화살표가 그려진 것 {len(drawn)}"
          f"  갈매기 수 중앙 {sorted(r[2] for r in sub)[len(sub) // 2]}")
print("  Label-display arrows 값 분포:",
      {v: sum(1 for r in rows if r[1] == v) for v in {r[1] for r in rows}})
for r in rows:
    if r[0] == "0":
        print(f"    flow-arrows=0 · arrows={r[1]} · 갈매기 {r[2]:3d}  {r[3][:52]}")
