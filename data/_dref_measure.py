# -*- coding: utf-8 -*-
"""PIPENET 원본 ISO PDF 의 표기법을 눈짐작이 아니라 수치로 잰다."""
from __future__ import annotations

import collections
import math
import pathlib
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
LIB = ROOT / "data" / "reference_library"


def rgb(c):
    if c is None:
        return None
    return tuple(round(v, 3) for v in c)


def describe(path: pathlib.Path) -> None:
    doc = fitz.open(path)
    page = doc[0]
    print("=" * 78)
    print(path.name, "|", page.rect)
    draws = page.get_drawings()
    kinds = collections.Counter()
    for d in draws:
        ops = "".join(it[0] for it in d["items"])
        kinds[(d["type"], ops, rgb(d.get("color")), rgb(d.get("fill")),
               round(d.get("width") or 0, 3))] += 1
    for key, n in kinds.most_common(40):
        print(f"  {n:5d}  type={key[0]:<3} ops={key[1][:20]:<20} "
              f"stroke={key[2]} fill={key[3]} w={key[4]}")
    doc.close()


def geometry(path: pathlib.Path) -> None:
    """닫힌 삼각형(노즐 후보)과 짧은 2획 도형(화살표 후보)을 골라 치수를 잰다."""
    doc = fitz.open(path)
    page = doc[0]
    tri, chev = [], []
    for d in page.get_drawings():
        items = d["items"]
        pts: list[tuple[float, float]] = []
        for it in items:
            if it[0] == "l":
                pts += [(it[1].x, it[1].y), (it[2].x, it[2].y)]
        if not pts:
            continue
        xs = [p[0] for p in pts]
        ys = [p[1] for p in pts]
        size = max(max(xs) - min(xs), max(ys) - min(ys))
        uniq = {(round(x, 2), round(y, 2)) for x, y in pts}
        rec = (size, len(items), rgb(d.get("color")), rgb(d.get("fill")),
               round(d.get("width") or 0, 3), d["type"], sorted(uniq))
        if len(uniq) == 3:
            tri.append(rec)
        elif len(uniq) == 3 or (len(items) == 2 and size < 20):
            chev.append(rec)
    print(f"  삼각형 후보 {len(tri)}  2획 소형 {len(chev)}")
    for name, bag in (("삼각형", tri), ("2획", chev)):
        agg = collections.Counter((round(r[0], 1), r[1], r[2], r[3], r[4], r[5])
                                  for r in bag)
        for key, n in agg.most_common(8):
            print(f"    [{name}] {n:4d}  size={key[0]} items={key[1]} "
                  f"stroke={key[2]} fill={key[3]} w={key[4]} type={key[5]}")
    if tri:
        print("    삼각형 첫 3개 좌표:")
        for r in tri[:3]:
            print("      ", r[6])
    doc.close()


if __name__ == "__main__":
    targets = sys.argv[1:]
    if not targets:
        targets = [str(p) for p in sorted(LIB.rglob("*3.유량.pdf"))[:2]]
    for t in targets:
        p = pathlib.Path(t)
        describe(p)
        geometry(p)
