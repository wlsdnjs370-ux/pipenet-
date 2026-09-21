# -*- coding: utf-8 -*-
"""노즐 기호의 붙는 자리를 확정하고, 그 자리를 실제 그림으로 오려 본다."""
from __future__ import annotations

import math
import pathlib
import sys

import fitz

PURE = {(1.0, 0.0, 0.0), (1.0, 0.676, 0.0), (0.0, 1.0, 0.0),
        (0.0, 1.0, 1.0), (0.0, 0.0, 1.0), (1.0, 0.0, 1.0)}
OUT = pathlib.Path(__file__).resolve().parent / "_dref"
OUT.mkdir(exist_ok=True)


def rgb(c):
    return None if c is None else tuple(round(v, 3) for v in c)


def main(path: pathlib.Path, tag: str) -> None:
    doc = fitz.open(path)
    page = doc[0]
    pipe_pts, poly5, squares, nodes = [], [], [], []
    for d in page.get_drawings():
        col, fill = rgb(d["color"]), rgb(d["fill"])
        s = [((it[1].x, it[1].y), (it[2].x, it[2].y))
             for it in d["items"] if it[0] == "l"]
        ops = "".join(it[0] for it in d["items"])
        if col in PURE and s:
            pipe_pts += [s[0][0]] + [q for _, q in s]
        elif ops == "lllll" and col == (0.0, 0.0, 0.0):
            poly5.append([s[0][0]] + [q for _, q in s])
        elif ops == "re" and fill == (1.0, 1.0, 0.0):
            r = d["rect"]
            squares.append((r.x0 + r.width / 2, r.y0 + r.height / 2, col))
        elif ops == "cccc" and fill == (0.0, 0.0, 0.0):
            r = d["rect"]
            nodes.append((r.x0 + r.width / 2, r.y0 + r.height / 2))

    print("=" * 78)
    print(path.name)
    for p in poly5[:4]:
        start, tip = p[0], p[3]
        dn_s = min(math.dist(start, n) for n in nodes)
        dn_t = min(math.dist(tip, n) for n in nodes)
        dp_s = min(math.dist(start, q) for q in pipe_pts)
        dq = min((math.dist(tip, (x, y)), (x, y)) for x, y, _ in squares)
        print(f"  화살표 시작{tuple(round(v,1) for v in start)}"
              f" 끝{tuple(round(v,1) for v in tip)}"
              f" | 시작→노드 {dn_s:.2f} 시작→관로점 {dp_s:.2f}"
              f" 끝→노드 {dn_t:.2f} 끝→노랑사각 {dq[0]:.2f}")
    for x, y, col in squares[:4]:
        dn = min(math.dist((x, y), n) for n in nodes)
        dp = min(math.dist((x, y), q) for q in pipe_pts)
        d5 = min(min(math.dist((x, y), pt) for pt in p) for p in poly5)
        print(f"  노랑사각({x:.1f},{y:.1f}) 테두리{col}"
              f" → 노드 {dn:.2f} 관로점 {dp:.2f} 화살표점 {d5:.2f}")

    # 첫 화살표 둘레를 크게 오려 본다
    if poly5:
        c = poly5[0][3]
        r = fitz.Rect(c[0] - 45, c[1] - 45, c[0] + 45, c[1] + 45)
        page.get_pixmap(clip=r, dpi=900).save(OUT / f"{tag}_노즐확대.png")
    if squares:
        x, y, _ = squares[1] if len(squares) > 1 else squares[0]
        r = fitz.Rect(x - 45, y - 45, x + 45, y + 45)
        page.get_pixmap(clip=r, dpi=900).save(OUT / f"{tag}_사각확대.png")
    page.get_pixmap(dpi=200).save(OUT / f"{tag}_전면.png")
    doc.close()


if __name__ == "__main__":
    for i, t in enumerate(sys.argv[1:]):
        main(pathlib.Path(t), f"ref{i + 1}")
