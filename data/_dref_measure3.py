# -*- coding: utf-8 -*-
"""남은 도형의 정체를 좌표로 확정한다 — 노즐 기호는 무엇인가."""
from __future__ import annotations

import collections
import math
import pathlib
import sys

import fitz

PURE = {(1.0, 0.0, 0.0), (1.0, 0.676, 0.0), (0.0, 1.0, 0.0),
        (0.0, 1.0, 1.0), (0.0, 0.0, 1.0), (1.0, 0.0, 1.0)}


def rgb(c):
    return None if c is None else tuple(round(v, 3) for v in c)


def segs(d):
    return [((it[1].x, it[1].y), (it[2].x, it[2].y))
            for it in d["items"] if it[0] == "l"]


def main(path: pathlib.Path) -> None:
    doc = fitz.open(path)
    page = doc[0]
    draws = page.get_drawings()
    print("=" * 78)
    print(path.name)

    chev, pipes, squares, poly5, circles = [], [], [], [], []
    for d in draws:
        col, fill = rgb(d["color"]), rgb(d["fill"])
        s = segs(d)
        ops = "".join(it[0] for it in d["items"])
        if col in PURE and len(s) == 2:
            a, b, c = s[0][0], s[0][1], s[1][1]
            l1, l2 = math.dist(a, b), math.dist(b, c)
            ang = math.degrees(abs(math.atan2(a[1] - b[1], a[0] - b[0])
                                   - math.atan2(c[1] - b[1], c[0] - b[0])))
            ang = min(ang, 360 - ang)
            if 40 < ang < 70 and abs(l1 - l2) < 0.25 * max(l1, l2) and max(l1, l2) < 3:
                chev.append((col, a, b, c, (l1 + l2) / 2, ang))
                continue
        if col in PURE and s:
            pipes.append((col, [s[0][0]] + [q for _, q in s]))
        elif ops == "re" and fill == (1.0, 1.0, 0.0):
            squares.append((d["rect"], col))
        elif ops == "lllll" and col == (0.0, 0.0, 0.0):
            poly5.append([s[0][0]] + [q for _, q in s])
        elif ops == "cccc" and fill == (0.0, 0.0, 0.0):
            circles.append((d["rect"].x0 + d["rect"].width / 2,
                            d["rect"].y0 + d["rect"].height / 2, d["rect"].width))

    print(f"  관로 {len(pipes)}  갈매기 {len(chev)}  노랑사각 {len(squares)}"
          f"  5획검정 {len(poly5)}  검은점 {len(circles)}")

    # 갈매기가 가리키는 쪽 — 꼭짓점이 앞인가 뒤인가. 관로 방향과 견준다.
    ahead = 0
    for col, a, b, c, wing, ang in chev:
        mid = ((a[0] + c[0]) / 2, (a[1] + c[1]) / 2)
        # 꼭짓점 b 에서 두 날개 중점 mid 로 가는 방향이 진행 반대편이면 b 가 앞이다
        ahead += 1 if math.dist(b, mid) > 0 else 0
    print(f"  갈매기 날개길이 중앙 {sorted(w for *_, w, _ in chev)[len(chev)//2]:.3f}"
          f"  벌어진각 중앙 {sorted(a for *_, a in chev)[len(chev)//2]:.1f}")

    # 5획 검정 도형의 실제 좌표 — 노즐 기호 후보
    for p in poly5[:3]:
        base = p[0]
        print("    5획 상대좌표:",
              [(round(x - base[0], 2), round(y - base[1], 2)) for x, y in p])
    # 노랑사각과 5획 도형이 짝인가
    if squares and poly5:
        for r, col in squares[:3]:
            cx, cy = r.x0 + r.width / 2, r.y0 + r.height / 2
            near = min((math.dist((cx, cy), pt), i)
                       for i, p in enumerate(poly5) for pt in p)
            print(f"    노랑사각({cx:.1f},{cy:.1f}) 크기 {r.width:.2f}x{r.height:.2f}"
                  f" 테두리{col} → 가장 가까운 5획까지 {near[0]:.2f}pt")
    # 노랑사각이 관로 끝점에 붙는가
    if squares and pipes:
        ends = [pt for _, p in pipes for pt in (p[0], p[-1])]
        for r, col in squares[:5]:
            cx, cy = r.x0 + r.width / 2, r.y0 + r.height / 2
            print(f"    노랑사각→관로끝점 최단 {min(math.dist((cx, cy), e) for e in ends):.2f}pt")
    # 검은점이 관로 끝점에 붙는가
    if circles and pipes:
        ends = [pt for _, p in pipes for pt in (p[0], p[-1])]
        ds = sorted(min(math.dist((x, y), e) for e in ends) for x, y, _ in circles)
        print(f"    검은점→관로끝점 최단 중앙 {ds[len(ds)//2]:.2f}pt  최대 {ds[-1]:.2f}pt")
        print(f"    검은점 지름 {collections.Counter(round(w,2) for *_ , w in circles).most_common(3)}")
    doc.close()


if __name__ == "__main__":
    for t in sys.argv[1:]:
        main(pathlib.Path(t))
