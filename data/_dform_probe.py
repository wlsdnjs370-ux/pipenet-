# -*- coding: utf-8 -*-
"""참조 PIPENET 지면의 판면(테두리·표제란·글씨)을 pt 로 실측한다.

인자 없이 부르면 참조 3장을 잰다. 경로를 주면 그 파일만 잰다.
"""
import pathlib
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent


def segments(page):
    """직선 획만 (수평/수직) 로 갈라 돌려준다. 사각형은 네 변으로 편다."""
    hs, vs = [], []
    for d in page.get_drawings():
        w = d.get("width") or 0.0
        pairs = []
        for item in d["items"]:
            if item[0] == "l":
                pairs.append((item[1], item[2]))
            elif item[0] == "re":
                r = item[1]
                c = [(r.x0, r.y0), (r.x1, r.y0), (r.x1, r.y1), (r.x0, r.y1)]
                pairs += [(fitz.Point(c[i]), fitz.Point(c[(i + 1) % 4])) for i in range(4)]
        for p, q in pairs:
            if abs(p.y - q.y) < 0.05 and abs(p.x - q.x) > 0.05:
                hs.append((round(p.y, 2), round(min(p.x, q.x), 2), round(max(p.x, q.x), 2),
                           round(w, 3)))
            elif abs(p.x - q.x) < 0.05 and abs(p.y - q.y) > 0.05:
                vs.append((round(p.x, 2), round(min(p.y, q.y), 2), round(max(p.y, q.y), 2),
                           round(w, 3)))
    return sorted(set(hs)), sorted(set(vs))


def report(path):
    doc = fitz.open(str(path))
    page = doc[0]
    W, H = page.rect.width, page.rect.height
    print(f"\n{'=' * 78}\n{path.name}   {W:.2f} x {H:.2f} pt\n{'=' * 78}")

    hs, vs = segments(page)
    # 판면은 지면 아래쪽 1/4 과 테두리에 몰려 있다. 망 자체의 획은 뺀다.
    frame_x = sorted({x for x, y0, y1, _ in vs if y1 - y0 > H * 0.8})
    frame_y = sorted({y for y, x0, x1, _ in hs if x1 - x0 > W * 0.8})
    print(f"-- 테두리: x {frame_x}  y {frame_y}")
    if frame_x and frame_y:
        print(f"   여백  좌 {frame_x[0]:.2f} · 우 {W - frame_x[-1]:.2f} · "
              f"상 {frame_y[0]:.2f} · 하 {H - frame_y[-1]:.2f}")

    # 표제란 = 테두리 하단에 붙은 가장 넓은 가로 괘선 뭉치.
    cand = [(y, x0, x1, w) for y, x0, x1, w in hs if y > H * 0.8 and x1 - x0 > W * 0.3]
    if cand:
        left = min(x0 for _, x0, _, _ in cand)
        top = min(y for y, _, _, _ in cand)
        right = max(x1 for _, _, x1, _ in cand)
        print(f"-- 표제란 상자: x {left:.2f}~{right:.2f} ({right - left:.2f})  "
              f"y {top:.2f}~{H - (H - max(y for y, _, _, _ in cand)):.2f}")
        print(f"   지면 우/하 기준: 오른쪽 {W - right:.2f} · 아래 "
              f"{H - max(y for y, _, _, _ in cand):.2f}")
        rows = [y for y, _, _, _ in cand]
        print(f"   가로 괘선 {len(rows)}개 y={[round(r, 2) for r in rows]}")
        print(f"   행 높이 {[round(b - a, 2) for a, b in zip(rows, rows[1:])]}")

    print("-- 표제란 영역 획 전부 (y > 700)")
    for y, x0, x1, w in hs:
        if y > 700:
            print(f"   가로 y={y:7.2f}  x {x0:7.2f}~{x1:7.2f}  길이 {x1 - x0:7.2f}  w={w}")
    for x, y0, y1, w in vs:
        if y1 > 700:
            print(f"   세로 x={x:7.2f}  y {y0:7.2f}~{y1:7.2f}  길이 {y1 - y0:7.2f}  w={w}")

    print("-- 글씨 (크기별 개수)")
    sizes = {}
    furniture = []
    for blk in page.get_text("dict")["blocks"]:
        for line in blk.get("lines", []):
            for sp in line["spans"]:
                t = sp["text"].strip()
                if not t:
                    continue
                key = (round(sp["size"], 2), sp["font"])
                sizes[key] = sizes.get(key, 0) + 1
                if sp["bbox"][1] > 700 or sp["bbox"][1] < H * 0.45:
                    furniture.append((round(sp["size"], 2), sp["font"],
                                      round(sp["bbox"][0], 2), round(sp["bbox"][1], 2), t))
    for (size, font), n in sorted(sizes.items()):
        print(f"   {size:6.2f}pt  {font:22s}  {n:4d}개")
    print("-- 판면 글씨 (표제란 + 상단 주기)")
    for size, font, x, y, t in sorted(furniture, key=lambda r: r[3]):
        print(f"   {size:6.2f}pt ({x:7.2f},{y:7.2f}) {font:20s} {t[:44]!r}")
    doc.close()


targets = ([pathlib.Path(a) for a in sys.argv[1:]]
           or sorted(ROOT.glob("공동주택_스프링클러*.pdf")))
for p in targets:
    report(p)
