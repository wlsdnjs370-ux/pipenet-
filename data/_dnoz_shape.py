# -*- coding: utf-8 -*-
"""노즐 기호를 모델 좌표로 잰다 — 스텁은 데이터, 머리는 상수인지 갈린다.

앞선 실측에서 `lllll` 5획 도형의 전체 길이 변동계수가 0.202 로 컸다. 스텁 길이는
SDF 의 `@` 노드 좌표가 정하는 값(데이터)이라 흔들리는 게 당연하고, 머리 삼각형만
상수일 수 있다. 둘을 갈라서 잰다.
"""
from __future__ import annotations

import math
import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402

LIB = ROOT / "data" / "reference_library"


def harvest(pdf: pathlib.Path):
    doc = fitz.open(pdf)
    dots, glyphs = [], []
    for d in doc[0].get_drawings():
        col = None if d["color"] is None else tuple(round(v, 3) for v in d["color"])
        fill = None if d["fill"] is None else tuple(round(v, 3) for v in d["fill"])
        ops = "".join(it[0] for it in d["items"])
        if ops == "cccc" and fill == (0.0, 0.0, 0.0):
            r = d["rect"]
            dots.append((r.x0 + r.width / 2, r.y0 + r.height / 2))
        elif ops == "lllll" and col == (0.0, 0.0, 0.0):
            s = [((it[1].x, it[1].y), (it[2].x, it[2].y)) for it in d["items"]]
            glyphs.append(([s[0][0]] + [q for _, q in s], fill))
    doc.close()
    return dots, glyphs


def main(limit: int) -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    head, halfw, stub, closed, filled = [], [], [], 0, 0
    files = 0
    for sdf in sorted(LIB.rglob("*.sdf")):
        pdf = pdfs.get(sdf.stem)
        if pdf is None:
            continue
        try:
            model = load_display_model(sdf)
        except Exception:                                  # noqa: BLE001
            continue
        real = [n for n in model.nodes if not n.virtual]
        if len(real) < 4 or not model.nozzles:
            continue
        dots, glyphs = harvest(pdf)
        if not glyphs or abs(len(real) - len(dots)) > 1:
            continue
        mspan = max(max(n.x for n in real) - min(n.x for n in real),
                    max(n.y for n in real) - min(n.y for n in real))
        dspan = max(max(x for x, _ in dots) - min(x for x, _ in dots),
                    max(y for _, y in dots) - min(y for _, y in dots))
        if mspan <= 0 or dspan <= 0:
            continue
        scale = dspan / mspan
        kept = 0
        for pts, fill in glyphs:
            # 노즐 기호는 관로 위 노드에서 시작한다. 시작점이 검은 점과 맞아떨어지지
            # 않으면 다른 도형이다 — 이 조건이 없으면 통계가 오염된다.
            if min(math.dist(pts[0], d) for d in dots) > 0.5:
                continue
            kept += 1
            if math.dist(pts[1], pts[5]) < 0.05:
                closed += 1
            if fill is not None:
                filled += 1
            stub.append(math.dist(pts[0], pts[1]) / scale)
            head.append(math.dist(pts[1], pts[3]) / scale)
            halfw.append(math.dist(pts[2], pts[4]) / 2 / scale)
        if kept:
            files += 1
        if files >= limit:
            break

    print(f"도면 {files}장  노즐 기호 {len(head)}개")
    print(f"  머리 뒤끝이 닫힌 것 {closed} / {len(head)}   채움이 있는 것 {filled}")
    for name, vals in (("스텁", stub), ("머리 길이", head), ("머리 반폭", halfw)):
        vals = sorted(vals)
        print(f"  {name:6s} 중앙 {statistics.median(vals):8.4f}"
              f"  최소 {vals[0]:7.4f}  최대 {vals[-1]:8.4f}"
              f"  변동계수 {statistics.pstdev(vals) / statistics.mean(vals):.3f}")
    ratio = [w / h for w, h in zip(halfw, head)]
    print(f"  반폭/길이 중앙 {statistics.median(ratio):.4f}"
          f"  변동계수 {statistics.pstdev(ratio) / statistics.mean(ratio):.3f}"
          f"   → 꼭지각 {2 * math.degrees(math.atan(statistics.median(ratio))):.2f}°")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 60)
