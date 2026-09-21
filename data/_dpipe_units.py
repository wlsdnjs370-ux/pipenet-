# -*- coding: utf-8 -*-
"""관로 획 굵기의 기준을 고른다 — 종이 pt 고정인가, 모델 단위 고정인가.

한 장에 굵기가 한 종류뿐이라는 건 이미 봤다(_dpipe_width). 그 하나가 어느 자에
고정돼 있는지 본다. 축척은 노드 점 중심 span 으로 잡는다.
"""
from __future__ import annotations

import collections
import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402

LIB = ROOT / "data" / "reference_library"


def spread(name: str, vals: list[float]) -> None:
    v = sorted(vals)
    q = statistics.quantiles(v, n=4)
    print(f"  {name:12s} 중앙 {statistics.median(v):8.4f}  4분위 {q[0]:8.4f}/{q[2]:8.4f}"
          f"  변동계수 {statistics.pstdev(v) / statistics.mean(v):.3f}")


def main(limit: int) -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    pt, units = [], []
    files = 0
    for sdf in sorted(LIB.rglob("*.sdf")):
        pdf = pdfs.get(sdf.stem)
        if pdf is None:
            continue
        try:
            model = load_display_model(sdf)
        except Exception:                                   # noqa: BLE001
            continue
        real = [n for n in model.nodes if not n.virtual]
        if len(real) < 4:
            continue
        doc = fitz.open(pdf)
        dots, widths = [], collections.Counter()
        for d in doc[0].get_drawings():
            fill = None if d["fill"] is None else tuple(round(v, 3) for v in d["fill"])
            if "".join(it[0] for it in d["items"]) == "cccc" and fill == (0.0, 0.0, 0.0):
                dots.append(d["rect"])
            if d["color"] is not None and d["width"] not in (None, 0):
                if set(i[0] for i in d["items"]) == {"l"}:
                    widths[round(d["width"], 3)] += len(d["items"])
        doc.close()
        if abs(len(real) - len(dots)) > 1 or len(dots) < 4 or not widths:
            continue
        mspan = max(max(n.x for n in real) - min(n.x for n in real),
                    max(n.y for n in real) - min(n.y for n in real))
        cx = [r.x0 + r.width / 2 for r in dots]
        cy = [r.y0 + r.height / 2 for r in dots]
        dspan = max(max(cx) - min(cx), max(cy) - min(cy))
        if mspan <= 0 or dspan <= 0:
            continue
        scale = dspan / mspan
        w = widths.most_common(1)[0][0]
        files += 1
        pt.append(w)
        units.append(w / scale)
        if files >= limit:
            break

    print(f"도면 {files}장")
    spread("① 종이 pt", pt)
    spread("② 모델 단위", units)


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 60)
