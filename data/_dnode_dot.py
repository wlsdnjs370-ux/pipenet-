# -*- coding: utf-8 -*-
"""노드 점의 크기 기준을 고른다 — 종이 pt 고정인가, 모델 단위 고정인가.

기호 셋(화살표·머리·스텁)은 모두 모델 단위에 붙어 있었다. 노드 점도 같은지
확인한다. 같은 표본을 두 기준으로 재서 변동계수가 작은 쪽이 진짜 기준이다.
"""
from __future__ import annotations

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
        dots = []
        for d in doc[0].get_drawings():
            fill = None if d["fill"] is None else tuple(round(v, 3) for v in d["fill"])
            if "".join(it[0] for it in d["items"]) == "cccc" and fill == (0.0, 0.0, 0.0):
                dots.append(d["rect"])
        doc.close()
        if abs(len(real) - len(dots)) > 1 or len(dots) < 4:
            continue
        mspan = max(max(n.x for n in real) - min(n.x for n in real),
                    max(n.y for n in real) - min(n.y for n in real))
        cx = [r.x0 + r.width / 2 for r in dots]
        cy = [r.y0 + r.height / 2 for r in dots]
        dspan = max(max(cx) - min(cx), max(cy) - min(cy))
        if mspan <= 0 or dspan <= 0:
            continue
        scale = dspan / mspan
        files += 1
        for r in dots:
            pt.append((r.width + r.height) / 2)
            units.append((r.width + r.height) / 2 / scale)
        if files >= limit:
            break

    print(f"도면 {files}장  노드 점 {len(pt)}개")
    spread("① 종이 pt", pt)
    spread("② 모델 단위", units)


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 60)
