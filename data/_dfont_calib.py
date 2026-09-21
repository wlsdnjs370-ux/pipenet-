# -*- coding: utf-8 -*-
"""망 위 글씨 크기가 종이에 고정인지 도면 크기를 따라가는지 가른다.

_dref_calib.py 와 같은 배율(검은 노드 점 무리 대 모델 노드 무리)을 쓴다.
같은 배율을 써야 화살표·기호와 같은 잣대로 비교할 수 있다.
"""
from __future__ import annotations

import collections
import pathlib
import re
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402

LIB = ROOT / "data" / "reference_library"
NUM = re.compile(r"^[-+]?\d[\d.,]*$")


def dots_and_spans(pdf: pathlib.Path):
    """검은 노드 점과, 망 위에 적힌 글씨 조각을 함께 걷는다."""
    doc = fitz.open(pdf)
    page = doc[0]
    W, H = page.rect.width, page.rect.height
    dots = []
    for d in page.get_drawings():
        fill = None if d["fill"] is None else tuple(round(v, 3) for v in d["fill"])
        if "".join(it[0] for it in d["items"]) == "cccc" and fill == (0.0, 0.0, 0.0):
            r = d["rect"]
            dots.append((r.x0 + r.width / 2, r.y0 + r.height / 2, r.width))
    spans = [s for b in page.get_text("dict")["blocks"] for ln in b.get("lines", ())
             for s in ln["spans"]]
    doc.close()
    return (W, H), dots, spans


def main(limit: int) -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    rows = []
    fonts = collections.Counter()
    for sdf in sorted(LIB.rglob("*.sdf")):
        pdf = pdfs.get(sdf.stem)
        if pdf is None:
            continue
        try:
            model = load_display_model(sdf)
        except Exception:                                  # noqa: BLE001
            continue
        (W, H), dots, spans = dots_and_spans(pdf)
        if W > H or len(dots) < 4:
            continue
        real = [n for n in model.nodes if not n.virtual]
        if len(real) < 4 or abs(len(real) - len(dots)) > 1:
            continue
        mspan = max(max(n.x for n in real) - min(n.x for n in real),
                    max(n.y for n in real) - min(n.y for n in real))
        dspan = max(max(x for x, _, _ in dots) - min(x for x, _, _ in dots),
                    max(y for _, y, _ in dots) - min(y for _, y, _ in dots))
        if mspan <= 0 or dspan <= 0:
            continue
        scale = dspan / mspan
        # 망 위에 적힌 값 글씨만. 표제란·범례는 그림틀 밖이라 빠진다.
        sizes = [s["size"] for s in spans
                 if 0.156 < 1 - s["bbox"][3] / H < 0.930 and NUM.match(s["text"].strip())]
        if len(sizes) < 10:
            continue
        for s in spans:
            if 0.156 < 1 - s["bbox"][3] / H < 0.930:
                fonts[s["font"]] += 1
        size = statistics.median(sizes)
        rows.append(dict(name=sdf.stem, size_pt=size, size_frac=size / dspan,
                         size_model=size / scale, n=len(sizes)))
        if len(rows) >= limit:
            break

    print(f"짝 {len(rows)} 세트, 글씨 중앙 {statistics.median(r['n'] for r in rows):.0f}개")
    for key in ("size_pt", "size_frac", "size_model"):
        vals = sorted(r[key] for r in rows)
        print(f"  {key:11s} 중앙 {statistics.median(vals):.5f}"
              f"  최소 {vals[0]:.5f}  최대 {vals[-1]:.5f}"
              f"  변동계수 {statistics.pstdev(vals) / statistics.mean(vals):.3f}")
    print(f"  글꼴 {fonts.most_common(4)}")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
