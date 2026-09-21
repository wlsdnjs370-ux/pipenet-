# -*- coding: utf-8 -*-
"""도면 주기(한글 글씨) 높이가 SDF 의 typesize 를 따라가는지 본다.

렌더러의 _note_size 는 typesize 를 print-font(48) 로 나눈 상대비로 주기
크기를 정한다. 그 환산이 맞다면, 실측한 주기 높이를 typesize 로 나눈 몫이
도면마다 같아야 한다.
"""
from __future__ import annotations

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


def main(limit: int) -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    val_model, note_model, ratios, quotients = [], [], [], []
    sheets = 0
    for sdf in sorted(LIB.rglob("*.sdf")):
        pdf = pdfs.get(sdf.stem)
        if pdf is None:
            continue
        try:
            model = load_display_model(sdf)
        except Exception:                                  # noqa: BLE001
            continue
        doc = fitz.open(pdf)
        page = doc[0]
        W, H = page.rect.width, page.rect.height
        dots = [d["rect"] for d in page.get_drawings()
                if "".join(it[0] for it in d["items"]) == "cccc"
                and d["fill"] is not None and max(d["fill"]) < 0.01]
        spans = [s for b in page.get_text("dict")["blocks"] for ln in b.get("lines", ())
                 for s in ln["spans"] if 0.156 < 1 - s["bbox"][3] / H < 0.930]
        doc.close()
        real = [n for n in model.nodes if not n.virtual]
        if W > H or len(dots) < 4 or len(real) < 4 or abs(len(real) - len(dots)) > 1:
            continue
        mspan = max(max(n.x for n in real) - min(n.x for n in real),
                    max(n.y for n in real) - min(n.y for n in real))
        dspan = max(max(r.x0 for r in dots) - min(r.x0 for r in dots),
                    max(r.y0 for r in dots) - min(r.y0 for r in dots))
        if mspan <= 0 or dspan <= 0:
            continue
        scale = dspan / mspan
        vals = [s["size"] for s in spans if NUM.match(s["text"].strip())]
        notes = [s["size"] for s in spans if not NUM.match(s["text"].strip())
                 and s["text"].strip()]
        sizes = [t.typesize for t in model.texts if t.typesize]
        if len(vals) < 10 or len(notes) < 3 or not sizes:
            continue
        sheets += 1
        v = statistics.median(vals) / scale
        n = statistics.median(notes) / scale
        val_model.append(v)
        note_model.append(n)
        ratios.append(n / v)
        quotients.append(n / statistics.median(sizes))
        if sheets >= limit:
            break

    print(f"주기가 있는 짝 {sheets} 세트")
    for name, vals in (("값 글씨(모델단위)", val_model), ("주기 글씨(모델단위)", note_model),
                       ("주기/값 비", ratios), ("주기/typesize 몫", quotients)):
        vals = sorted(vals)
        print(f"  {name:20s} 중앙 {statistics.median(vals):9.4f}"
              f"  최소 {vals[0]:9.4f}  최대 {vals[-1]:9.4f}"
              f"  변동계수 {statistics.pstdev(vals) / statistics.mean(vals):.3f}")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
