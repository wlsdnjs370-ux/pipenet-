# -*- coding: utf-8 -*-
"""참조 도면 한 장을 PIPENET 원본과 우리 렌더러로 나란히 뽑는다."""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402
from core.d_iso_renderer import render_iso  # noqa: E402

LIB = ROOT / "data" / "reference_library"
OUT = ROOT / "data" / "_dcmp"
OUT.mkdir(exist_ok=True)


def main() -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    for sdf in sorted(LIB.rglob("*.sdf")):
        ref = pdfs.get(sdf.stem)
        if ref is None:
            continue
        model = load_display_model(sdf)
        if len({p.bore_m for p in model.pipes if p.bore_m}) < 4 or len(model.pipes) < 40:
            continue
        mine = OUT / "ours.pdf"
        render_iso(model, mine, link_item="None", node_item="None")
        for tag, path in (("pipenet", ref), ("ours", mine)):
            doc = fitz.open(path)
            doc[0].get_pixmap(dpi=200).save(OUT / f"{tag}.png")
            doc.close()
        print(sdf.stem, "관로", len(model.pipes))
        return
    print("짝지을 도면을 못 찾았다")


if __name__ == "__main__":
    main()
