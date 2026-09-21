# -*- coding: utf-8 -*-
"""관로 선 굵기가 관경을 따라가는가 — PIPENET 원본 PDF 의 획 굵기 분포를 본다.

우리 렌더러는 bore 에 따라 선을 굵게 그린다. 근거가 있는지 확인한다. 원본이
굵기를 하나만 쓴다면 그건 우리가 지어낸 표기다.
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


def main(limit: int) -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    per_file_kinds, all_widths = [], collections.Counter()
    bores_per_file = []
    files = 0
    for sdf in sorted(LIB.rglob("*.sdf")):
        pdf = pdfs.get(sdf.stem)
        if pdf is None:
            continue
        try:
            model = load_display_model(sdf)
        except Exception:                                   # noqa: BLE001
            continue
        bores = {p.bore_m for p in model.pipes if p.bore_m}
        if len(bores) < 2:                                  # 관경이 하나면 판별이 안 된다
            continue
        doc = fitz.open(pdf)
        widths = collections.Counter()
        for d in doc[0].get_drawings():
            if d["color"] is None or d["width"] in (None, 0):
                continue
            # 직선 도막만 — 기호는 굵기가 따로일 수 있다
            if set(i[0] for i in d["items"]) == {"l"}:
                widths[round(d["width"], 3)] += len(d["items"])
        doc.close()
        if not widths:
            continue
        files += 1
        all_widths.update(widths)
        per_file_kinds.append(len(widths))
        bores_per_file.append(len(bores))
        if files >= limit:
            break

    print(f"도면 {files}장")
    print(f"  한 장에 쓰인 획 굵기 종류  중앙 {statistics.median(per_file_kinds)}"
          f"  최소 {min(per_file_kinds)}  최대 {max(per_file_kinds)}")
    print(f"  한 장에 있는 관경 종류    중앙 {statistics.median(bores_per_file)}"
          f"  최소 {min(bores_per_file)}  최대 {max(bores_per_file)}")
    print("  전체 굵기 상위:", all_widths.most_common(8))


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 60)
