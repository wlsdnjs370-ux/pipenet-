# -*- coding: utf-8 -*-
"""망 위에 적힌 글씨를 잰다 — 크기, 글꼴, 무슨 내용을 얼마나 적는가."""
from __future__ import annotations

import collections
import pathlib
import re
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent

import fitz  # noqa: E402

LIB = ROOT / "data" / "reference_library"

NUM = re.compile(r"^[-+]?\d[\d.,]*$")


def classify(s: str) -> str:
    s = s.strip()
    if not s:
        return "빈칸"
    if NUM.match(s):
        return "숫자"
    if re.match(r"^[A-Za-z][A-Za-z0-9_\-]*$", s):
        return "이름"
    return "기타"


def main(pattern: str, limit: int) -> None:
    sizes = collections.Counter()
    fonts = collections.Counter()
    kinds = collections.Counter()
    per_sheet = []
    files = 0
    for pdf in sorted(LIB.rglob(pattern)):
        doc = fitz.open(pdf)
        page = doc[0]
        W, H = page.rect.width, page.rect.height
        d = page.get_text("dict")
        doc.close()
        if W > H:
            continue
        n = 0
        for block in d["blocks"]:
            for line in block.get("lines", ()):
                for span in line["spans"]:
                    # 망이 차지하는 구간만. 표제란·범례는 뺀다.
                    if not 0.156 < 1 - span["bbox"][3] / H < 0.930:
                        continue
                    sizes[round(span["size"], 2)] += 1
                    fonts[span["font"]] += 1
                    kinds[classify(span["text"])] += 1
                    n += 1
        if n:
            files += 1
            per_sheet.append(n)
        if files >= limit:
            break

    print(f"{pattern} 세로 도면 {files}장, 망 위 글씨 중앙 {statistics.median(per_sheet):.0f}개"
          f" (최소 {min(per_sheet)} 최대 {max(per_sheet)})")
    print(f"  크기: {sizes.most_common(6)}")
    print(f"  글꼴: {fonts.most_common(4)}")
    print(f"  내용: {kinds.most_common(6)}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "*3.유량.pdf",
         int(sys.argv[2]) if len(sys.argv) > 2 else 40)
