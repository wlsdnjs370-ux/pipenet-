# -*- coding: utf-8 -*-
"""PIPENET 원본 PDF 역산 — §9-2 페이지 분할 / §9-4 표시 프리셋 (일회성)."""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
LIB = ROOT / "data" / "reference_library"

# ── §9-2: ISO 도면은 몇 쪽인가 (전수) ────────────────────────────────────────
iso = [p for p in LIB.rglob("*.pdf") if re.search(r"[23]\.(압력|유량)\.pdf$", p.name)]
pages, sizes, rots, headers = Counter(), Counter(), Counter(), Counter()
for p in iso:
    try:
        r = PdfReader(str(p))
    except Exception as exc:
        print("EXC", p.name, exc)
        continue
    pages[len(r.pages)] += 1
    for pg in r.pages:
        b = pg.mediabox
        sizes[(round(float(b.width) / 72 * 25.4), round(float(b.height) / 72 * 25.4))] += 1
        rots[pg.get("/Rotate", 0)] += 1
    m = re.search(r"Page (\d+) of (\d+)", r.pages[0].extract_text() or "")
    headers[m.groups() if m else None] += 1
print(f"ISO PDF {len(iso)}개")
print("  쪽수 분포:", pages.most_common())
print("  용지(mm):", sizes.most_common())
print("  회전:", rots.most_common())
print("  머리글 'Page a of b':", headers.most_common(5))

# ── §9-4: 두 프리셋이 무엇을 켜는가 ─────────────────────────────────────────
LEGEND = re.compile(r"^[<>]\s*-?[\d.]+")


def legend_of(path: Path) -> list[str]:
    text = PdfReader(str(path)).pages[0].extract_text() or ""
    lines = [t.strip() for t in text.splitlines() if t.strip()]
    out, i = [], 0
    while i < len(lines):
        # 범례는 '이름' / '(단위)' / 밴드 줄 이 잇달아 나온다
        if i + 1 < len(lines) and lines[i + 1].startswith("("):
            bands = []
            j = i + 2
            while j < len(lines) and LEGEND.match(lines[j]):
                bands.append(lines[j])
                j += 1
            if bands:
                out.append(f"{lines[i]} {lines[i + 1]} :: " + " | ".join(bands))
                i = j
                continue
        i += 1
    return out


print()
print("=== 프리셋 (범례 이름으로 역산)")
combos, bands = Counter(), Counter()
for p in iso:
    try:
        legend = legend_of(p)
    except Exception:
        continue
    names = tuple(l.split(" ::")[0] for l in legend)
    combos[("압력" if "압력" in p.name else "유량", names)] += 1
    for line in legend:
        head, rest = line.split(" :: ")
        bands[(head, len(rest.split(" | ")[0].split()) + len(rest.split(" | ")[1].split()))] += 1
for (tag, names), n in combos.most_common(12):
    print(f"  [{tag}] {n:4d}건  {names}")
print()
print("밴드 개수:", bands.most_common(8))
