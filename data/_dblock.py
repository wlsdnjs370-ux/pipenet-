# -*- coding: utf-8 -*-
"""PIPENET 아래쪽 표제란 블록의 자리를 잰다 — 세로 도면만, 글줄 x·y 와 테두리."""
from __future__ import annotations

import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent

import fitz  # noqa: E402

LIB = ROOT / "data" / "reference_library"


def med(vals):
    return statistics.median(vals) if vals else float("nan")


def main(limit: int) -> None:
    rows: dict[str, list[tuple[float, float, float, float]]] = {}
    sw_x0, sw_x1, sw_y0, sw_y1, sw_w = [], [], [], [], []
    frames, files = [], 0
    for pdf in sorted(LIB.rglob("*3.유량.pdf")):
        doc = fitz.open(pdf)
        page = doc[0]
        W, H = page.rect.width, page.rect.height
        if (W > H) != (len(sys.argv) > 2):
            doc.close()
            continue
        sw = [d["rect"] for d in page.get_drawings()
              if d["fill"] is not None and 3 < d["rect"].width < 30
              and 3 < d["rect"].height < 30
              and tuple(round(v, 3) for v in d["fill"]) != (1.0, 1.0, 1.0)
              and 1 - d["rect"].y1 / H < 0.14]
        if len(sw) < 4:
            doc.close()
            continue
        files += 1
        sw_x0.append(min(r.x0 for r in sw) / W)
        sw_x1.append(max(r.x1 for r in sw) / W)
        sw_y0.append(1 - max(r.y1 for r in sw) / H)
        sw_y1.append(1 - min(r.y0 for r in sw) / H)
        sw_w.append(med([r.width / W for r in sw]))
        # 블록 테두리로 쓸 만한 긴 가로선
        for d in page.get_drawings():
            r = d["rect"]
            if d["color"] is not None and r.height < 1.5 and r.width > W * 0.3 \
                    and 1 - r.y1 / H < 0.2:
                frames.append((r.x0 / W, r.x1 / W, 1 - r.y1 / H))
        for b in page.get_text("blocks"):
            if 1 - b[3] / H >= 0.16:
                continue
            key = b[4].strip().splitlines()[0][:24]
            kind = ("legend-title" if "flow" in key.lower() or "velocity" in key.lower()
                    else "band-label" if key.startswith(("<", ">", "≥")) else "block")
            rows.setdefault(kind, []).append(
                (b[0] / W, b[2] / W, 1 - b[3] / H, 1 - b[1] / H))
        doc.close()
        if files >= limit:
            break

    print(f"세로 도면 {files}장")
    print(f"  범례 칸: x {med(sw_x0):.3f}~{med(sw_x1):.3f}  y {med(sw_y0):.3f}~{med(sw_y1):.3f}"
          f"  칸 너비 {med(sw_w):.4f}")
    for kind, rs in rows.items():
        print(f"  [{kind}] {len(rs)}줄  x0 {med([r[0] for r in rs]):.3f}"
              f"  x1 {med([r[1] for r in rs]):.3f}"
              f"  y {med([r[2] for r in rs]):.3f}~{med([r[3] for r in rs]):.3f}")
    if frames:
        print(f"  아래 20% 안 긴 가로선 {len(frames)}개, y 중앙 {med([f[2] for f in frames]):.3f}"
              f"  x {med([f[0] for f in frames]):.3f}~{med([f[1] for f in frames]):.3f}")
    else:
        print("  아래 20% 안에 블록 테두리로 볼 긴 가로선이 없다")


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 30)
