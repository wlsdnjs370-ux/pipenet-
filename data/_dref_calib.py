# -*- coding: utf-8 -*-
"""SDF 모델 좌표와 PIPENET PDF 페이지 좌표의 배율을 맞춰 기호 치수를 확정한다.

배율은 노드 점 무리의 크기 대 모델 노드 무리의 크기로 잡는다. 그러면 기호가
"페이지에 고정"인지 "모델 단위에 고정"인지 둘 중 어느 쪽인지 갈린다.
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
PURE = {(1.0, 0.0, 0.0), (1.0, 0.676, 0.0), (0.0, 1.0, 0.0),
        (0.0, 1.0, 1.0), (0.0, 0.0, 1.0), (1.0, 0.0, 1.0)}


def harvest(pdf: pathlib.Path):
    doc = fitz.open(pdf)
    page = doc[0]
    dots, chev, stubs, pipes = [], [], [], []
    for d in page.get_drawings():
        col = None if d["color"] is None else tuple(round(v, 3) for v in d["color"])
        fill = None if d["fill"] is None else tuple(round(v, 3) for v in d["fill"])
        s = [((it[1].x, it[1].y), (it[2].x, it[2].y))
             for it in d["items"] if it[0] == "l"]
        ops = "".join(it[0] for it in d["items"])
        if ops == "cccc" and fill == (0.0, 0.0, 0.0):
            r = d["rect"]
            dots.append((r.x0 + r.width / 2, r.y0 + r.height / 2, r.width))
        elif ops == "lllll" and col == (0.0, 0.0, 0.0) and len(s) == 5:
            pts = [s[0][0]] + [q for _, q in s]
            stubs.append((math.dist(pts[0], pts[3]), math.dist(pts[1], pts[3]),
                          math.dist(pts[2], pts[4]) / 2))
        elif col in PURE and len(s) == 2:
            a, b, c = s[0][0], s[0][1], s[1][1]
            l1, l2 = math.dist(a, b), math.dist(b, c)
            ang = math.degrees(abs(math.atan2(a[1] - b[1], a[0] - b[0])
                                   - math.atan2(c[1] - b[1], c[0] - b[0])))
            ang = min(ang, 360 - ang)
            if 40 < ang < 70 and abs(l1 - l2) < 0.25 * max(l1, l2) and max(l1, l2) < 3:
                chev.append(((l1 + l2) / 2, ang, d["width"], col, b))
            else:
                pipes.append(([a, b, c], d["width"], col))
        elif col in PURE and s:
            pipes.append(([s[0][0]] + [q for _, q in s], d["width"], col))
    doc.close()
    return dots, chev, stubs, pipes


def main(limit: int) -> None:
    pdfs = {p.name.rsplit(" 3.유량.pdf", 1)[0]: p for p in LIB.rglob("*3.유량.pdf")}
    rows = []
    for sdf in sorted(LIB.rglob("*.sdf")):
        pdf = pdfs.get(sdf.stem)
        if pdf is None:
            continue
        try:
            model = load_display_model(sdf)
        except Exception as exc:                       # noqa: BLE001
            print("  파싱 실패", sdf.stem[:40], exc)
            continue
        dots, chev, stubs, pipes = harvest(pdf)
        if len(dots) < 4 or not chev:
            continue
        # 검은 점은 실노드에만 찍힌다 — 노즐 방출단('@')과 텍스트는 빼고 잰다.
        real = [n for n in model.nodes if not n.virtual]
        if len(real) < 4:
            continue
        mspan = max(max(n.x for n in real) - min(n.x for n in real),
                    max(n.y for n in real) - min(n.y for n in real))
        dspan = max(max(x for x, _, _ in dots) - min(x for x, _, _ in dots),
                    max(y for _, y, _ in dots) - min(y for _, y, _ in dots))
        if mspan <= 0 or dspan <= 0:
            continue
        # 점 하나하나가 실노드와 맞아떨어질 때만 배율을 믿는다. 어긋나면 배율이
        # 틀리고, 틀린 배율로 잰 치수는 통계를 통째로 오염시킨다.
        if abs(len(real) - len(dots)) > 1:
            continue
        scale = dspan / mspan
        wing = statistics.median(w for w, *_ in chev)
        rows.append(dict(
            name=sdf.stem, nodes=len(real), dots=len(dots),
            scale=scale,
            wing_pt=wing,
            wing_frac=wing / dspan,               # 도면 크기 대비 — 우리 상수와 같은 뜻
            wing_model=wing / scale,
            stub_model=(statistics.median(s[0] for s in stubs) / scale) if stubs else None,
            dot_model=statistics.median(d for *_, d in dots) / scale,
            ang=statistics.median(a for _, a, *_ in chev),
            dot_pt=statistics.median(d for *_, d in dots),
            dot_frac=statistics.median(d for *_, d in dots) / dspan,
            stub_pt=statistics.median(s[0] for s in stubs) if stubs else None,
            stub_frac=(statistics.median(s[0] for s in stubs) / dspan) if stubs else None,
            head_ratio=statistics.median(s[1] / s[0] for s in stubs) if stubs else None,
            half_ratio=statistics.median(s[2] / s[0] for s in stubs) if stubs else None,
            lw=statistics.median(c[2] for c in chev),
        ))
        if len(rows) >= limit:
            break

    print(f"짝 {len(rows)} 세트")
    for key in ("wing_pt", "wing_frac", "wing_model", "ang", "dot_frac",
                "dot_model", "stub_frac", "stub_model", "head_ratio",
                "half_ratio", "lw"):
        vals = [r[key] for r in rows if r[key] is not None]
        if not vals:
            continue
        vals.sort()
        print(f"  {key:11s} 중앙 {statistics.median(vals):.5f}"
              f"  최소 {vals[0]:.5f}  최대 {vals[-1]:.5f}"
              f"  변동계수 {statistics.pstdev(vals) / statistics.mean(vals):.3f}")
    print("  노드 수 일치:",
          sum(1 for r in rows if r["nodes"] == r["dots"]), "/", len(rows))


if __name__ == "__main__":
    main(int(sys.argv[1]) if len(sys.argv) > 1 else 40)
