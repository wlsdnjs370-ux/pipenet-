# -*- coding: utf-8 -*-
"""갈매기표를 관로 위 어디에 몇 개나 놓는가 — 호길이 비율로 잰다."""
from __future__ import annotations

import collections
import math
import pathlib
import statistics
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

import fitz  # noqa: E402

LIB = ROOT / "data" / "reference_library"
PURE = {(1.0, 0.0, 0.0), (1.0, 0.676, 0.0), (0.0, 1.0, 0.0),
        (0.0, 1.0, 1.0), (0.0, 0.0, 1.0), (1.0, 0.0, 1.0)}


def split(pdf: pathlib.Path):
    doc = fitz.open(pdf)
    chev, pipes = [], []
    for d in doc[0].get_drawings():
        col = None if d["color"] is None else tuple(round(v, 3) for v in d["color"])
        s = [((it[1].x, it[1].y), (it[2].x, it[2].y))
             for it in d["items"] if it[0] == "l"]
        if col not in PURE or not s:
            continue
        if len(s) == 2:
            a, b, c = s[0][0], s[0][1], s[1][1]
            l1, l2 = math.dist(a, b), math.dist(b, c)
            ang = math.degrees(abs(math.atan2(a[1] - b[1], a[0] - b[0])
                                   - math.atan2(c[1] - b[1], c[0] - b[0])))
            ang = min(ang, 360 - ang)
            if 40 < ang < 70 and abs(l1 - l2) < 0.25 * max(l1, l2) and max(l1, l2) < 3:
                chev.append((col, a, b, c))
                continue
        pipes.append((col, [s[0][0]] + [q for _, q in s]))
    doc.close()
    return chev, pipes


fracs, per_pipe, tip_forward, offsets = [], [], [], []
files = 0
for pdf in sorted(LIB.rglob("*3.유량.pdf"))[:60]:
    chev, pipes = split(pdf)
    if not chev or not pipes:
        continue
    files += 1
    hits = collections.Counter()
    for col, a, b, c in chev:
        axis = ((b[0] - (a[0] + c[0]) / 2), (b[1] - (a[1] + c[1]) / 2))
        best = None
        for pi, (pcol, pts) in enumerate(pipes):
            if pcol != col:
                continue
            acc, total = 0.0, sum(math.dist(p, q) for p, q in zip(pts, pts[1:]))
            for p, q in zip(pts, pts[1:]):
                L = math.dist(p, q)
                if L < 1e-9:
                    continue
                t = max(0.0, min(1.0, ((b[0] - p[0]) * (q[0] - p[0])
                                       + (b[1] - p[1]) * (q[1] - p[1])) / (L * L)))
                proj = (p[0] + t * (q[0] - p[0]), p[1] + t * (q[1] - p[1]))
                dd = math.dist(b, proj)
                # 관로 진행 방향과 갈매기 축이 같은 쪽을 향해야 그 관로의 것이다
                dot = ((q[0] - p[0]) * axis[0] + (q[1] - p[1]) * axis[1]) / L
                if best is None or dd < best[0]:
                    best = (dd, (acc + t * L) / total, pi, dot, total)
                acc += L
        if best and best[0] < 0.4:
            fracs.append(best[1])
            hits[best[2]] += 1
            tip_forward.append(best[3] > 0)
            offsets.append(best[0])
    for pi in range(len(pipes)):
        per_pipe.append(hits.get(pi, 0))

print(f"도면 {files}장  갈매기 {len(fracs)}개")
print(f"  관로 위치(호길이 비) 중앙 {statistics.median(fracs):.4f}"
      f"  4분위 {statistics.quantiles(fracs, n=4)}")
print("  0.05 폭 히스토그램:",
      collections.Counter(round(f * 20) / 20 for f in fracs).most_common(8))
print("  관로당 갈매기 수:", collections.Counter(per_pipe).most_common(6))
print(f"  꼭짓점이 진행 방향 쪽: {sum(tip_forward)} / {len(tip_forward)}")
print(f"  관로에서 떨어진 거리 중앙 {statistics.median(offsets):.3f}pt")
