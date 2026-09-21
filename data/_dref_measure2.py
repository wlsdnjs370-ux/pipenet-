# -*- coding: utf-8 -*-
"""화살표·노즐 기호의 정확한 모양을 잰다 — 꺾인 각도, 붙은 자리, 크기 비율."""
from __future__ import annotations

import collections
import math
import pathlib
import sys

import fitz

ROOT = pathlib.Path(__file__).resolve().parent.parent
LIB = ROOT / "data" / "reference_library"

PURE = {(1.0, 0.0, 0.0), (1.0, 0.676, 0.0), (0.0, 1.0, 0.0),
        (0.0, 1.0, 1.0), (0.0, 0.0, 1.0), (1.0, 0.0, 1.0)}


def rgb(c):
    return None if c is None else tuple(round(v, 3) for v in c)


def segs(d):
    out = []
    for it in d["items"]:
        if it[0] == "l":
            out.append(((it[1].x, it[1].y), (it[2].x, it[2].y)))
    return out


def main(path: pathlib.Path) -> None:
    doc = fitz.open(path)
    page = doc[0]
    draws = page.get_drawings()
    print("=" * 78)
    print(path.name)

    pipes, chevrons, others = [], [], []
    for d in draws:
        s = segs(d)
        if not s or d["type"] != "s":
            others.append(d)
            continue
        col = rgb(d["color"])
        pts = [s[0][0]] + [b for _, b in s]
        length = sum(math.dist(a, b) for a, b in s)
        if col in PURE and len(s) == 2 and length < 6.0:
            chevrons.append((col, pts, s))
        elif col in PURE:
            pipes.append((col, pts, length))
        else:
            others.append(d)

    print(f"관로 {len(pipes)}  화살표후보 {len(chevrons)}  기타 {len(others)}")

    # 화살표: 꼭짓점이 가운데인가(갈매기) 끝인가, 벌어진 각, 날개 길이
    angles, wings, apex_mid = [], [], 0
    for col, pts, s in chevrons:
        a, b, c = pts[0], pts[1], pts[2]
        v1 = (a[0] - b[0], a[1] - b[1])
        v2 = (c[0] - b[0], c[1] - b[1])
        ang = math.degrees(abs(math.atan2(v1[1], v1[0]) - math.atan2(v2[1], v2[0])))
        ang = min(ang, 360 - ang)
        angles.append(round(ang, 1))
        wings.append(round((math.hypot(*v1) + math.hypot(*v2)) / 2, 3))
        apex_mid += 1
    print("  화살표 벌어진각 분포:", collections.Counter(angles).most_common(6))
    print("  화살표 날개길이:", collections.Counter(wings).most_common(6))
    print("  화살표 색 분포:", collections.Counter(c for c, _, _ in chevrons).most_common())

    # 화살표가 관로 위 어디에 붙는가 — 가장 가까운 관로의 어느 지점인가
    frac = []
    for col, pts, s in chevrons:
        apex = pts[1]
        best = None
        for pcol, ppts, plen in pipes:
            if pcol != col:
                continue
            acc = 0.0
            for i in range(len(ppts) - 1):
                p, q = ppts[i], ppts[i + 1]
                L = math.dist(p, q)
                if L < 1e-9:
                    continue
                t = max(0.0, min(1.0, ((apex[0] - p[0]) * (q[0] - p[0]) +
                                       (apex[1] - p[1]) * (q[1] - p[1])) / (L * L)))
                proj = (p[0] + t * (q[0] - p[0]), p[1] + t * (q[1] - p[1]))
                dist = math.dist(apex, proj)
                if best is None or dist < best[0]:
                    best = (dist, (acc + t * L) / plen, plen)
                acc += L
        if best:
            frac.append((round(best[0], 2), round(best[1], 2), round(best[2], 1)))
    print("  화살표→관로 거리 중앙:",
          sorted(f[0] for f in frac)[len(frac) // 2] if frac else "NA")
    print("  관로상 위치(0~1) 분포:",
          collections.Counter(f[1] for f in frac).most_common(8))
    print("  화살표 날개 / 관로길이 비:",
          collections.Counter(round(w / f[2], 4) for w, f in zip(wings, frac)
                              ).most_common(5) if frac else "NA")

    # 기타 도형: 노즐·기기·노드 후보
    agg = collections.Counter()
    samples = {}
    for d in others:
        ops = "".join(it[0] for it in d["items"])
        r = d["rect"]
        key = (d["type"], ops, rgb(d.get("color")), rgb(d.get("fill")),
               round(max(r.width, r.height), 2))
        agg[key] += 1
        samples.setdefault(key, d)
    for key, n in agg.most_common(14):
        print(f"  기타 {n:4d} type={key[0]:<3} ops={key[1][:14]:<14} "
              f"stroke={key[2]} fill={key[3]} size={key[4]}")

    print("  글자:", [t[4] for t in page.get_text("words")[:0]])
    doc.close()


if __name__ == "__main__":
    for t in sys.argv[1:]:
        main(pathlib.Path(t))
