# -*- coding: utf-8 -*-
"""그래프에 급수원에서 물을 흘린다. 화면 없음.

본체 물채움(`pipeline.flow.pipeline`)과 다른 함수다.
"""
import math
from collections import defaultdict, deque

from services.cad_import.pipeline.expand import gnear, gput, seg_dist

SRC_SNAP = 2500.0   # 급수원 스냅 한계 — 본체와 같은 값


def water(g, sources):
    adj = g.adj()
    egrid = defaultdict(list)
    for (i, j) in g.edges:
        a, b = g.pts[i], g.pts[j]
        steps = max(1, int(math.hypot(b[0] - a[0], b[1] - a[1]) / 1000.0))
        for k in range(steps + 1):
            t = k / steps
            gput(egrid, 1000.0, a[0] + (b[0] - a[0]) * t,
                 a[1] + (b[1] - a[1]) * t, (i, j))
    starts = []
    for src in sources:
        tag, sx, sy = src[0], src[1], src[2]
        best = (1e18, None)
        for (i, j) in set(gnear(egrid, 1000.0, sx, sy, rings=3)):
            d, _t = seg_dist(g.pts[i], g.pts[j], sx, sy)
            if d < best[0]:
                best = (d, (i, j))
        if best[1] is None or best[0] > SRC_SNAP:
            print(f"    ★급수원 {tag}: 스냅 실패 ({best[0]:.0f}mm)")
            continue
        print(f"    급수원 {tag}: 스냅 {best[0]:.0f}mm")
        starts += list(best[1])
    reach, q = set(starts), deque(starts)
    while q:
        u = q.popleft()
        for v in adj[u]:
            if v not in reach:
                reach.add(v)
                q.append(v)
    wet = {(i, j) for (i, j) in g.edges if i in reach and j in reach}
    return reach, wet
