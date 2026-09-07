# -*- coding: utf-8 -*-
"""[설명용] 설계면적을 고르는 세 가지 자가 **실제로 어떤 헤드를 집는지** 지도.

말로는 «가지관 기반» 과 «공간 직사각형» 의 차이가 안 와닿는다. 같은 도면
같은 앵커에서 각각 무엇을 집는지 찍어 보면 한눈에 갈린다.

  A  앵커(가장 불리한 헤드)
  #  지금 방식이 집은 것 (앵커에서 «배관거리» 로 가까운 K개)
  O  공간 직사각형이 집을 것
  *  둘 다 집은 것
  .  그 밖의 헤드
"""
from __future__ import annotations

import heapq
import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(_ROOT)
sys.path.insert(0, _ROOT)

from routes.module_f.common import _boot          # noqa: E402

_boot()

from services.cad_import.design.worst import worst_k_heads   # noqa: E402
from services.cad_import.edit.session import EditSession     # noqa: E402


def _dij(pts, edges, seeds):
    adj: dict = {}
    for a, b in edges:
        adj.setdefault(a, []).append(b)
        adj.setdefault(b, []).append(a)
    dist: dict = {}
    pq: list = []
    for s in seeds:
        if isinstance(s, int) and 0 <= s < len(pts):
            dist[s] = 0.0
            heapq.heappush(pq, (0.0, s))
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, float("inf")):
            continue
        for v in adj.get(u, ()):
            nd = d + math.dist(pts[u], pts[v])
            if nd < dist.get(v, float("inf")):
                dist[v] = nd
                heapq.heappush(pq, (nd, v))
    return dist


def _rect_pick(far, xy, anchor, k):
    """공간 직사각형 — 앵커를 품는 최소 사각형을 K개 담길 때까지 넓힌다."""
    a = xy[anchor]
    order = sorted(far, key=lambda h: max(abs(xy[h][0] - a[0]),
                                          abs(xy[h][1] - a[1])))
    return order[:k]


def main() -> int:
    key = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")
    k = int(os.environ.get("MF_K", "30"))
    es = EditSession.open(key, out_dir=None, load_saved=True, use_cache=True)
    b = es.board
    src = _dij(b.pts, b.edges, [list(b.sources)[0]])
    far, node = {}, {}
    for hi, nodes in enumerate(b.hnodes):
        reach = [n for n in nodes if n in src]
        if reach:
            n0 = min(reach, key=lambda n: src[n])
            node[hi], far[hi] = n0, src[n0]
    xy = {hi: (float(b.disks[hi][0]), float(b.disks[hi][1]))
          for hi in far if hi < len(b.disks)}
    far = {hi: v for hi, v in far.items() if hi in xy}

    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k, source_index=0)
    anchor = w["worst_head"]
    cur = set(w["heads"])
    rect = set(_rect_pick(far, xy, anchor, k))

    a = xy[anchor]
    R = 30000.0                       # 앵커 둘레 30 m 만 그린다
    near = [h for h in far
            if abs(xy[h][0] - a[0]) <= R and abs(xy[h][1] - a[1]) <= R]
    W, H = 78, 26
    x0 = a[0] - R
    y0 = a[1] - R
    grid = [[" "] * W for _ in range(H)]
    for h in near:
        cx = int((xy[h][0] - x0) / (2 * R) * (W - 1))
        cy = int((1 - (xy[h][1] - y0) / (2 * R)) * (H - 1))
        cx = max(0, min(W - 1, cx))
        cy = max(0, min(H - 1, cy))
        ch = "."
        if h in cur and h in rect:
            ch = "*"
        elif h in cur:
            ch = "#"
        elif h in rect:
            ch = "O"
        if h == anchor:
            ch = "A"
        # 이미 더 «센» 글자가 있으면 안 덮는다
        cur_ch = grid[cy][cx]
        if cur_ch == "A" or (cur_ch in "*#O" and ch == "."):
            continue
        grid[cy][cx] = ch

    print(f"[{key}]  앵커 h{anchor} ({far[anchor] / 1000:.1f} m) 둘레 "
          f"±{R / 1000:.0f} m · 헤드 {len(near)}개 · K={k}")
    print("  A 앵커 · # 지금(배관거리) · O 공간 직사각형 · * 둘 다 · . 나머지\n")
    for row in grid:
        print("   " + "".join(row).rstrip())

    def stat(tag, sel):
        d = sorted((far[h] for h in sel), reverse=True)
        P = [xy[h] for h in sel]
        span = max(math.dist(p, q) for p in P for q in P)
        ys = sorted({round(xy[h][1] / 1000) for h in sel})
        # 가지관은 대개 나란하다 — y 값이 몇 줄로 뭉치는지 센다
        lines, prev = 1, ys[0]
        for v in ys[1:]:
            if v - prev > 2:
                lines += 1
            prev = v
        print(f"  {tag:<18} 유하 {d[0] / 1000:>6.1f}~{d[-1] / 1000:>6.1f} m"
              f" (평균 {sum(d) / len(d) / 1000:>6.1f})"
              f" · 공간폭 {span / 1000:>5.1f} m · 가로줄 약 {lines}개")

    print()
    stat("지금(배관거리)", list(cur))
    stat("공간 직사각형", list(rect))
    print(f"\n  겹침 {len(cur & rect)}개 / {k}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
