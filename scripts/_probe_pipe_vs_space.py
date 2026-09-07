# -*- coding: utf-8 -*-
"""[감사] 배관거리 / 직선거리 비 — «자» 문제인가 «그래프» 문제인가.

앞 감사에서 「공간으로 11.1m 옆인 헤드가 배관으로 111.6m」가 나왔다. 두 가지
설명이 가능하고, 고칠 자리가 완전히 다르다.

  ⑴ 자(尺) 문제 — 옆 가지관은 원래 주관을 돌아가므로 배관거리가 길다.
     그렇다면 설계면적을 배관거리로 고르는 것이 개념부터 안 맞는다.
  ⑵ 그래프 문제 — 있어야 할 연결(주관·크로스메인)이 빠져서 우회가 생겼다.
     그렇다면 자를 바꾸기 전에 그래프를 고쳐야 한다.

가르는 법: 앵커 주변 헤드들의 **배관/직선 비**를 분포로 본다.
정상적인 스프링클러 배치라면 옆 가지관은 2~5배쯤이다. 10배가 넘는 것이
수두룩하면 그래프가 끊겨 있다는 뜻이다.
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


def main() -> int:
    key = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")
    es = EditSession.open(key, out_dir=None, load_saved=True, use_cache=True)
    b = es.board
    src = _dij(b.pts, b.edges, [list(b.sources)[0]])
    node, far = {}, {}
    for hi, nodes in enumerate(b.hnodes):
        reach = [n for n in nodes if n in src]
        if reach:
            n0 = min(reach, key=lambda n: src[n])
            node[hi], far[hi] = n0, src[n0]
    xy = {hi: (float(b.disks[hi][0]), float(b.disks[hi][1]))
          for hi in far if hi < len(b.disks)}
    far = {hi: v for hi, v in far.items() if hi in xy}

    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30,
                      source_index=0)
    anchor = w["worst_head"]
    a_xy = xy[anchor]
    an = _dij(b.pts, b.edges, [node[anchor]])

    print(f"[{key}] 앵커 h{anchor} · 도달 헤드 {len(far)}")

    # 앵커에서 공간으로 가까운 60개를 본다 — 설계면적 후보가 될 자리들이다.
    near = sorted(far, key=lambda h: math.dist(xy[h], a_xy))[:60]
    rows = []
    for h in near:
        s = math.dist(xy[h], a_xy) / 1000.0
        p = an.get(node[h], float("inf")) / 1000.0
        if s < 0.5:
            continue
        rows.append((p / s if p != float("inf") else float("inf"), s, p, h))
    rows.sort()
    band = {"≤2배": 0, "≤5배": 0, "≤10배": 0, "10배 초과": 0, "못 이음": 0}
    for r, _s, p, _h in rows:
        if p == float("inf"):
            band["못 이음"] += 1
        elif r <= 2:
            band["≤2배"] += 1
        elif r <= 5:
            band["≤5배"] += 1
        elif r <= 10:
            band["≤10배"] += 1
        else:
            band["10배 초과"] += 1
    print(f"  앵커 둘레 {len(rows)}개의 배관/직선 비: "
          + " · ".join(f"{k} {v}" for k, v in band.items()))
    mid = rows[len(rows) // 2]
    print(f"  중앙값 {mid[0]:.1f}배 (직선 {mid[1]:.1f} m → 배관 {mid[2]:.1f} m)")
    print("\n  ★가장 심한 10개 (공간으론 코앞인데 배관으론 멀다)")
    for r, s, p, h in rows[-10:][::-1]:
        print(f"      h{h:<5} 직선 {s:>6.1f} m → 배관 {p:>7.1f} m  ({r:>5.1f}배)"
              f"  유하 {far[h] / 1000:.1f} m")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
