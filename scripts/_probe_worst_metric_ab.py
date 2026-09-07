# -*- coding: utf-8 -*-
"""[감사 A/B] 설계면적을 «배관거리» 로 고르는 것이 맞나 — 자를 바꿔 재 본다.

사용자 지적: 「가장 먼 헤드는 인식 못 하고 덜 먼 헤드를 먼저 엮는다」.

지금 알고리즘 ②단계는 앵커에서 **배관거리(pipe distance)** 로 가까운 K개를
고른다. 그런데 옆 가지관의 헤드는 공간으로는 3m 옆인데 배관으로는 «가지 끝까지
내려갔다 주관 타고 옆 가지로 올라오는» 40m 일 수 있다. 그러면 옆 가지(더 불리)
대신 **같은 가지를 따라 급수원 쪽으로 되돌아가며**(덜 불리) 뽑게 된다.

설계면적은 «면적» 이다. 그래서 세 가지 자를 같은 도면에 대 본다:

  A) 지금 것       — 앵커에서 «배관거리» 로 가까운 K개
  B) 공간거리      — 앵커에서 «직선거리» 로 가까운 K개
  C) 먼 순서       — 급수원에서 먼 K개 (설계면적 개념은 없지만 비교용)

각각에 대해 ①유하거리 분포(불리한가) ②공간 폭(한 구역인가)을 잰다.
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


def _report(tag, picked, far, xy, k):
    d = sorted((far[h] for h in picked), reverse=True)
    P = [xy[h] for h in picked]
    span_xy = max(math.dist(a, b) for a in P for b in P) if len(P) > 1 else 0.0
    xs = [p[0] for p in P]
    ys = [p[1] for p in P]
    print(f"  {tag:<26} 유하 {d[0] / 1000:>7.2f}~{d[-1] / 1000:>7.2f} m"
          f" (평균 {sum(d) / len(d) / 1000:>7.2f})"
          f" · 공간폭 {span_xy / 1000:>6.2f} m"
          f" · bbox {(max(xs) - min(xs)) / 1000:.1f}×{(max(ys) - min(ys)) / 1000:.1f} m")


def main() -> int:
    key = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")
    k = int(os.environ.get("MF_K", "30"))
    es = EditSession.open(key, out_dir=None, load_saved=True, use_cache=True)
    b = es.board
    if not b.sources:
        print(f"«{key}» 에 급수원이 없다")
        return 0

    src = _dij(b.pts, b.edges, [list(b.sources)[0]])
    far: dict = {}
    node: dict = {}
    for hi, nodes in enumerate(b.hnodes):
        reach = [n for n in nodes if n in src]
        if reach:
            n0 = min(reach, key=lambda n: src[n])
            node[hi] = n0
            far[hi] = src[n0]
    xy = {hi: (float(b.disks[hi][0]), float(b.disks[hi][1]))
          for hi in far if hi < len(b.disks)}
    far = {hi: v for hi, v in far.items() if hi in xy}

    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k, source_index=0)
    anchor = w["worst_head"]
    print(f"[{key}]  도달 헤드 {len(far)} · K={k}"
          f" · 앵커 h{anchor} ({far[anchor] / 1000:.2f} m)")
    print(f"  ★도달 못 한 헤드 {len(b.hnodes) - len(far)}개"
          f" / 전체 {len(b.hnodes)}개"
          f" ({(len(b.hnodes) - len(far)) / max(len(b.hnodes), 1) * 100:.0f}%)"
          " — 알고리즘이 아예 못 보는 헤드다\n")

    a_xy = xy[anchor]
    an = _dij(b.pts, b.edges, [node[anchor]])

    A = [h for h in w["heads"] if h in far]
    B = sorted(far, key=lambda h: math.dist(xy[h], a_xy))[:k]
    C = sorted(far, key=far.get, reverse=True)[:k]

    _report("A) 지금 — 배관거리", A, far, xy, k)
    _report("B) 공간거리(직선)", B, far, xy, k)
    _report("C) 먼 순서(비교용)", C, far, xy, k)

    print(f"\n  A∩B {len(set(A) & set(B))}개 · A∩C {len(set(A) & set(C))}개"
          f" · B∩C {len(set(B) & set(C))}개")

    # A 가 «되돌아간» 정도 — 앵커보다 급수원에 가까운 헤드를 얼마나 물었나
    back = [(far[anchor] - far[h]) / 1000 for h in A]
    back.sort(reverse=True)
    print(f"  A 가 앵커보다 급수원 쪽으로 되돌아간 거리: 최대 {back[0]:.2f} m"
          f" · 중앙 {back[len(back) // 2]:.2f} m")
    bb = [(far[anchor] - far[h]) / 1000 for h in B]
    bb.sort(reverse=True)
    print(f"  B 는                                    최대 {bb[0]:.2f} m"
          f" · 중앙 {bb[len(bb) // 2]:.2f} m")

    # 앵커에서 배관거리 vs 직선거리 — 같은 헤드에 두 자가 얼마나 다른가
    pair = [(an.get(node[h], float("inf")) / 1000,
             math.dist(xy[h], a_xy) / 1000) for h in B]
    pair = [(p, q) for p, q in pair if p != float("inf")]
    if pair:
        worst = max(pair, key=lambda t: t[0] - t[1])
        print(f"\n  ★같은 헤드인데 배관거리 {worst[0]:.1f} m vs 직선 {worst[1]:.1f} m"
              f" — 차이 {worst[0] - worst[1]:.1f} m")
        print("     (옆 가지관은 공간으로 코앞인데 배관으로는 주관을 돌아간다)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
