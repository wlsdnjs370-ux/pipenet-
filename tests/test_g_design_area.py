# -*- coding: utf-8 -*-
"""[BLOCKED §31] 설계면적 = «가장 불리한 K개». 규칙은 두 줄이다.

■ 확정된 규칙 (2026-09-07 · 사용자)

  1순위  사람이 «영역» 을 지정했으면 그 안의 헤드만 후보다.
         (영역이 없으면 도면 전체가 후보 — 곧바로 2순위로 간다.)
  2순위  알람밸브(급수원)에서 **배관을 따라 잰 길이**가 긴 순서로 K 개.
         가장 긴 것부터.

■ 여기까지 온 길

  종전에는 ①앵커를 잡고 ②그 둘레를 «채우는» 방식이었다. 채우는 자를
  배관거리 → 직사각형 → 가지관으로 세 번 바꿔 봤지만 문제는 같았다 —
  어느 자든 «가장 불리한 헤드» 를 빠뜨린다. 실측(B1F·K=30)에서 가장 먼
  30개 중 7개(23%)만 뽑히고 347.9 m 짜리 대신 306.9 m 짜리를 골랐다.

  지금 규칙은 그 자를 없앤다. 불리한 순서 그대로 K 개다.

■ 대가도 시험으로 남긴다

  뽑힌 K 개는 도면 여러 곳에 흩어질 수 있다. 한 구역으로 모으고 싶으면
  사람이 «영역» 을 지정한다 — 그것이 1순위인 이유다. 프로그램이 대신
  «구역» 을 지어내지 않는다.
"""
from __future__ import annotations

import math
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "cad_project_editor_g"),
           os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.design.worst import worst_k_heads    # noqa: E402


def _comb():
    """빗 모양 망 — 주관 하나에 가지관 둘. 실도면의 최소 판이다.

        가지 A (x=100)   가지 B (x=90)
          h y=50            h y=50
          …                 …
        ────────── 주관 y=0 ──────────  ← 급수원 x=0

    급수원에서 배관을 따라 재면 A 가지가 B 보다 10 씩 멀다(주관이 그만큼
    더 길다). 그래서 «먼 순서» 는 A 가지가 통째로 앞선다.
    """
    pts: list = [(0.0, 0.0)]                 # 0 = 급수원
    edges: list = []
    hnodes: list = []
    xy: list = []
    pts.append((90000.0, 0.0))               # 1
    pts.append((100000.0, 0.0))              # 2
    edges += [(0, 1), (1, 2)]
    for bx, base in ((100000.0, 2), (90000.0, 1)):
        prev = base
        for j in range(1, 6):
            pts.append((bx, j * 10000.0))
            n = len(pts) - 1
            edges.append((prev, n))
            hnodes.append({n})
            xy.append((bx, j * 10000.0))
            prev = n
    return pts, edges, hnodes, xy


def _far(pts, edges, hnodes, src=0):
    """급수원에서 배관을 따라 잰 헤드별 유하거리 — 시험이 제 자로 다시 잰다."""
    import heapq
    adj: dict = {}
    for a, b in edges:
        adj.setdefault(a, []).append(b)
        adj.setdefault(b, []).append(a)
    dist = {src: 0.0}
    pq = [(0.0, src)]
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, float("inf")):
            continue
        for v in adj.get(u, ()):
            nd = d + math.dist(pts[u], pts[v])
            if nd < dist.get(v, float("inf")):
                dist[v] = nd
                heapq.heappush(pq, (nd, v))
    return {hi: min(dist[n] for n in ns if n in dist)
            for hi, ns in enumerate(hnodes) if any(n in dist for n in ns)}


def test_기준_헤드는_가장_먼_헤드다():
    pts, edges, hnodes, xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=6, head_xy=xy)
    assert xy[w["worst_head"]] == (100000.0, 50000.0), xy[w["worst_head"]]


def test_먼_순서_그대로_K개를_추린다():
    """★이것이 규칙의 전부다 — 다른 자를 대지 않는다."""
    pts, edges, hnodes, xy = _comb()
    far = _far(pts, edges, hnodes)
    for k in (3, 6, 8):
        w = worst_k_heads(pts, edges, hnodes, [0], k=k, head_xy=xy)
        want = [h for h, _ in sorted(far.items(),
                                     key=lambda kv: (-kv[1], kv[0]))][:k]
        assert set(w["heads"]) == set(want), (
            k, sorted(xy[h] for h in w["heads"]))


def test_가장_먼_것을_절대_빠뜨리지_않는다():
    """종전 방식이 낸 결함이 바로 이것이었다 — 347.9 m 를 두고 306.9 m 를 뽑음."""
    pts, edges, hnodes, xy = _comb()
    far = _far(pts, edges, hnodes)
    order = sorted(far, key=lambda h: (-far[h], h))
    for k in (1, 4, 7, 10):
        w = worst_k_heads(pts, edges, hnodes, [0], k=k, head_xy=xy)
        got = set(w["heads"])
        # 뽑힌 것 중 가장 가까운 것보다 «먼» 헤드가 안 뽑혔으면 안 된다.
        cut = min(far[h] for h in got)
        missed = [h for h in order if far[h] > cut and h not in got]
        assert not missed, [xy[h] for h in missed]


def test_배관을_따라_재지_직선으로_안_잰다():
    """★«변위 아님» 이 규칙의 핵심이다 — 둘이 정반대인 판으로 못 박는다.

        급수원 (0,0)
          ├── 곧게 → 헤드 B (200, 0)      직선 200 · 배관 200
          └── 멀리 돌아 → 헤드 A (10, 60)  직선  61 · 배관 994

    A 는 급수원 «코앞» 에 있지만 물은 500 m 를 돌아온다. 직선으로 재면 B 가
    불리해 보이고, 배관으로 재면 A 가 압도적으로 불리하다. 규칙은 배관이다.
    """
    pts = [(0.0, 0.0), (500000.0, 0.0), (10000.0, 60000.0), (200000.0, 0.0)]
    edges = [(0, 1), (1, 2), (0, 3)]
    hnodes = [{2}, {3}]                      # 헤드 A, 헤드 B
    xy = [(10000.0, 60000.0), (200000.0, 0.0)]
    w = worst_k_heads(pts, edges, hnodes, [0], k=1, head_xy=xy)
    assert w["heads"] == [0], "직선으로 쟀다 — 배관으로 재야 한다"
    assert w["far_m"] > 900.0, w["far_m"]


def test_영역이_1순위다():
    """영역을 주면 그 안에서만 고른다 — 밖의 더 먼 헤드는 안 본다."""
    pts, edges, hnodes, xy = _comb()
    only = {hi for hi, p in enumerate(xy) if p[0] == 90000.0}   # B 가지만
    w = worst_k_heads(pts, edges, hnodes, [0], k=3, head_xy=xy,
                      only_heads=only)
    assert {xy[h][0] for h in w["heads"]} == {90000.0}
    assert xy[w["worst_head"]] == (90000.0, 50000.0)


def test_같은_입력에_같은_산출():
    """유하거리가 똑같은 헤드가 여럿일 때 순서가 흔들리면 안 된다."""
    pts, edges, hnodes, xy = _comb()
    a = worst_k_heads(pts, edges, hnodes, [0], k=4, head_xy=xy)
    b = worst_k_heads(pts, edges, hnodes, [0], k=4, head_xy=xy)
    assert a["heads"] == b["heads"]


def test_흩어진_정도를_숨기지_않고_낸다():
    """★대가를 수치로 남긴다 — 흩어지면 사람이 «영역» 을 쓸지 판단해야 한다."""
    pts, edges, hnodes, xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=8, head_xy=xy)
    assert w["area_w_m"] > 0 and w["area_h_m"] > 0
    assert abs(w["area_m2"] - w["area_w_m"] * w["area_h_m"]) < 0.11
    assert w["span_m"] == max(w["area_w_m"], w["area_h_m"])


def test_헤드_좌표를_안_주면_부착_노드로_잰다():
    """옛 호출부(데스크톱 G 등)가 그대로 돌아야 한다 — 터지지 않는다."""
    pts, edges, hnodes, _xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=6)
    assert len(w["heads"]) == 6
    assert w["area_m2"] > 0
