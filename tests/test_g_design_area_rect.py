# -*- coding: utf-8 -*-
"""[BLOCKED §31] 설계면적은 «면적» 이다 — 직사각형으로 고른다.

■ 무엇이 틀려 있었나

  ②단계가 앵커에서 **배관 거리**로 가까운 K개를 골랐다. 배관 거리는 공간
  거리와 딴판이다 — 옆 가지관의 헤드는 공간으로 코앞인데 주관을 돌아가느라
  «멀다» 고 읽힌다(B1F 실측: 직선 3.0 m 인데 배관 100.4 m · 34배).

  그래서 설계면적이 **두 조각으로 갈라졌다**: 앵커 줄에서 8개를 집고 4.3 m 를
  건너뛰어 24~36 m 떨어진 곳의 22개를 집었다. 어느 규범을 따르든 설계면적은
  하나의 연속한 구역이라야 한다.

■ 지금 규칙

  앵커를 품는 **직사각형**(체비셰프 상자)을 K개가 담길 때까지 넓힌다.
  사람이 화면에서 «영역 지정» 으로 그리는 그 사각형과 같은 개념이다.

  ★한계는 코드 주석에 적어 뒀다 — 이 상자는 방호구역을 모른다. 그 판단은
    사람이 영역 지정으로 한다.
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
          h y=50 ← 앵커     h y=50
          h y=40            h y=40
          …                 …
        ────────── 주관 y=0 ──────────  ← 급수원 x=0
    """
    pts: list = [(0.0, 0.0)]                 # 0 = 급수원
    edges: list = []
    hnodes: list = []
    xy: list = []
    # 주관 0 → (90,0) → (100,0)
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


def test_앵커는_가장_먼_헤드다():
    pts, edges, hnodes, xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=6, head_xy=xy)
    # 가지 A 의 맨 끝(100,50) 이 급수원에서 가장 멀다.
    assert xy[w["worst_head"]] == (100000.0, 50000.0), xy[w["worst_head"]]


def test_옆_가지관을_건너뛰지_않는다():
    """★이것이 사용자가 본 결함이다.

    앵커가 달린 가지관에 헤드가 5개뿐인데 K=6 이면 한 개를 옆에서 가져와야
    한다. 배관 거리로 고르면 옆 가지관의 **맨 아래**(공간으론 가장 먼) 헤드를
    집는다 — 주관을 돌아가는 길이가 그쪽이 짧기 때문이다. 공간으로 고르면
    바로 옆(90,50)을 집는다.
    """
    pts, edges, hnodes, xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=6, head_xy=xy)
    got = {xy[h] for h in w["heads"]}
    assert (90000.0, 50000.0) in got, sorted(got)
    assert (90000.0, 10000.0) not in got, "옆 가지관 맨 아래를 집었다"


def test_설계면적이_한_덩어리다():
    """두 조각으로 갈라지면 설계면적이 아니다 — 헤드 간 최대 «빈틈» 을 본다."""
    pts, edges, hnodes, xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=6, head_xy=xy)
    P = sorted(xy[h] for h in w["heads"])
    # 각 헤드에서 가장 가까운 다른 헤드까지 — 이것이 크면 흩어진 것이다.
    worst_gap = max(min(math.dist(p, q) for q in P if q != p) for p in P)
    assert worst_gap <= 12000.0, f"가장 가까운 이웃까지 {worst_gap / 1000:.1f} m"


def test_직사각형_크기를_산출한다():
    """설계면적은 규정이 ㎡ 로 말하는 값이다 — 그 수를 직접 낸다."""
    pts, edges, hnodes, xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=6, head_xy=xy)
    assert w["area_w_m"] > 0 and w["area_h_m"] > 0
    assert abs(w["area_m2"] - w["area_w_m"] * w["area_h_m"]) < 0.11
    assert w["span_m"] == max(w["area_w_m"], w["area_h_m"])


def test_같은_입력에_같은_산출():
    """상자 경계에서 값이 같으면 순서가 흔들린다 — 못 박아 뒀다."""
    pts, edges, hnodes, xy = _comb()
    a = worst_k_heads(pts, edges, hnodes, [0], k=4, head_xy=xy)
    b = worst_k_heads(pts, edges, hnodes, [0], k=4, head_xy=xy)
    assert a["heads"] == b["heads"]


def test_헤드_좌표를_안_주면_부착_노드로_잰다():
    """옛 호출부(데스크톱 G 등)가 그대로 돌아야 한다 — 터지지 않는다."""
    pts, edges, hnodes, _xy = _comb()
    w = worst_k_heads(pts, edges, hnodes, [0], k=6)
    assert len(w["heads"]) == 6
    assert w["area_m2"] > 0
