# -*- coding: utf-8 -*-
"""최불리 선정 — 영역 지정(모듈 A 의 zones)과 최원 유하거리 «경로».

두 가지를 못박는다:

  ① 영역으로 후보를 가두면 앵커도 그 안에서 나온다.
     도면 장 나누기는 자동으로 잰 경계라 실무에서 늘 맞지는 않는다 — 한 층에
     방화구획이 여럿이거나 주차장·기계실이 섞이면 앵커가 그리로 튄다.

  ② far_m 이 «어느 줄» 인지가 나온다.
     앵커까지의 거리는 전부터 냈지만 그 거리가 어느 관을 타고 오는지는
     화면에서 읽을 수 없었다. worst_path 가 그 줄이다.
"""
from __future__ import annotations

import math
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.design.worst import worst_k_heads  # noqa: E402


def _line(n: int, dx: float = 1000.0):
    """급수원 0 에서 오른쪽으로 뻗은 직선 — 노드 i 마다 헤드 하나."""
    pts = [(i * dx, 0.0) for i in range(n)]
    edges = [(i, i + 1) for i in range(n - 1)]
    hnodes = [[i] for i in range(1, n)]      # 헤드 i-1 → 노드 i
    return pts, edges, hnodes, [0]


def test_앵커는_가장_먼_헤드다():
    pts, edges, hnodes, src = _line(6)
    w = worst_k_heads(pts, edges, hnodes, src, k=3)
    assert w["worst_head"] == len(hnodes) - 1          # 맨 끝 헤드
    assert w["far_m"] == 5.0                       # 5,000mm


def test_최원_경로가_급수원에서_앵커까지다():
    pts, edges, hnodes, src = _line(6)
    w = worst_k_heads(pts, edges, hnodes, src, k=3)
    assert w["worst_path"] == [0, 1, 2, 3, 4, 5]
    assert w["worst_path_m"] == 5.0


def test_최원_경로_길이가_far_m와_같다():
    """같은 거리를 두 곳에서 재므로 어긋나면 둘 중 하나가 틀린 것이다."""
    pts, edges, hnodes, src = _line(9)
    w = worst_k_heads(pts, edges, hnodes, src, k=4)
    assert math.isclose(w["worst_path_m"], w["far_m"], abs_tol=0.01)


def test_경로는_급수원에서_시작한다():
    pts, edges, hnodes, src = _line(6)
    w = worst_k_heads(pts, edges, hnodes, src, k=3)
    assert w["worst_path"][0] == 0


def test_헤드가_없으면_경로도_빈다():
    pts, edges, hnodes, src = _line(4)
    w = worst_k_heads(pts, edges, hnodes, src, k=3, only_heads=set())
    assert w["heads"] == [] and w["worst_path"] == []
    assert w["worst_path_m"] == 0.0


# ─────────────────────────────────────────────── 영역(only_heads)
def test_영역으로_가두면_앵커가_그_안에서_나온다():
    """가두지 않으면 맨 끝이 앵커지만, 가두면 그 안의 최원이 앵커다."""
    pts, edges, hnodes, src = _line(8)
    free = worst_k_heads(pts, edges, hnodes, src, k=2)
    assert free["worst_head"] == 6                       # 맨 끝 헤드(노드 7)

    # 헤드 0~2(노드 1~3)만 후보로 — 「영역」이 하는 일과 같다.
    caged = worst_k_heads(pts, edges, hnodes, src, k=2, only_heads={0, 1, 2})
    assert caged["worst_head"] == 2
    assert caged["far_m"] == 3.0
    assert set(caged["heads"]) <= {0, 1, 2}


def test_영역_밖_헤드는_설계면적에_안_들어온다():
    pts, edges, hnodes, src = _line(10)
    w = worst_k_heads(pts, edges, hnodes, src, k=5, only_heads={0, 1, 2, 3})
    assert set(w["heads"]) <= {0, 1, 2, 3}
    assert w["reachable"] == 4, "후보 수도 가둔 만큼만 센다"


def test_영역_안_경로도_급수원까지_이어진다():
    """가둔 것은 «후보» 지 «망» 이 아니다 — 경로는 급수원까지 그대로 간다."""
    pts, edges, hnodes, src = _line(8)
    w = worst_k_heads(pts, edges, hnodes, src, k=2, only_heads={4, 5})
    assert w["worst_head"] == 5
    assert w["worst_path"][0] == 0            # 영역 밖을 지나서라도 급수원까지
    assert len(w["worst_path"]) == 7          # 노드 0..6


def test_영역이_한_헤드만_잡아도_돈다():
    pts, edges, hnodes, src = _line(6)
    w = worst_k_heads(pts, edges, hnodes, src, k=30, only_heads={1})
    assert w["heads"] == [1]
    assert w["worst_head"] == 1
    assert w["worst_path_m"] == 2.0
