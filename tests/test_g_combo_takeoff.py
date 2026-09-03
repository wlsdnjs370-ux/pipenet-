# -*- coding: utf-8 -*-
"""상하향식(combo) 하향 팔은 «가지관을 따라» 뻗는다.

종전에는 언제나 +X 였다. 가지가 어느 쪽으로 놓였든 하향 팔만 동쪽으로
삐져나와, 수리계산 화면에서 그 헤드가 «엉뚱한 데 박힌» 것으로 보였다
(실측: 사람이 찍은 자리에서 300mm 치우침 — 그 도면에서는 헤드 두 개).

하향식(`build_pendant`)은 이미 `_pendant_arm_to_tee` 로 실제 배관을 따라간다.
상하향식만 이 규칙 밖에 있었다.

★수리계산 값은 이 고침에 영향받지 않는다. 세로·팔 배관의 길이는 좌표로
  재지 않고 입력값(①②③④)을 그대로 쓴다(`make_vert_pipe`). 실측으로도
  배관 272 · 노즐 26 · 부속 135 · 표고 273 이 전부 그대로였다.
"""
from __future__ import annotations

import math
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
for _p in (_ROOT, _G):
    if _p not in sys.path:
        sys.path.append(_p)


def _dir(line_xy, other_xy):
    """`_takeoff_dir` 을 한 줄짜리 그래프로 부른다."""
    from services.cad_import.convert.engine import _takeoff_dir
    nodes = {
        "L": {"coords": [line_xy[0], line_xy[1], 0.0]},
        "O": {"coords": [other_xy[0], other_xy[1], 0.0]},
    }
    pipes = {"P1": {"start": "L", "end": "O"}}
    return _takeoff_dir("L", ["P1"], pipes, nodes)


def test_가지가_X축이면_X로_뻗는다():
    ux, uy = _dir((0.0, 0.0), (5.0, 0.0))
    assert (round(ux, 6), round(uy, 6)) == (1.0, 0.0)


def test_가지가_Y축이면_Y로_뻗는다():
    """★종전 코드는 여기서도 +X 로 뻗었다 — 팔이 관을 가로질러 삐져나왔다."""
    ux, uy = _dir((0.0, 0.0), (0.0, 5.0))
    assert (round(ux, 6), round(uy, 6)) == (0.0, 1.0)


def test_반대쪽으로_놓인_가지도_그_쪽을_따른다():
    ux, uy = _dir((10.0, 10.0), (10.0, 4.0))
    assert (round(ux, 6), round(uy, 6)) == (0.0, -1.0)


def test_비스듬한_가지는_그_축의_단위벡터다():
    ux, uy = _dir((0.0, 0.0), (3.0, 4.0))
    assert abs(math.hypot(ux, uy) - 1.0) < 1e-9
    assert abs(ux - 0.6) < 1e-9 and abs(uy - 0.8) < 1e-9


def test_방향을_못_정하면_종전대로_X():
    """★기본값은 반드시 +X 다 — 못 정할 때 그림이 갑자기 달라지면 안 된다."""
    # 수직관(수평 성분 0)
    assert _dir((2.0, 3.0), (2.0, 3.0)) == (1.0, 0.0)
    # 붙은 관이 없음
    from services.cad_import.convert.engine import _takeoff_dir
    nodes = {"L": {"coords": [0.0, 0.0, 0.0]}}
    assert _takeoff_dir("L", [], {}, nodes) == (1.0, 0.0)
    # 좌표를 못 읽음
    assert _takeoff_dir("없음", ["P1"], {}, {}) == (1.0, 0.0)


def test_상대_노드의_좌표가_없으면_다음_관을_본다():
    """한 관이 못 쓰겠다고 바로 +X 로 물러서지 않는다."""
    from services.cad_import.convert.engine import _takeoff_dir
    nodes = {"L": {"coords": [0.0, 0.0, 0.0]},
             "B": {"coords": [0.0, -7.0, 0.0]}}
    pipes = {"P1": {"start": "L", "end": "A"},      # A 는 좌표 없음
             "P2": {"start": "L", "end": "B"}}
    ux, uy = _takeoff_dir("L", ["P1", "P2"], pipes, nodes)
    assert (round(ux, 6), round(uy, 6)) == (0.0, -1.0)
