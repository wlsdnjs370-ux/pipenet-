# -*- coding: utf-8 -*-
"""[위상] 아이소가 평면과 «달라 보이던» 진짜 원인 — 담당 헤드 수의 소실.

■ 사용자 지적

  「평면에서 보기에서 본 배관망이랑 30° 아이소매트릭에서 본 배관망이랑 위상이
  조금 다른데? 형태를 옮기는 과정에서 뭔가가 훼손된 것 같다.」

■ 재어 보니 위상은 같았다 (`scripts/_probe_design_topology.py` · 대명동 K30)

    평면 corridor   갈림 21 · 끝 23 · 사슬 43 · 고리 0
    설계 표         갈림 21 · 끝 23 · 사슬 43 · 고리 0

  설계가 가리킨 도면 간선 중 corridor «집합» 밖에 있는 49개도 실제로는
  corridor 선 위에 **거리 0mm** 로 얹혀 있었다 — 다른 길이 아니라, 평면
  그래프를 세우며 절점을 다시 매겨 (i,j) 키만 어긋난 것이다.

■ 훼손은 «굵기» 에 있었다

  간선 굵기는 담당 헤드 수다. 그 수를 board 간선 키로 찾는데, 위의 키 어긋남
  때문에 **216개 중 49개(22%)** 가 빗나갔다. 그런데 못 찾은 자리에 `0` 을 박아
  넣고 있어서, 뒤따르는 `tree_loads` 폴백이 `setdefault` 라 영영 못 들어왔다.

  결과: 주배관이 가지 끝처럼 가늘게 그려졌다 — 사람 눈에는 «위상이 다른» 망이다.
  고친 뒤 굵기 0 인 배관은 51 → 2 로 떨어졌다.

■ 여기서 못박는 것

  ⑴ 못 찾으면 0 을 적지 않는다 — 다른 곳이 아는 값이 들어올 자리를 막지 않는다.
  ⑵ 두 곳이 다 알면 board 쪽(worst.loads)이 이긴다.
  ⑶ 둘 다 모르면 그 배관은 아예 빠진다(0 으로 «아는 척» 하지 않는다).
"""
from __future__ import annotations

import os

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _fn():
    from routes.module_f.api_design import _load_map
    return _load_map


def test_board_간선에서_찾으면_그_값을_쓴다():
    got = {"worst": {"loads": {(1, 2): 7}}, "edge_ref": {"P1": [2, 1]},
           "tree_loads": {"P1": 99}}
    assert _fn()(got) == {"P1": 7}


def test_못_찾으면_0_을_박지_않는다():
    """★이것이 실제 결함이었다 — 0 을 박아 폴백을 막았다."""
    got = {"worst": {"loads": {(5, 6): 3}}, "edge_ref": {"P1": [1, 2]},
           "tree_loads": {"P1": 12}}
    assert _fn()(got) == {"P1": 12}, "tree_loads 가 못 들어온다"


def test_역참조가_없는_배관은_tree_loads_가_채운다():
    """헤드 접속관·가지 상승은 board 간선이 없다 — 그 자리는 tree 가 안다."""
    got = {"worst": {"loads": {}}, "edge_ref": {},
           "tree_loads": {"P9": 4}}
    assert _fn()(got) == {"P9": 4}


def test_둘_다_모르면_아는_척하지_않는다():
    got = {"worst": {"loads": {}}, "edge_ref": {"P1": [1, 2]},
           "tree_loads": {}}
    assert _fn()(got) == {}


def test_망가진_역참조는_건너뛴다():
    got = {"worst": {"loads": {(1, 2): 5}},
           "edge_ref": {"P1": None, "P2": ["a", "b"], "P3": [1, 2]},
           "tree_loads": {}}
    assert _fn()(got) == {"P3": 5}


def test_0_은_모름이_아니라_값이다():
    """진짜로 0 인 배관(급수원 쪽 등)은 0 으로 남아야 한다."""
    got = {"worst": {"loads": {(1, 2): 0}}, "edge_ref": {"P1": [1, 2]},
           "tree_loads": {"P1": 8}}
    assert _fn()(got) == {"P1": 0}, "board 가 0 이라고 아는 것을 덮었다"


def test_미리보기가_이_함수를_쓴다():
    """두 곳이 각자 셈하면 화면과 시험이 다른 값을 본다."""
    py = open(os.path.join(_ROOT, "routes", "module_f", "api_design.py"),
              encoding="utf-8").read()
    assert "load_of = _load_map(got)" in py
    assert "loads.get((min(i, j), max(i, j)), 0)" not in py, \
        "0 을 박는 옛 셈이 남아 있다"


def test_평면과_아이소를_같은_망으로_볼_수_있다():
    """«설계망만» 을 고르면 평면이 아이소와 같은 망만 보인다.

    값은 손질의 select 하나뿐이고 수리계산 쪽은 그 얼굴이다 — 두 곳이 각자
    값을 들면 화면끼리 다른 말을 한다.
    """
    html = open(os.path.join(_ROOT, "templates", "module_f.html"),
                encoding="utf-8").read()
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    assert 'id="dg-plan-view"' in html
    i = js.index('$("dg-plan-view").onchange')
    seg = js[i:i + 200]
    assert '$("ed-worst-view").value = $("dg-plan-view").value' in seg, seg
    j = js.index("function renderPlanUnderlay()")
    assert '$("dg-plan-view").value = $("ed-worst-view").value' in js[j:j + 600]
