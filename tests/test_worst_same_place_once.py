# -*- coding: utf-8 -*-
"""[최불리] «같은 자리» 의 헤드는 하나로 센다 — 기준개수 K 가 진짜 K 가 되도록.

■ 무엇이 문제였나 (2026-09-09 사용자 지적)

  「기준개수 30개로 헤드 배관망을 추출해도 30개 밑으로 헤드가 검출된다」

  도면이 헤드 기호를 겹쳐 그리면(대명동 실측: 노랑 r42 위에 색12 r42 가 그대로
  포개져 있다) 헤드는 둘인데 **자리는 하나**다. 선정은 그것을 둘로 셌고
  (`design/worst.py`), 표는 하나로 만들었다(`convert/planar.py` 의
  `head_vid[vid]` 가 같은 절점을 덮는다). 그래서::

      선정이 고른 헤드   30
      서로 다른 «자리»   26        ← 겹친 자리 4곳
      표에 온 노즐       26        ← 26 = 26, 겹침이 원인

  잃어버린 헤드는 없었다. **세는 자가 두 벌**이었던 것이다.

■ 조치 — 자리로 묶어 세고, 그만큼 다음 순위를 채운다

  없는 헤드를 지어내는 것이 아니라 **이미 하나인 것을 하나로 센다.**
  자는 «헤드 중심 좌표»다 — 표가 합치는 자와 같은 것을 봐야 두 수가 안 갈린다.

  조치 뒤 실측(대명동 K=30): 고른 헤드 30 · 서로 다른 자리 30 · **노즐 30**.
  골든(K=10): far_m 52.69·span_m 11.1·max_load 10 **불변**, corridor 총연장
  73.16 → 74.02 m, 최불리 kfp 노드 59 → **67** (접혀 사라지던 헤드가 올라왔다).
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _wk():
    from services.cad_import.design.worst import worst_k_heads
    return worst_k_heads


# 급수원 0 ─ 1 ─ 2 ─ 3 ─ 4  (100 mm 간격의 한 줄)
_PTS = [(0.0, 0.0), (100.0, 0.0), (200.0, 0.0), (300.0, 0.0), (400.0, 0.0)]
_EDGES = [(0, 1), (1, 2), (2, 3), (3, 4)]


def _case(dup=True):
    """헤드 넷 — 그중 둘이 **같은 자리**(겹쳐 그린 기호)."""
    hnodes = [[4], [4], [3], [2]]
    xy = [(400.0, 0.0, 42.0),
          (400.0, 0.0, 42.0) if dup else (350.0, 0.0, 42.0),
          (300.0, 0.0, 42.0),
          (200.0, 0.0, 42.0)]
    return hnodes, xy


def test_같은_자리는_한_개로_센다():
    """★K=2 를 넣으면 «서로 다른 자리» 2곳이 나와야 한다."""
    hnodes, xy = _case(dup=True)
    w = _wk()(_PTS, _EDGES, hnodes, [0], k=2, head_xy=xy)
    got = {(round(xy[hi][0], 1), round(xy[hi][1], 1)) for hi in w["heads"]}
    assert len(w["heads"]) == 2, w["heads"]
    assert len(got) == 2, got            # ← 종전에는 둘 다 (400,0) 이었다
    assert w["merged"] == 1


def test_겹침이_없으면_한_바이트도_안_바뀐다():
    """★골든이 지키는 것 — 겹치지 않은 도면의 산출은 종전 그대로다."""
    hnodes, xy = _case(dup=False)
    w = _wk()(_PTS, _EDGES, hnodes, [0], k=2, head_xy=xy)
    assert w["heads"] == [0, 1]
    assert w["merged"] == 0 and w["merged_xy"] == []


def test_닿는_헤드도_자리로_센다():
    """「기준개수 K인데 닿는 헤드가 N개뿐」이라는 말이 정직해야 한다.

    헤드 수로 세면 4라 K=4 가 통과하는데, 자리는 3곳뿐이라 표는 3개가 된다 —
    통과시켜 놓고 뒤에서 모자라는 것이 가장 나쁘다.
    """
    hnodes, xy = _case(dup=True)
    w = _wk()(_PTS, _EDGES, hnodes, [0], k=4, head_xy=xy)
    assert w["reachable"] == 3
    assert len(w["heads"]) == 3


def test_최원_유하거리는_안_바뀐다():
    """겹침을 접어도 «1등» 은 그대로다 — 대표를 같은 자리에서 고르기 때문."""
    hnodes, xy = _case(dup=True)
    a = _wk()(_PTS, _EDGES, hnodes, [0], k=2, head_xy=xy)
    assert a["far_m"] == round(400.0 / 1000.0, 2)
    assert xy[a["worst_head"]][:2] == (400.0, 0.0)


def test_겹친_자리를_좌표로_말한다():
    """조용히 접지 않는다 — 어디를 접었는지 사람이 볼 수 있어야 한다."""
    hnodes, xy = _case(dup=True)
    w = _wk()(_PTS, _EDGES, hnodes, [0], k=2, head_xy=xy)
    assert w["merged_xy"] == [[400.0, 0.0, 2]], w["merged_xy"]


def test_중심을_모르면_부착_절점으로_센다():
    """`head_xy` 가 없어도 «자리» 개념은 살아 있어야 한다(하위호환 경로)."""
    hnodes, _xy = _case(dup=True)
    w = _wk()(_PTS, _EDGES, hnodes, [0], k=2)
    assert len(w["heads"]) == 2 and w["merged"] == 1


def test_대표는_번호로_못_박는다():
    """같은 입력에 같은 산출 — 무리 안에서 «가장 작은 번호» 를 대표로."""
    hnodes, xy = _case(dup=True)
    for _ in range(3):
        w = _wk()(_PTS, _EDGES, hnodes, [0], k=2, head_xy=xy)
        assert w["heads"] == [0, 2], w["heads"]


def test_왜_그렇게_세는지_적어_두었다():
    """숫자만 고치고 이유를 안 적으면 다음 사람이 되돌린다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "design" / "worst.py").read_text(encoding="utf-8")
    i = src.index("def _place(")
    seg = src[max(0, i - 1800):i]
    assert "표는 하나로 만들었다" in seg or "표는 하나로 만든다" in seg
    assert "노즐이 **26개**만" in seg, "실측이 안 적혀 있다"
    assert "바이트도 안 바뀐다" in seg, "무변경 조건이 안 적혀 있다"
