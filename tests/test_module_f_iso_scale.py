# -*- coding: utf-8 -*-
"""[아이소] 표고를 **평면과 같은 자**로 그린다 · 교차는 끊어 그린다.

■ 사용자 지적

  「여전히 아이소 변환 시 위상이 깨지는데, 좀 더 면밀하게 검토를 부탁해.」

■ 세 번 재고서야 자리를 찾았다 (`scripts/_probe_iso_*.py`)

  ① 자료의 위상은 같다        갈림 21 · 끝 23 · 사슬 43 · 고리 0 (양쪽)
  ② 배치도 같다              board→설계 잔차 중앙 0.047% · 최대 0.09%
  ③ 그런데 아이소 그림에는 없던 교차가 **15건** 생긴다.

  범인은 lift 였다. 옛 식은 `대각선 · 0.5 · 배율 / 표고폭` — 표고 폭이 얼마든
  **화면 절반**으로 늘린다. 단위세대(표고 폭 3.3 m)에서는 1 m 가 508 단위가
  되는데 같은 화면의 평면 1 m 는 83 단위다. 세로만 **6배** 부푼 그림이라
  0.3 m 짜리 헤드 접속관이 옆 가지를 가로질렀다.

    lift 508 → 교차 18 · lift 83(평면과 같은 자) → 6 · lift 0 → 0

■ 고침 둘

  ⑴ 표고를 평면과 같은 자로 그린다. 더 펼쳐 보려면 «고도 펼침 배율» 을 쓴다
     — 자를 기본으로 부풀려 두면 얼마나 부풀었는지 아무도 모른다.
     헤드 스텁도 «표가 말하는 그 길이» 로 세운다(표고 차를 알 때).
  ⑵ 남는 교차는 대부분 **진짜 3차원 교차**다(7건 중 6건 — 교차점에서 두 배관의
     표고가 다르다). 제도 규약대로 **아래로 지나가는 쪽을 끊어** 그린다.
     표고가 같으면 끊지 않는다 — 그것은 실제로 거기서 만나는 것일 수 있다.

■ ★산출은 안 바뀐다

  자리는 241/241 절점이 바뀌지만 표고 0건 · 배관행 동일 · 노즐행 동일.
  바뀌는 것은 그림뿐이다.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


class _T:
    """PipeTables 흉내 — bake_isometric 이 보는 것은 nodes 뿐이다."""

    def __init__(self, nodes):
        self.nodes = nodes


def _bake():
    import sys
    g = os.path.join(_ROOT, "cad_project_editor_g")
    if g not in sys.path:
        sys.path.insert(0, g)
    from services.cad_import.design.sdf_post import bake_isometric
    return bake_isometric


def _nodes():
    return [
        {"label": "1", "x": 0.0, "y": 0.0, "elevation": 0.0},
        {"label": "2", "x": 100.0, "y": 0.0, "elevation": 0.0},
        {"label": "3", "x": 100.0, "y": 0.0, "elevation": -0.3},   # 헤드
    ]


def test_평면과_같은_자로_표고를_그린다():
    """★lift 는 «1 m 가 화면에서 몇 단위인가» 다 — 화면 절반이 아니다."""
    t = _T(_nodes())
    got = _bake()(t, units_per_m=83.0)
    assert got["lift"] == pytest.approx(83.0)


def test_고도_펼침_배율이_그_자를_곱한다():
    t = _T(_nodes())
    got = _bake()(t, units_per_m=83.0, iso_z_scale=3.0)
    assert got["lift"] == pytest.approx(249.0)


def test_자를_안_주면_옛_식_그대로():
    """다른 부르는 곳(모듈 A 계통도 등)의 동작을 바꾸지 않는다."""
    t = _T(_nodes())
    got = _bake()(t)
    assert got["lift"] != pytest.approx(83.0)
    assert got["lift"] > 0


def test_헤드는_표가_말하는_길이로_선다():
    """표는 0.3 m 라는데 그림이 그보다 긴 토막을 그리면 그림이 표를 배신한다."""
    t = _T(_nodes())
    got = _bake()(t, units_per_m=100.0,
                  head_nodes=["3"], head_parent={"3": "2"})
    assert got["vertical"] == 1
    n2 = next(n for n in t.nodes if n["label"] == "2")
    n3 = next(n for n in t.nodes if n["label"] == "3")
    assert n3["x"] == pytest.approx(n2["x"])          # 수직으로 선다
    assert n3["y"] == pytest.approx(n2["y"] - 30.0)   # 0.3 m × 100 · 아래로


def test_표고_차를_모르면_옛_고정_길이로_선다():
    """정보가 없을 때까지 «진짜 길이» 인 척하지 않는다."""
    nodes = _nodes()
    nodes[2]["elevation"] = 0.0
    t = _T(nodes)
    got = _bake()(t, units_per_m=100.0, head_stub_ratio=0.1,
                  head_nodes=["3"], head_parent={"3": "2"})
    n2 = next(n for n in t.nodes if n["label"] == "2")
    n3 = next(n for n in t.nodes if n["label"] == "3")
    assert n3["y"] == pytest.approx(n2["y"] + got["stub"])


def test_표시_변환이_그_자를_넘긴다():
    """`display_tables` 가 안 넘기면 위 규칙이 한 번도 안 걸린다."""
    src = open(os.path.join(_ROOT, "cad_project_editor_g", "services",
                            "cad_import", "design", "emit.py"),
               encoding="utf-8").read()
    assert "units_per_m=scale * 1000.0" in src, src[:200]


def test_표고는_미리보기로_함께_간다():
    """화면이 «어느 배관이 위로 지나가는지» 를 알아야 끊어 그릴 수 있다."""
    py = open(os.path.join(_ROOT, "routes", "module_f", "api_design.py"),
              encoding="utf-8").read()
    assert '"e": round(float(n.get("elevation", 0) or 0), 3)' in py


# ─────────────────────────────── 화면 — 교차 끊기
_STUB = """
const S = {design: {}};
"""


def _run(view, body):
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("  function crossGaps(v)")
    src = js[i:js.index("\n  }\n", i) + 4]
    prog = "\n".join([f"const VIEW = {json.dumps(view)};", _STUB, src, body])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-800:]
    return json.loads(out.stdout)


def _x_view(e_lo, e_hi):
    """가운데서 교차하는 두 배관 — 표고만 인자로 바꾼다."""
    return {"nodes": [{"label": "A", "x": 0, "y": 0, "e": e_lo},
                      {"label": "B", "x": 10, "y": 10, "e": e_lo},
                      {"label": "C", "x": 0, "y": 10, "e": e_hi},
                      {"label": "D", "x": 10, "y": 0, "e": e_hi}],
            "pipes": [{"label": "P1", "a": "A", "b": "B"},
                      {"label": "P2", "a": "C", "b": "D"}]}


def test_높이가_다르면_아래쪽을_끊는다():
    got = _run(_x_view(0.0, 1.0),
               "const g = crossGaps(VIEW);"
               "console.log(JSON.stringify([...g.entries()]));")
    assert got == [[0, [0.5]]], got      # 낮은 P1(index 0) 만 끊는다


def test_높이가_같으면_끊지_않는다():
    """같은 높이에서 겹치는 것은 실제로 거기서 만나는 것일 수 있다."""
    got = _run(_x_view(0.0, 0.0),
               "const g = crossGaps(VIEW);"
               "console.log(JSON.stringify([...g.entries()]));")
    assert got == []


def test_이웃한_배관은_교차로_보지_않는다():
    v = {"nodes": [{"label": "A", "x": 0, "y": 0, "e": 0},
                   {"label": "B", "x": 10, "y": 0, "e": 0},
                   {"label": "C", "x": 20, "y": 0, "e": 1}],
         "pipes": [{"label": "P1", "a": "A", "b": "B"},
                   {"label": "P2", "a": "B", "b": "C"}]}
    got = _run(v, "const g = crossGaps(VIEW);"
                  "console.log(JSON.stringify([...g.entries()]));")
    assert got == []


def test_그리기가_그_자리를_비운다():
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("  function drawDesign()")
    src = js[i:js.index("\n  }\n", i) + 4]
    assert "strokeWithGaps(" in src, "끊어 그리지 않는다"
    assert "S.design.gaps" in js, "끊을 자리를 안 셈한다"
