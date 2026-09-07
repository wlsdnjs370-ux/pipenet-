# -*- coding: utf-8 -*-
"""[§27-1] 이름 사전이 «정반대로» 읽는 자리를 기하가 바로잡는다.

■ 무엇이 틀렸나

  대명동 계통도의 `SP` 는 이름 사전이 「스프링클러 배관」으로 읽는다. 실제
  내용은 **헤드 기호 918개**(전부 닫힌 폴리라인)이고, 진짜 입상관은 `LSP`·
  `HSP` 에 있다. 그래서 「배관 레이어 고르기」가 사람에게 **틀린 추천**을 했다.

■ 왜 이 지문인가 — 두 번 기각한 뒤에 남은 것

  §27 은 선분 길이와 접점 수를 실제로 재서 둘 다 기각했다(평면도의 진짜
  배관은 짧은 선분이 대부분이라 「길면 배관」이 정확히 거꾸로 작동한다).
  그때 같이 적어 둔 문장이 길을 남겼다 — **「entity 단위(닫힘·도형 지름)로는
  갈리지만 그 정보는 묶음에 없다」**. `_categorize_layer` 는 이름만 받지만,
  부르는 쪽은 entity 를 갖고 있다.

  지문은 그래프가 이미 쓰는 사실 그대로다: 첫점≈끝점인 도형은 배관이 아니다.
  **선형 도형이 하나도 없는** 레이어로는 배관망을 세울 수 없다.

■ 이 시험이 지키는 셋

  ⑴ 교정 4장에서 **진짜 배관 레이어는 전부 닫힘 0%** 라 안 건드린다.
  ⑵ **그래프가 안 바뀐다** — 닫힌 PL 은 이미 컷에서 잘리고 있었다. 바뀌는
     것은 사람에게 하는 «말» 뿐이다.
  ⑶ 「헤드다」라고 말하지 않는다 — 아는 것은 «닫힌 기호» 까지다. 없는 확신을
     만들지 않고 틀린 확신 하나를 거둘 뿐이다.
"""
from __future__ import annotations

import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from remote30_prototype import (                    # noqa: E402
    _auto_pipe_layer_filter, _closed_open_counts, _is_symbol_only,
    categorize_layers)
from remote30_constants import SYMBOL_LAYER_MIN_CLOSED   # noqa: E402


def _box(layer, x=0.0, y=0.0, r=100.0):
    """닫힌 사각형 하나 — 헤드 기호를 흉내 낸다."""
    return {"t": "PL", "l": layer,
            "p": [[x, y], [x + r, y], [x + r, y + r], [x, y + r], [x, y]]}


def _line(layer, x=0.0, y=0.0, dx=1000.0):
    return {"t": "L", "l": layer, "p": [x, y, x + dx, y]}


def _sym_layer(layer, n=SYMBOL_LAYER_MIN_CLOSED + 5):
    return [_box(layer, x=i * 500.0) for i in range(n)]


def test_이름이_배관이어도_닫힌_기호뿐이면_안_믿는다():
    ents = _sym_layer("SP")
    assert categorize_layers(ents)["SP"] == "OTHER"


def test_선형_도형이_하나라도_있으면_안_내린다():
    """★배관과 기호가 섞인 레이어를 통째로 내리면 진짜 배관을 같이 버린다."""
    ents = _sym_layer("SP") + [_line("SP")]
    assert categorize_layers(ents)["SP"] == "PIPE"


def test_너무_적으면_안_내린다():
    """도형 몇 개짜리가 어쩌다 다 닫혀 있는 경우(범례 한둘)까지 내리지 않는다."""
    ents = _sym_layer("SP", n=SYMBOL_LAYER_MIN_CLOSED - 1)
    assert categorize_layers(ents)["SP"] == "PIPE"


def test_배관_아닌_레이어는_손대지_않는다():
    """PIPE 로 읽힌 것만 대상이다 — 다른 카테고리를 재분류하지 않는다."""
    ents = _sym_layer("HEAD") + _sym_layer("TEX")
    got = categorize_layers(ents)
    assert got["HEAD"] == "HEAD" and got["TEX"] == "TEXT"


def test_헤드라고_말하지_않는다():
    """★아는 것은 «닫힌 기호» 까지다. HEAD 로 올리면 없는 확신을 만든다."""
    assert categorize_layers(_sym_layer("SP"))["SP"] != "HEAD"


def test_자동_배관_필터도_같은_사실을_본다():
    """화면 추천과 추출이 다른 말을 하면 사람은 화면을 믿는다."""
    ents = _sym_layer("SP") + [_line("LSP", y=i * 10.0) for i in range(5)]
    got = _auto_pipe_layer_filter(ents)
    assert "LSP" in got and "SP" not in got, sorted(got)


def test_닫힘_판정은_그래프와_같은_자를_쓴다():
    """따로 재면 언젠가 한쪽만 바뀐다 — 첫점≈끝점, 같은 허용치."""
    from remote30_constants import CLOSED_PL_TOL_MM
    almost = {"t": "PL", "l": "X",
              "p": [[0, 0], [100, 0], [100, 100],
                    [0, CLOSED_PL_TOL_MM * 0.5]]}
    far = {"t": "PL", "l": "Y",
           "p": [[0, 0], [100, 0], [100, 100],
                 [0, CLOSED_PL_TOL_MM * 5]]}
    got = _closed_open_counts([almost, far])
    assert got["X"] == (1, 0), "허용치 안인데 열린 것으로 셌다"
    assert got["Y"] == (0, 1), "허용치 밖인데 닫힌 것으로 셌다"


def test_원도_닫힌_도형이다():
    got = _closed_open_counts([{"t": "C", "l": "S", "c": [0, 0], "r": 50}])
    assert got["S"] == (1, 0)


@pytest.mark.parametrize("counts, want", [
    ((SYMBOL_LAYER_MIN_CLOSED, 0), True),
    ((SYMBOL_LAYER_MIN_CLOSED, 1), False),
    ((SYMBOL_LAYER_MIN_CLOSED - 1, 0), False),
    ((0, 0), False),
])
def test_판정_경계(counts, want):
    assert _is_symbol_only(counts) is want


# ── 실도면 — 교정에 쓴 4장 중 있는 것만 ──────────────────────────
_REAL = [
    ("계통도", ["LSP", "HSP"], ["SP"]),
    ("기계실", ["-소화(SP-고)", "-소화(SP-저)"], []),
    ("평면도", ["-소화(SP가지관)", "SP 후렉시블"], []),
]


@pytest.mark.parametrize("name, keep, drop", _REAL)
def test_실도면에서_진짜_배관은_안_내려간다(name, keep, drop):
    """★교정 4장 실측 — 진짜 배관 레이어는 전부 닫힘 0% 라 안전하다."""
    path = os.path.join(_ROOT, "routes", "제출용[최종]",
                        f"1. 입력도면 대명동 단위세대 {name}.dxf")
    if not os.path.isfile(path):
        pytest.skip(f"표본 도면 없음: {name}")
    from routes.module_f.subdrawing import parse_subdrawing
    ents, _ = parse_subdrawing(path)
    cats = categorize_layers(ents)
    for lay in keep:
        assert cats.get(lay) == "PIPE", f"{lay} 이 배관에서 내려갔다"
    for lay in drop:
        assert cats.get(lay) == "OTHER", f"{lay} 이 아직 배관으로 읽힌다"


def test_그래프는_한_글자도_안_바뀐다():
    """★이 교정의 값은 «말» 에 있지 «그래프» 에 있지 않다.

    닫힌 PL 은 `_build_graph` 의 컷에서 이미 잘리고 있었다. 그러니 SP 를
    빼도 경로 그래프는 같아야 한다 — 다르면 배관을 같이 버린 것이다.
    """
    path = os.path.join(_ROOT, "routes", "제출용[최종]",
                        "1. 입력도면 대명동 단위세대 계통도.dxf")
    if not os.path.isfile(path):
        pytest.skip("표본 도면 없음")
    from remote30_prototype import build_system_graph
    from routes.module_f.subdrawing import parse_subdrawing
    ents, _ = parse_subdrawing(path)
    _g, _e, with_gate = build_system_graph(ents, force_connect=False)
    _g2, _e2, raw = build_system_graph(
        ents, layer_filter={"HSP", "LSP", "SP", "감압밸브"},
        force_connect=False)
    for k in ("node_count", "edge_count", "components_before_bridge"):
        assert with_gate[k] == raw[k], (k, with_gate[k], raw[k])
