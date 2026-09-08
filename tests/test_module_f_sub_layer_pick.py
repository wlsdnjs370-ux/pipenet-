# -*- coding: utf-8 -*-
"""[계통도 꼬임] 경로가 **계통 사이를 넘나들지 않는다**.

■ 사용자 지적

  「계통도 경로 추출이 그래도 일단 그 레이어 안에서 색깔별로 최소 길이가
  나와야 되는데, 뭔가 경로가 꼬여서 추적이 된다. 기계실도 문제가 있는
  것으로 판단되면 조치를 취하고.」

■ 실측으로 잡은 것 (대명동, `scripts/_probe_sub_bridge_path.py`)

  레이어 필터를 안 걸면 엔진의 이름 사전이 배관 레이어를 **여러 장 한꺼번에**
  고른다. 그 셋이 한 그래프에 들어가니 최단경로가 계통을 넘나든다.

    자동(섞음) — LSP → HSP → LSP · 넘나듦 4회 · 연장 126.3 m
    HSP 만     — HSP            · 넘나듦 0회 · 연장 118.0 m
    기계실 자동 — -소화(SP-저) → -소화(SP-고) · 넘나듦 2회

  범인이 아닌 것도 적어 둔다: 추정 이음(허용오차 다리)은 경로 126.3 m 중
  3곳 1.0 m(1%)뿐이고, 그것을 뒤로 미뤄도 경로가 그대로였다. 다리를 손대는
  대신 **계통을 하나로 좁힌다.**

■ 여기서 지키는 것

  ⑴ 두 점이 한 레이어 안에 있으면 그 레이어를 고른다.
  ⑵ 못 담으면 **고르지 않는다** — 억지로 좁히면 엉뚱한 계통을 돈다.
  ⑶ 사람이 직접 고른 뒤에는 자동이 그것을 덮지 않는다.
  ⑷ 미리보기와 추출이 같은 레이어를 쓴다(세션 한 곳에 남는다).
"""
from __future__ import annotations

import math
import os

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _grid(layer, x0, y0, n=14, step=1000.0):
    """한 계통 — 세로 한 줄에 가로 가지 n개."""
    ents = [{"t": "L", "l": layer,
             "p": [x0, y0, x0, y0 + step * n]}]
    for i in range(n):
        y = y0 + step * i
        ents.append({"t": "L", "l": layer, "p": [x0, y, x0 + step * 2, y]})
    return ents


@pytest.fixture()
def two_systems():
    """서로 안 닿는 두 계통 — 다만 **가깝다**(허용오차 다리가 붙을 만큼)."""
    return _grid("HSP", 0.0, 0.0) + _grid("LSP", 300.0, 0.0)


def _pick(ents, a, b, **kw):
    from routes.module_f.subdrawing import pick_system_layer
    return pick_system_layer(ents, a, b, **kw)


def test_두_점이_한_계통에_있으면_그_계통을_고른다(two_systems):
    nm, diag = _pick(two_systems, (0.0, 0.0), (0.0, 5000.0))
    assert nm == "HSP", diag
    nm2, diag2 = _pick(two_systems, (300.0, 0.0), (300.0, 5000.0))
    assert nm2 == "LSP", diag2


def test_짧은_경로보다_찍은_자리가_먼저다(two_systems):
    """★고칠 때 실제로 틀렸던 자리다.

    두 계통이 300mm 로 나란히 붙어 있으면 클릭이 옆 계통의 허용거리 안에도
    든다. 그때 «경로가 짧은 쪽» 만 보면 찍은 점 위의 계통(HSP·13.0 m)을 두고
    옆 계통(LSP·9.0 m)을 고른다 — 사람이 찍은 자리가 먼저다.
    """
    nm, diag = _pick(two_systems, (0.0, 0.0), (0.0, 5000.0))
    got = {r["layer"]: r for r in diag["candidates"]}
    assert got["LSP"]["ok"] is True, "옆 계통도 허용거리 안이라야 시험이 된다"
    assert got["LSP"]["path_m"] < got["HSP"]["path_m"], "짧은 쪽이 옆 계통이라야"
    assert nm == "HSP", diag


def test_고른_근거를_함께_돌려준다(two_systems):
    """조용히 바꾸지 않는다 — 화면이 무엇을 왜 골랐는지 말할 수 있어야 한다."""
    nm, diag = _pick(two_systems, (0.0, 0.0), (0.0, 5000.0))
    assert nm and diag.get("reason")
    got = {r["layer"]: r for r in diag["candidates"]}
    assert set(got) >= {"HSP", "LSP"}
    assert got["HSP"]["ok"] is True
    assert got["HSP"]["path_m"] > 0


def test_두_점을_한_계통에_못_담으면_고르지_않는다(two_systems):
    """★억지로 좁히면 엉뚱한 계통 안에서 먼 길을 돈다 — 섞은 채로 둔다."""
    # 두 점이 서로 다른 계통에 있고, 상대 계통에서는 허용거리 밖이다.
    nm, diag = _pick(two_systems, (0.0, 0.0), (300.0, 5000.0),
                     snap_tolerance_mm=50.0)
    assert nm is None, diag
    assert "못" in diag["reason"] or "섞" in diag["reason"]


def test_섞인_레이어가_없으면_좁힐_것도_없다():
    ents = _grid("HSP", 0.0, 0.0)
    nm, diag = _pick(ents, (0.0, 0.0), (0.0, 5000.0))
    assert nm is None and diag["candidates"] == []


def _path_layers(ents, layers, a, b):
    """경로가 지나는 레이어 열 — 넘나듦을 세기 위한 자."""
    from remote30_prototype import (_nearest_graph_node, _shortest_path,
                                    build_system_graph)
    g, el, _st = build_system_graph(ents, layer_filter=layers,
                                    force_connect=True)
    own = {}
    for nm in ("HSP", "LSP"):
        g1, el1, _s = build_system_graph(ents, layer_filter={nm},
                                         force_connect=False)
        own[nm] = set(el1)
    na = _nearest_graph_node(g, a)
    nb = _nearest_graph_node(g, b)
    p = _shortest_path(g, el, na, nb)
    seq = []
    for u, v in zip(p or (), (p or ())[1:]):
        k = (min(u, v), max(u, v))
        hit = [nm for nm, ks in own.items() if k in ks]
        s = hit[0] if len(hit) == 1 else "?"
        if not seq or seq[-1] != s:
            seq.append(s)
    return seq


def test_좁히면_한_계통_안에서만_돈다(two_systems):
    """★이것이 «꼬임» 의 정체다 — 좁히기 전후를 같은 자로 잰다."""
    a, b = (0.0, 0.0), (2000.0, 5000.0)
    mixed = _path_layers(two_systems, None, a, b)
    one = _path_layers(two_systems, {"HSP"}, a, b)
    assert set(one) <= {"HSP"}, one
    # 섞으면 넘나들 수 있다(도면에 따라 다르므로 «좁힌 쪽이 더 낫다» 만 본다).
    assert len([s for s in one if s != "?"]) <= len(
        [s for s in mixed if s != "?"])


def test_추정_이음도_세어_내보낸다(two_systems):
    """허용오차 다리는 «알고리즘이 이은» 것이다 — 세지 않으면 안 보인다."""
    from remote30_prototype import build_system_graph
    _g, _el, st = build_system_graph(two_systems, layer_filter={"HSP"},
                                     force_connect=False)
    assert "tolerance_bridges" in st and "tolerance_bridge_edges" in st
    assert st["tolerance_bridges"] == len(st["tolerance_bridge_edges"])


# ─────────────────────────────── 라우트: 미리보기와 추출이 같은 레이어를 쓴다
def _app():
    from routes.module_f import api_sub
    from flask import Flask
    app = Flask(__name__)
    api_sub.register(app)
    return app


def test_사람이_고르면_자동이_덮지_않는다():
    """★사람 결정을 다음 클릭에 기계가 덮으면 «고를 수 있다» 가 거짓이 된다."""
    from routes.module_f import api_sub
    sess = {}
    api_sub._layers(sess, {"layers": ["LSP"]})
    assert sess["sub_layers"] == ["LSP"]
    assert sess["sub_layers_auto"] is False


def test_아무것도_안_고른_세션은_자동_좁히기가_열려_있다():
    from routes.module_f import api_sub
    sess = {}
    assert api_sub._layers(sess, {}) is None
    assert sess.get("sub_layers_auto", True) is True


def test_좁힌_레이어가_세션에_남아_추출도_같은_것을_쓴다(two_systems):
    """미리보기 그래프와 추출이 다른 레이어를 쓰면 화면이 거짓말을 한다."""
    from routes.module_f import api_sub
    from routes.module_f.subdrawing import pick_system_layer
    sess = {"entities": two_systems}
    nm, _diag = pick_system_layer(two_systems, (0.0, 0.0), (0.0, 5000.0))
    sess["sub_layers"] = [nm]
    sess["sub_layers_auto"] = True
    # 추출이 부르는 바로 그 함수 — 같은 값을 받아야 한다.
    assert api_sub._layers(sess, {}) == {nm}


def test_멀리_찍어도_더_붙은_계통을_고른다(two_systems):
    """★«허용거리 안이면 된다» 로 걸었다가 한 번 틀렸다 — 그 자리를 못박는다.

    추출은 클릭을 거리 제한 없이 가장 가까운 절점에 붙인다. 화면에서 사람이
    찍는 자리는 배관에서 수십 m 떨어져 있는 것이 보통이라(실측: 계통도 632 m),
    절대 허용거리로 후보를 자르면 좁히기가 **한 번도 안 걸린다.** 자는
    «어느 계통에 더 붙었나» 라는 견줌이어야 한다.
    """
    far = 100_000.0
    nm, diag = _pick(two_systems, (far, 0.0), (far, 5000.0),
                     snap_tolerance_mm=10.0)
    assert nm == "LSP", diag          # x=300 쪽이 x=0 쪽보다 가깝다
    assert all(r["ok"] for r in diag["candidates"]), diag


def test_두_점이_서로_다른_계통에_붙으면_안_좁힌다(two_systems):
    """한쪽은 HSP 위, 한쪽은 LSP 위 — 어느 하나로 좁히면 반드시 틀린다."""
    nm, diag = _pick(two_systems, (0.0, 0.0), (2300.0, 5000.0))
    assert nm is None, diag
    assert "다른 계통" in diag["reason"], diag["reason"]


def test_거리는_실배관으로_잰다(two_systems):
    """고른 근거의 «연장» 은 직선이 아니라 배관을 따라간 값이어야 한다."""
    nm, diag = _pick(two_systems, (0.0, 0.0), (2000.0, 5000.0))
    row = [r for r in diag["candidates"] if r["layer"] == nm][0]
    straight = math.hypot(2000.0, 5000.0) / 1000.0
    assert row["path_m"] >= straight - 1e-6
