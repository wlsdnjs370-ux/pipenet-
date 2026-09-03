# -*- coding: utf-8 -*-
"""[F-10e] 아이소 밑그림 — board 평면을 설계 좌표계에 «어긋남 없이» 얹는다.

■ 오래 «불가» 였던 것

  지시서는 밑그림을 아이소 아래에 깔라며 「좌표계가 1픽셀이라도 어긋나면 밑그림의
  의미가 없다」고 못 박았고, BLOCKED §17 은 그 전제가 성립하지 않는다고 기록했다
  (board → 설계에 전역 변환이 없다 · 최대 잔차 도면의 9.3%).

  그 9.3% 는 **기하가 아니라 짝짓기가 만든 값**이었다. 대응을 `edge_ref` 로
  잡았는데 그 표는 노드정리 «전» 을 가리킨다 — 믿을 수 있는 대응으로 다시 재면
  0.018% 다(BLOCKED §17 정정 · `scripts/_probe_f10e_affine2.py`).

■ 이 시험이 지키는 것

  변환을 **엔진이 쓰는 수 그대로** 만든다(별도 수학 금지). 그러므로 board 점을
  그 식으로 옮기면 표의 절점과 같은 자리에 떨어져야 한다 — 평면에서도, 아이소
  에서도.
"""
from __future__ import annotations

import math
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
for _p in (_ROOT, _G):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.design.emit import display_tables          # noqa: E402
from services.cad_import.design.sdf_post import (                   # noqa: E402
    COS30, SIN30, node_norm_params, normalize_node_coords)
from services.cad_import.design.tables import build_design_tables    # noqa: E402

ORIGIN = (646_763.6, 88_854.1)      # 실도면처럼 원점에서 먼 값


def _net():
    """접속점 — 수평배관 — 헤드. 헤드만 표고가 다르다(전개가 세로로 올린 것)."""
    def k(mx, my):
        return ((mx - ORIGIN[0]) / 1000.0 + 1.0, (my - ORIGIN[1]) / 1000.0 + 1.0)

    a = k(ORIGIN[0], ORIGIN[1])
    b = k(ORIGIN[0] + 8000.0, ORIGIN[1])
    c = k(ORIGIN[0] + 8000.0, ORIGIN[1] + 5000.0)
    nodes = {
        "N1": {"coords": [a[0], a[1], 0.0], "type_id": "pump"},
        "N2": {"coords": [b[0], b[1], 0.0], "type_id": "base"},
        "N3": {"coords": [c[0], c[1], 0.6], "type_id": "head"},
    }
    pipes = {"P1": {"start": "N1", "end": "N2", "length_m": 8.0},
             "P2": {"start": "N2", "end": "N3", "length_m": 5.0}}
    board_pts = [(ORIGIN[0], ORIGIN[1]), (ORIGIN[0] + 8000.0, ORIGIN[1]),
                 (ORIGIN[0] + 8000.0, ORIGIN[1] + 5000.0)]
    worst = {"heads": [0], "worst_head": 0, "worst_path": [0, 1, 2],
             "loads": {(0, 1): 1}}
    return ({"pipe_data": pipes, "nodes_meta_runtime": nodes},
            worst, board_pts)


def _xf(mx, my, u):
    """화면이 쓰는 그 식 — `static/module_f.js: drawUnderlay` 와 같아야 한다."""
    nx = u["k"] * mx + u["tx"]
    ny = u["k"] * my + u["ty"]
    if not u["iso"]:
        return nx, ny
    dz = (u["e"] - u["e_ref"]) * u["lift"]
    return (nx - ny) * u["cos30"], (nx + ny) * u["sin30"] + dz


def _build(iso):
    from routes.module_f.api_design import _underlay_xf
    net, worst, pts = _net()
    tbl = build_design_tables(net, worst, {"P1": (0, 1), "P2": (1, 2)}, [],
                              board_pts=pts, origin_mm=ORIGIN)
    view, stood = display_tables(tbl, iso=iso, canvas_units=3000.0)
    sess = {"design": {"got": {"origin_mm": ORIGIN}}}
    u = _underlay_xf(sess, view, stood, {"iso": iso})
    at = {str(n["label"]): n for n in view.nodes}
    return view, stood, u, at, pts


def test_평면_보기에서_board_점이_표_절점에_떨어진다():
    _v, _s, u, at, pts = _build(False)
    root = next(n for n in at.values() if str(n.get("io_node")) == "Input")
    X, Y = _xf(pts[0][0], pts[0][1], u)
    d = math.hypot(X - float(root["x"]), Y - float(root["y"]))
    assert d < 1e-6, f"평면에서 {d} 만큼 어긋났다"


def test_아이소에서도_board_평면이_제자리에_떨어진다():
    """★§17 이 「불가」라 했던 바로 그것."""
    _v, _s, u, at, pts = _build(True)
    assert u["iso"] is True
    root = next(n for n in at.values() if str(n.get("io_node")) == "Input")
    X, Y = _xf(pts[0][0], pts[0][1], u)
    d = math.hypot(X - float(root["x"]), Y - float(root["y"]))
    assert d < 1e-6, f"아이소에서 {d} 만큼 어긋났다"


def test_밑그림은_접속점_표고에_눕는다():
    """board 는 평면도다 — 수평망 높이(=접속점 표고)에 놓여야 한다.

    ★lift 영점(e_ref)에 두면 안 된다. 영점은 «보기» 의 기준일 뿐 board 평면이
      실제로 놓인 높이가 아니다. 실측(B1F)으로 그 차 0.300m × lift 2,821.9 =
      846.9 단위 = 한 변의 21.8% 가 그대로 어긋났다.
    """
    _v, stood, u, at, _pts = _build(True)
    root = next(n for n in at.values() if str(n.get("io_node")) == "Input")
    assert u["e"] == float(root["elevation"])
    assert u["e_ref"] == float(stood["e_ref"])
    # 이 판은 헤드가 0.6, 나머지가 0 이라 영점과 평면이 «다르다» — 그래야 이
    # 시험이 둘을 구분한다.
    assert abs(u["e"] - u["e_ref"]) > 1e-9, "영점과 평면이 같아 구분이 안 된다"


def test_변환은_엔진이_쓰는_수_그대로다():
    """별도 수학 금지 — 정규화 파라미터가 엔진 것과 한 글자도 다르지 않아야 한다."""
    net, worst, pts = _net()
    tbl = build_design_tables(net, worst, {"P1": (0, 1), "P2": (1, 2)}, [],
                              board_pts=pts, origin_mm=ORIGIN)
    cx, cy, scale = node_norm_params(tbl, canvas_units=3000.0)
    view, _ = display_tables(tbl, iso=False, canvas_units=3000.0)
    assert (view.norm["cx"], view.norm["cy"], view.norm["scale"]) == \
        (cx, cy, scale)
    assert (view.norm["cos30"], view.norm["sin30"]) == (COS30, SIN30)


def test_정규화는_파라미터_함수와_같은_결과를_낸다():
    """`node_norm_params` 를 갈라 놓고도 `normalize_node_coords` 가 안 변했는가.

    두 함수가 갈라지면 밑그림만 어긋난다 — 표는 멀쩡해 보여 눈치채기 어렵다.
    """
    import types
    nodes = [{"label": "1", "x": 1000, "y": 2000, "elevation": 0.0},
             {"label": "2", "x": 5000, "y": 4000, "elevation": 0.0}]
    t = types.SimpleNamespace(nodes=[dict(n) for n in nodes], pipes=[])
    cx, cy, scale = node_norm_params(t, canvas_units=3000.0)
    got = normalize_node_coords(t, canvas_units=3000.0)
    assert got == scale
    for before, after in zip(nodes, t.nodes):
        assert after["x"] == (before["x"] - cx) * scale
        assert after["y"] == (before["y"] - cy) * scale


def test_화면_식이_서버_식과_한_글자도_다르지_않다():
    """★밑그림이 조용히 어긋나는 가장 그럴듯한 길 — JS 가 제 식으로 흘러가는 것.

    변환은 서버가 «엔진이 쓰는 수» 를 보내고 화면이 그대로 얹는다. 그런데 그
    «얹는 식» 은 JS 에 손으로 적혀 있어, 언젠가 한쪽만 고쳐질 수 있다. 표는
    멀쩡한데 밑그림만 어긋나므로 눈치채기 어렵다.

    그래서 JS 소스에서 식을 그대로 뽑아 node 로 돌리고, 같은 입력에 대해
    파이썬 셈과 맞대 본다.
    """
    import json
    import re
    import shutil
    import subprocess

    import pytest

    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — JS 식을 돌릴 수 없다")

    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index("function drawUnderlay()")
    body = js[i:js.index("\n  }\n", i)]
    m = re.search(r"const px = \(mx, my\) => \{(.+?)\};", body, re.S)
    assert m, "drawUnderlay 의 좌표식을 못 찾았다 — 이름이 바뀌었나"
    expr = m.group(1)

    u = {"k": 0.017767249, "tx": -11500.25, "ty": 3300.5,
         "cos30": COS30, "sin30": SIN30, "iso": True,
         "lift": 2821.9, "e_ref": 0.3, "e": 0.0}
    pts = [[646763.6, 88854.1], [700000.0, -12345.6], [654763.6, 93854.1]]
    prog = ("const u=%s; const dz=(u.e-u.e_ref)*u.lift;"
            "const px=(mx,my)=>{%s};"
            "console.log(JSON.stringify(%s.map(p=>px(p[0],p[1]))));"
            % (json.dumps(u), expr, json.dumps(pts)))
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True)
    assert out.returncode == 0, out.stderr[:300]
    got = json.loads(out.stdout)

    for p, a in zip(pts, got):
        b = _xf(p[0], p[1], u)
        assert math.hypot(a[0] - b[0], a[1] - b[1]) < 1e-9, (p, a, b)
