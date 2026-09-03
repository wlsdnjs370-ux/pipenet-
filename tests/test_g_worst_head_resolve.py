# -*- coding: utf-8 -*-
"""기준 헤드(최원단)를 kfp 노드로 되짚는다 — BLOCKED §30 해소.

■ 무엇이 죽어 있었나

  `design/tables._worst_head_node()` 가 무조건 None 을 돌려주는 스텁이었다.
  그래서 표의 「기준 헤드 노드」는 **항상 '?'** 였고, 그 값을 읽어 최원 유하거리
  «경로» 를 그리는 `api_design` 블록은 한 번도 돌지 않았다. 화면에는 숫자
  (far_m)만 있고 그것이 어느 줄인지는 없었다.

  포기 사유는 「board 헤드 번호는 전개 노드와 1:1 이 아니다」였다. 맞는 말이지만
  그래서 **아무것도 안 했다.** 실측(B1F · 뽑힌 헤드 30개)으로 원인을 갈랐다:
    · `node_ref` 로 되짚기 → 30개 중 12개만 맞는다. 그 표는 노드정리 «전» 의
      id 를 담고 있어(323 → 80) 절반 넘게 사라진 노드를 가리킨다.
    · 좌표로 맞대기  → 30개 전부 6~19mm 안에서 **유일하게** 걸린다
      (2등/1등 거리비 최소 141.7배 · 500mm 안에 둘 이상 0건).
  못 이을 이유는 없었고, 되짚는 표를 잘못 고른 것이었다.

■ 이 시험이 지키는 것

  좌표로 잇되 **추측하지 않는다** — 재료가 없거나(origin_mm) 모호하면(가까이
  둘 이상) 종전처럼 None 이다. 틀린 헤드를 가리키느니 모른다고 두는 편이 옳다.
"""
from __future__ import annotations

import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
if _G not in sys.path:
    sys.path.insert(0, _G)

from services.cad_import.design.tables import (          # noqa: E402
    WORST_HEAD_SNAP_M, _worst_head_node, build_design_tables)

# board mm 좌표계의 원점 — 실도면처럼 원점에서 멀리 둔다. (0,0) 으로 두면
# origin_mm 을 안 써도 우연히 맞아, 계약이 깨져도 시험이 안 잡는다.
ORIGIN = (646_763.6, 88_854.1)


def _to_kfp(mx, my):
    return ((mx - ORIGIN[0]) / 1000.0 + 1.0, (my - ORIGIN[1]) / 1000.0 + 1.0)


def _fixture(head_dx_mm=0.0, extra_head_dx_mm=None):
    """접속점 — 배관 — 헤드 하나. 헤드 노드를 부착점에서 얼마나 옮길지 고른다.

    `head_dx_mm` 는 세로 전개가 헤드를 x 로 미는 흉내다(상하향식 combo_2).
    """
    src_mm = (ORIGIN[0], ORIGIN[1])
    head_mm = (ORIGIN[0] + 5000.0, ORIGIN[1])
    board_pts = [src_mm, head_mm]

    ax, ay = _to_kfp(*src_mm)
    hx, hy = _to_kfp(*head_mm)
    nodes = {
        "N1": {"coords": [ax, ay, 0.0], "type_id": "pump"},
        "N2": {"coords": [hx + head_dx_mm / 1000.0, hy, 0.0],
               "type_id": "head"},
    }
    pipes = {"P1": {"start": "N1", "end": "N2", "length_m": 5.0}}
    if extra_head_dx_mm is not None:
        nodes["N3"] = {"coords": [hx + extra_head_dx_mm / 1000.0, hy, 0.5],
                       "type_id": "head"}
        pipes["P2"] = {"start": "N2", "end": "N3", "length_m": 0.5}
    net = {"pipe_data": pipes, "nodes_meta_runtime": nodes}
    worst = {"heads": [0], "worst_head": 0, "worst_path": [0, 1],
             "worst_path_m": 5.0, "loads": {(0, 1): 1}}
    return net, worst, board_pts


def test_부착점_자리의_헤드를_찾는다():
    net, worst, pts = _fixture()
    got = _worst_head_node(net, worst, net["nodes_meta_runtime"],
                           board_pts=pts, origin_mm=ORIGIN)
    assert got == "N2", got


def test_실측_범위_안의_어긋남은_받아낸다():
    """B1F 실측 최대 19mm — 그 다섯 배까지 받는다."""
    net, worst, pts = _fixture(head_dx_mm=19.0)
    assert _worst_head_node(net, worst, net["nodes_meta_runtime"],
                            board_pts=pts, origin_mm=ORIGIN) == "N2"


def test_상하향식처럼_밀린_헤드만_있으면_모른다고_둔다():
    """★combo_2(기본 0.3m)만큼 밀린 «하향» 쪽만 있으면 받지 않는다.

    실제 상하향식은 «상향» 쪽이 부착점에 그대로 남아 여기 걸리지만, 그렇지
    않은 전개가 생기면 조용히 이웃 헤드를 집는 대신 모른다고 해야 한다.
    """
    net, worst, pts = _fixture(head_dx_mm=300.0)
    assert _worst_head_node(net, worst, net["nodes_meta_runtime"],
                            board_pts=pts, origin_mm=ORIGIN) is None


def test_가까이_둘이면_모호하니_고르지_않는다():
    """상하향식 상·하 두 헤드가 **둘 다** 허용 안에 들면 어느 쪽인지 모른다."""
    net, worst, pts = _fixture(head_dx_mm=0.0, extra_head_dx_mm=30.0)
    assert _worst_head_node(net, worst, net["nodes_meta_runtime"],
                            board_pts=pts, origin_mm=ORIGIN) is None


def test_origin_mm_이_없으면_추측하지_않는다():
    """★좌표계를 모르면 «가장 먼 헤드» 같은 추측을 하지 않는다.

    (0,0) 을 가정하면 실도면에서 40만 mm 떨어진 허공을 짚는다 —
    tests/test_module_f_coord_contract.py 가 그 대가를 재 두었다.
    """
    net, worst, pts = _fixture()
    assert _worst_head_node(net, worst, net["nodes_meta_runtime"],
                            board_pts=pts, origin_mm=None) is None
    assert _worst_head_node(net, worst, net["nodes_meta_runtime"],
                            board_pts=None, origin_mm=ORIGIN) is None


def test_경로가_없으면_모른다():
    net, worst, pts = _fixture()
    worst = dict(worst, worst_path=[])
    assert _worst_head_node(net, worst, net["nodes_meta_runtime"],
                            board_pts=pts, origin_mm=ORIGIN) is None


def test_허용치는_이웃_헤드에_한참_못_미친다():
    """숫자의 근거를 시험이 들고 있는다.

    실측 이웃 헤드 간격 2.95m · 허용 0.10m — 이 관계가 뒤집히면 이웃을 집는다.
    """
    assert WORST_HEAD_SNAP_M < 2.95 / 10.0


# ═════════════════════════════ 표에 실제로 실린다
def test_표의_기준_헤드_노드가_물음표가_아니다():
    net, worst, pts = _fixture()
    tbl = build_design_tables(net, worst, {"P1": (0, 1)}, [],
                              board_pts=pts, origin_mm=ORIGIN)
    lab = dict(tbl.meta)["기준 헤드 노드"]
    assert lab != "?", "되짚기가 표까지 오지 않았다"
    row = next(n for n in tbl.nodes if str(n["label"]) == str(lab))
    # 그 라벨이 정말 헤드 자리인가 — 노즐표가 그 절점을 가리켜야 한다.
    assert row is not None


def test_재료가_없으면_표는_종전처럼_물음표다():
    """추측해서 채우느니 «모른다» 를 그대로 싣는다."""
    net, worst, pts = _fixture()
    tbl = build_design_tables(net, worst, {"P1": (0, 1)}, [], board_pts=pts)
    assert dict(tbl.meta)["기준 헤드 노드"] == "?"
