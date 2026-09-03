# -*- coding: utf-8 -*-
"""좌표계 계약 — 모듈 F/G 파이프라인의 세 공간과, 그 사이를 건너는 유일한 문.

■ 세 공간

  A. 판 mm     — DXF 원좌표. 찍기판·손질판·payload 의 ho/sources/valve_picks.
  B. kfp m     — 수리계산 그래프. A 에서 «도면 왼쪽아래(minx,miny)를 (1m,1m)
                 에 놓고 1/1000» 한 것. 문은 xf_mm_to_m 하나다.
  C. 표시 좌표 — PIPENET 스키매틱 캔버스. B 에서 «bbox 중심→(0,0), 최장축→
                 canvas_units» 로 정규화하고, 아이소 보기면 그 위에 베이크.
                 문은 display_tables 하나다(미리보기·파일 방출이 같은 사본).

■ 계약

  · B 로 들어가려면 origin_mm 이 필요하다. planar 산출 payload 가 싣고 다닌다
    — (minx,miny) 를 잃으면 B 좌표를 A 로 되돌릴 수 없다.
    ★origin_mm 을 (0,0) 으로 «가정» 하면 안 된다. 실측 도면은 원점에서 수십만
      mm 떨어져 있어, 그 가정으로 되돌린 좌표는 도면 바깥의 허공이다.
  · origin_mm=None 은 «입력이 이미 kfp 단위» 라는 뜻의 스모크 전용 모드다.
    제품 경로에서는 planar 가 항상 origin_mm 을 채우고 engine 이 백필한다.
  · C 는 표시 전용이다 — elevation 은 정규화가 건드리지 않는다. 그것은 표시가
    아니라 수리계산 입력이다(§G12). iso/iso_z_scale/canvas_units 를 어떻게
    두어도 계산 결과는 같아야 한다.
"""
from __future__ import annotations

import os
import random
import sys
import types

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
if _G not in sys.path:
    sys.path.insert(0, _G)

from services.cad_import.convert.main_walk import (   # noqa: E402
    ho_to_kfp_units, xf_mm_to_m)
from services.cad_import.design.emit import display_tables   # noqa: E402
from services.cad_import.design.sdf_post import (     # noqa: E402
    normalize_node_coords)


# ═════════════════════════════ A → B (판 mm → kfp m)
def test_왼쪽아래가_1m_1m_에_놓인다():
    minx, miny = 412_345.6, -98_765.4          # 실측 도면처럼 원점에서 먼 값
    assert xf_mm_to_m(minx, miny, minx, miny) == (1.0, 1.0)
    # 1000mm = 1m — 배율은 1/1000 하나뿐이다.
    assert xf_mm_to_m(minx + 1000.0, miny + 2500.0, minx, miny) == (2.0, 3.5)


def test_B에서_A로는_origin_mm_로만_돌아온다():
    """역변환 mm = (m − 1)·1000 + origin. origin 을 잃으면 못 돌아온다.

    ★(0,0) 가정의 말로 — 검증 스크립트에서 실제로 저지른 실수다: 되돌린
      좌표가 도면에서 40만 mm 떨어진 허공에 떨어져, 헤드가 «이상한 위치»
      에 있다고 오판하게 만든다.
    """
    rng = random.Random(20260903)
    minx, miny = 403_210.0, 187_654.0
    for _ in range(50):
        x = minx + rng.uniform(0, 90_000)
        y = miny + rng.uniform(0, 90_000)
        mx, my = xf_mm_to_m(x, y, minx, miny)
        back = ((mx - 1.0) * 1000.0 + minx, (my - 1.0) * 1000.0 + miny)
        assert abs(back[0] - x) < 1e-6 and abs(back[1] - y) < 1e-6
    mx, my = xf_mm_to_m(minx + 5000.0, miny, minx, miny)
    wrong = (mx - 1.0) * 1000.0 + 0.0          # origin=(0,0) 가정
    assert abs(wrong - (minx + 5000.0)) > 100_000, \
        "origin 없이도 얼추 맞는다면 이 시험 도면이 원점에 너무 가깝다"


def test_ho_는_중심도_반지름도_m_가_된다():
    out = ho_to_kfp_units(
        [{"cx": 1000.0, "cy": 2000.0, "r": 1500.0, "k": "알람밸브"}],
        0.0, 0.0)
    assert (out[0]["cx"], out[0]["cy"]) == (2.0, 3.0)
    assert out[0]["r"] == 1.5
    assert out[0]["k"] == "알람밸브", "변환이 좌표 아닌 칸을 건드렸다"


def test_origin_이_없으면_이미_kfp_단위다():
    """스모크 전용 모드 — mm 를 m 인 척 섞는 게 아니라, «변환 없음» 이 계약이다.

    제품 경로에서는 planar 가 origin_mm=(minx,miny) 를 payload 에 항상 싣고,
    engine._load_kfp 가 비어 있으면 백필한다. 이 모드가 제품에서 밟히면
    좌표가 천 배로 틀어지므로, 새 입력 경로를 만들 때는 origin 을 함께 나를 것.
    """
    from services.cad_import.convert.engine import _prepare_ho
    ho = [{"cx": 2.0, "cy": 3.0, "r": 0.1, "k": "메인"}]
    got, _src = _prepare_ho({"ho": ho}, None)
    assert got == ho                            # 그대로 — 변환하지 않는다
    got2, _src = _prepare_ho({"ho": [{"cx": 1000.0, "cy": 0.0, "r": 100.0}],
                              "origin_mm": (0.0, 0.0)}, None)
    assert (got2[0]["cx"], got2[0]["cy"], got2[0]["r"]) == (2.0, 1.0, 0.1)


# ═════════════════════════════ B → C (kfp m → 표시)
def _tables(nodes):
    return types.SimpleNamespace(nodes=nodes, pipes=[])


def test_정규화는_중심을_0으로_최장축을_canvas로_고도는_그대로():
    t = _tables([
        {"label": "N1", "x": 10.0, "y": 20.0, "elevation": 3.2},
        {"label": "N2", "x": 70.0, "y": 50.0, "elevation": 0.0},
    ])
    scale = normalize_node_coords(t, canvas_units=3000.0)
    xs = [n["x"] for n in t.nodes]
    ys = [n["y"] for n in t.nodes]
    assert abs(min(xs) + max(xs)) < 1e-9 and abs(min(ys) + max(ys)) < 1e-9
    assert abs((max(xs) - min(xs)) - 3000.0) < 1e-9      # 최장축 = x(60)
    assert abs((max(ys) - min(ys)) - 1500.0) < 1e-9      # 비율 보존 (30/60)
    assert scale == 50.0
    # ★elevation 은 수리계산 입력이다 — 표시 정규화가 손대면 계산이 틀어진다.
    assert [n["elevation"] for n in t.nodes] == [3.2, 0.0]


def test_display_tables_는_사본이다_평면보기는_세움도_없다():
    """부른 쪽의 표를 바꾸면, 같은 표로 두 번 그릴 때 두 번 정규화된다."""
    src = _tables([
        {"label": "N1", "x": 1.0, "y": 1.0, "elevation": 0.0},
        {"label": "N2", "x": 9.0, "y": 5.0, "elevation": 0.0},
    ])
    before = [dict(n) for n in src.nodes]
    view, stood = display_tables(src, iso=False, canvas_units=3000.0)
    assert src.nodes == before, "display_tables 가 원본 표를 바꿨다"
    assert stood is None                        # 평면 보기 — 베이크 없음
    assert view.nodes is not src.nodes
    xs = [n["x"] for n in view.nodes]
    assert abs((max(xs) - min(xs)) - 3000.0) < 1e-9
