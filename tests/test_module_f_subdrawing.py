# -*- coding: utf-8 -*-
"""[H-2 · H-3] 계통도 · 기계실 어댑터 — 특허 S720 · S730.

엔진(모듈 A)은 검증된 것을 그대로 부른다. 여기서 못박는 것은 **어댑터**다:
A 의 entity 목록이 F 캔버스가 읽는 World 모양으로 정확히 옮겨지는가.

어댑터가 틀리면 도면이 안 보이거나 좌표가 어긋난 채 보이고, 그 위에서 사람이
펌프·알람밸브를 찍는다 — 즉 **틀린 좌표로 경로를 뽑는다.** 조용히 그른다.
"""
from __future__ import annotations

import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import pytest  # noqa: E402

from routes.module_f.subdrawing import (  # noqa: E402
    entities_to_world, riser_summary)
from routes.module_f.world import _world_payload  # noqa: E402


@pytest.fixture(scope="module", autouse=True)
def _engine():
    """G 엔진(`services`)을 올린다 — `_world_payload` 가 색 이름을 거기서 읽는다.

    직접 sys.path 를 만지지 않는다: E/G 두 트리의 패키지 이름이 같아 어느 쪽이
    올라오는지가 순서 우연에 걸린다(common._boot 가 그것을 막는다).
    """
    from routes.module_f.common import _boot
    _boot()


def test_선분이_옮겨진다():
    w = entities_to_world([{"t": "L", "l": "PIPE", "p": [0, 0, 100, 50]}])
    assert w.segs == [("PIPE", 7, (0.0, 0.0), (100.0, 50.0))]


def test_폴리선은_마디마다_선분으로_편다():
    w = entities_to_world([
        {"t": "PL", "l": "PIPE", "p": [[0, 0], [10, 0], [10, 10]]},
    ])
    assert len(w.segs) == 2
    assert w.segs[0][2:] == ((0.0, 0.0), (10.0, 0.0))
    assert w.segs[1][2:] == ((10.0, 0.0), (10.0, 10.0))


def test_마디가_하나뿐인_폴리선은_선분이_없다():
    w = entities_to_world([{"t": "PL", "l": "X", "p": [[0, 0]]}])
    assert w.segs == []


def test_원과_호():
    w = entities_to_world([
        {"t": "C", "l": "SYM", "c": [5, 6], "r": 2},
        {"t": "A", "l": "SYM", "c": [1, 2], "r": 3, "a": [0, 90]},
    ])
    assert w.circles == [("SYM", 7, 5.0, 6.0, 2.0)]
    assert w.arcs == [("SYM", 7, 1.0, 2.0, 3.0)]
    assert w.arc_ang == [(0.0, 90.0)]


def test_호의_sweep은_360으로_감싼다():
    """끝각이 시작각보다 작으면 한 바퀴를 돌아온 것이다."""
    w = entities_to_world([{"t": "A", "l": "S", "c": [0, 0], "r": 1,
                            "a": [350, 10]}])
    assert w.arc_ang == [(350.0, 20.0)]


def test_전체가_한바퀴면_360():
    w = entities_to_world([{"t": "A", "l": "S", "c": [0, 0], "r": 1,
                            "a": [0, 360]}])
    assert w.arc_ang == [(0.0, 360.0)]


def test_텍스트는_높이가_양수다():
    """0 이면 치수 텍스트 판독이 그 줄을 버린다(stage1 의 h>0 조건)."""
    w = entities_to_world([{"t": "T", "l": "TXT", "p": [1, 2], "v": "100"}])
    assert len(w.texts) == 1
    lay, col, x, y, h, s = w.texts[0]
    assert (lay, x, y, s) == ("TXT", 1.0, 2.0, "100")
    assert h > 0


def test_치수_텍스트가_그대로_읽힌다():
    """어댑터를 거친 뒤에도 관경 판독이 되는가 — 계통도 관경의 근거다."""
    from services.cad_import.design.bore import extract_dia_text_points
    w = entities_to_world([
        {"t": "T", "l": "TXT", "p": [0, 0], "v": "100A"},
        {"t": "T", "l": "TXT", "p": [1, 1], "v": "옥내소화전"},   # 노이즈
    ])
    got = extract_dia_text_points(w.texts)
    assert got == [(0.0, 0.0, 100)]


def test_모르는_종류는_조용히_지나간다():
    """INSERT 표지 등은 그릴 것이 없다 — 죽지 않고 넘어가야 한다."""
    w = entities_to_world([
        {"t": "I", "l": "BLK", "p": [0, 0], "n": "HEAD"},
        {"t": "?", "l": "X"},
        {"t": "L", "l": "PIPE", "p": [0, 0, 1, 1]},
    ])
    assert len(w.segs) == 1


def test_망가진_entity를_건너뛴다():
    w = entities_to_world([
        {"t": "L", "l": "P", "p": [0, 0]},          # 짧다
        {"t": "C", "l": "P", "c": []},              # 중심 없음
        {"t": "T", "l": "P", "p": [0]},             # 좌표 하나
        {"t": "L", "l": "P", "p": [0, 0, 1, 1]},    # 멀쩡
    ])
    assert len(w.segs) == 1 and w.circles == [] and w.texts == []


def test_캔버스_payload로_이어진다():
    """평면도와 같은 렌더 코드를 타는지 — 이것이 어댑터의 존재 이유다."""
    w = entities_to_world([
        {"t": "L", "l": "PIPE", "p": [0, 0, 100, 0]},
        {"t": "L", "l": "PIPE", "p": [100, 0, 100, 200]},
        {"t": "C", "l": "SYM", "c": [50, 50], "r": 10},
    ])
    pay = _world_payload(w)
    assert pay["counts"]["segs"] == 2 and pay["counts"]["circles"] == 1
    assert pay["bounds"]["maxy"] == 200.0
    assert pay["bounds"]["minx"] == 0.0
    ids = {b["layer"] for b in pay["bundles"]}
    assert ids == {"PIPE", "SYM"}


def test_빈_도면도_경계가_생긴다():
    pay = _world_payload(entities_to_world([]))
    assert pay["bounds"]["minx"] == 0.0 and pay["bounds"]["maxx"] == 1.0


# ─────────────────────────────────────────────── 요약
def test_요약이_길이를_합친다():
    got = riser_summary({
        "nodes": [{"label": "1"}, {"label": "10"}],
        "pipes": [{"label": "p1", "length": 1.5},
                  {"label": "p2", "length_m": 2.25}],
        "av_node_label": "10",
    })
    assert got["nodes"] == 2 and got["pipes"] == 2
    assert got["total_m"] == 3.75
    assert got["av_node_label"] == "10"


def test_요약은_빈것도_받는다():
    got = riser_summary(None)
    assert got["nodes"] == 0 and got["pipes"] == 0 and got["total_m"] == 0.0


def test_요약이_숫자아닌_길이를_건너뛴다():
    got = riser_summary({"nodes": [], "pipes": [{"length": "없음"},
                                                {"length": 2.0}]})
    assert got["total_m"] == 2.0


# ── 도면 색 (2026-09-08 · 사용자) ──────────────────────────────────
#
# 「캐드 도면의 색깔별로 보이게 해서(기계실도 같은 논리로) 구분이 쉽게」.
# 종전에는 모든 도형에 색 7 을 박아 계통도·기계실이 통째로 한 색이었다 —
# 배관·기호·건축선이 눈으로 안 갈렸다. 평면도는 처음부터 도면 색으로
# 그려 왔으니, 두 화면이 서로 다른 규칙을 쓰고 있었던 셈이다.


def test_레이어_색을_읽는다():
    from routes.module_f.subdrawing import layer_colors

    got = layer_colors({"layers": [
        {"name": "PIPE", "color": 1},
        {"name": "SYM", "color": 3},
    ]})
    assert got == {"PIPE": 1, "SYM": 3}


def test_꺼진_레이어의_음수_색을_편다():
    """★CAD 는 꺼진 레이어 색을 음수로 준다.

    계통도는 꺼둔 레이어에 배관이 있는 일이 흔해 우리는 그것도 읽는다
    (`include_hidden_layers`). 음수를 그대로 넘기면 색표에서 못 찾아
    **전부 같은 색**이 된다 — 고치려던 바로 그 증상으로 되돌아간다.
    """
    from routes.module_f.subdrawing import layer_colors

    got = layer_colors({"layers": [{"name": "OFF", "color": -5}]})
    assert got == {"OFF": 5}


def test_색을_주면_도면_색으로_그린다():
    from routes.module_f.subdrawing import entities_to_world

    ents = [{"t": "L", "l": "PIPE", "p": [0, 0, 1, 1]},
            {"t": "C", "l": "SYM", "c": [0, 0], "r": 1},
            {"t": "T", "l": "TEX", "p": [0, 0], "v": "x"}]
    w = entities_to_world(ents, {"PIPE": 1, "SYM": 3})
    assert w.segs[0][1] == 1
    assert w.circles[0][1] == 3
    assert w.texts[0][1] == 7, "모르는 레이어는 종전 색으로"


def test_색을_안_주면_종전_그대로():
    """옛 호출부(진단 스크립트·통합 시험)가 그대로 돌아야 한다."""
    from routes.module_f.subdrawing import entities_to_world

    w = entities_to_world([{"t": "L", "l": "PIPE", "p": [0, 0, 1, 1]}])
    assert w.segs[0][1] == 7


def test_모르는_ACI_번호도_제_색으로_풀린다():
    """★손으로 적은 색표는 흔한 번호 열댓 개뿐이었다.

    실도면에는 51·105·255 같은 번호가 흔하고, 그것들이 전부 같은 대체색으로
    떨어지면 «색으로 구분» 이 성립하지 않는다(실측: 계통도 11색 중 3개,
    기계실 13색 중 6개가 한 색으로 뭉쳤다).

    근사식을 새로 쓰지 않고 ezdxf 의 256색 표준표를 부른다 — 이 저장소가
    DXF 를 읽는 데 쓰는 바로 그 라이브러리다.
    """
    import sys as _sys
    from pathlib import Path as _P
    _g = str(_P(__file__).resolve().parent.parent / "cad_project_editor_g")
    if _g not in _sys.path:
        _sys.path.insert(0, _g)
    from services.cad_import.colors import rgb_dark

    got = {c: rgb_dark(c) for c in (51, 105, 11, 55, 115, 135)}
    assert len(set(got.values())) == len(got), got
    assert all(v.startswith("#") and len(v) == 7 for v in got.values()), got
    # 손으로 고른 다크용 색이 우선이다 — 검정(7)은 어두운 캔버스에서 흰색.
    assert rgb_dark(7) == "#ffffff"
