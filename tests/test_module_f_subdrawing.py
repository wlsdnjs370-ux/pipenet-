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
