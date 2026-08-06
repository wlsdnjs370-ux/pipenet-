# -*- coding: utf-8 -*-
"""지시서 §3.0 — C1 인식 사슬.

여기서 지키는 것은 인식률이 아니다(그건 부록 B 가 말한 대로 도면 벤치마크의
몫이다). **인식기가 채운 값이 확정으로 둔갑하지 않는가**, 그리고 **사슬이 멈출
때 조용히 빈 결과를 돌려주지 않는가** 둘이다. 전자를 놓치면 사람이 볼 기회가
사라지고, 후자를 놓치면 화면이 "인식했는데 방이 없다" 로 읽힌다.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design import gate as GA  # noqa: E402
from core.design.recognize import pipeline as PL  # noqa: E402


def _room_dxf(x=0.0, y=0.0, width=6000.0, height=5000.0, thickness=200.0,
              layer="A-WALL"):
    """두 겹 선으로 그린 방 하나. C150 이 짝지을 수 있는 최소 도면이다."""
    t = thickness
    out = []
    for x0, y0, x1, y1 in (
            (0, 0, width, 0), (0, height, width, height),
            (0, 0, 0, height), (width, 0, width, height),
            (t, t, width - t, t), (t, height - t, width - t, height - t),
            (t, t, t, height - t), (width - t, t, width - t, height - t)):
        out.append({"t": "L", "l": layer,
                    "p": [x + x0, y + y0, x + x1, y + y1]})
    return out


def _clutter_dxf(layer="A-FURN", n=12):
    """평행쌍은 있지만 벽으로 보기엔 너무 짧은 선들 — WALL 이 안 나오는 도면."""
    out = []
    for i in range(n):
        y = i * 900.0
        out.append({"t": "L", "l": layer, "p": [0.0, y, 400.0, y]})
        out.append({"t": "L", "l": layer, "p": [0.0, y + 100.0, 400.0, y + 100.0]})
    return out


# 실 폴리곤은 bbox 의 절반을 넘으면 외곽 오검출로 버려진다(§3.5). 실제 도면처럼
# 방보다 넉넉한 범위를 준다 — 방이 bbox 를 꽉 채우면 그 방이 사라진다.
_BBOX = {"minx": -5000.0, "miny": -5000.0, "maxx": 25000.0, "maxy": 25000.0}


def _run(entities, bbox=_BBOX, **kw):
    return list(PL.recognize(entities, bbox, **kw))


def _last(messages) -> dict:
    return messages[-1]


# ── bbox 표기 흡수 ──────────────────────────────────────────────────────

def test_두_가지_bbox_표기를_모두_읽는다():
    """inspect 는 `x_min`, C130 은 `minx` 를 쓴다. 한쪽만 읽으면 단위가 None 이 된다."""
    a = PL.normalize_bbox({"minx": 0, "miny": 0, "maxx": 10, "maxy": 20})
    b = PL.normalize_bbox({"x_min": 0, "y_min": 0, "x_max": 10, "y_max": 20})
    assert a == b == {"minx": 0.0, "miny": 0.0, "maxx": 10.0, "maxy": 20.0}


def test_한_축이_비면_bbox_를_쓰지_않는다():
    assert PL.normalize_bbox({"minx": 0, "miny": 0, "maxx": 10}) is None
    assert PL.normalize_bbox(None) is None


# ── 메시지 순서 (§11.1) ─────────────────────────────────────────────────

def test_메시지가_11_1_순서로_나온다():
    types = [m["type"] for m in _run(_room_dxf(), wall_layers=["A-WALL"])]
    assert types == ["fingerprint", "phase", "centerlines", "virtual_edges",
                     "rooms", "cores", "building"]


def test_방_하나짜리_도면에서_실_하나를_뜬다():
    msgs = _run(_room_dxf(), wall_layers=["A-WALL"])
    rooms = next(m for m in msgs if m["type"] == "rooms")["rooms"]
    assert len(rooms) == 1
    assert 25.0 < rooms[0]["area_m2"] < 30.0


# ── WALL 이 없을 때 (§3.2 판정표의 한계) ────────────────────────────────

def test_WALL_이_없으면_빈_결과가_아니라_막혔다고_말한다():
    """조용히 실 0개를 돌려주면 화면은 '인식했는데 방이 없다' 로 읽는다."""
    msgs = _run(_clutter_dxf())
    phase = next(m for m in msgs if m["type"] == "phase")
    assert phase["blocked"] == "no_wall_layer"
    assert _last(msgs)["blocked"] == "no_wall_layer"
    assert [m["type"] for m in msgs] == ["fingerprint", "phase", "building"]


def test_막혔을_때_사람이_고를_차선을_함께_준다():
    msgs = _run(_clutter_dxf())
    phase = next(m for m in msgs if m["type"] == "phase")
    assert [c["name"] for c in phase["candidates"]] == ["A-FURN"]
    assert phase["candidates"][0]["offset_peaks_mm"]


def test_사람이_고른_레이어가_C140_판정을_이긴다():
    msgs = _run(_room_dxf(layer="A-FURN"), wall_layers=["A-FURN"])
    assert _last(msgs)["wall_source"] == PL.WALL_SOURCE_OPERATOR
    assert _last(msgs)["wall_layers"] == ["A-FURN"]


def test_없는_레이어를_넘기면_무시한다():
    """오타 하나로 사슬이 조용히 빈 벽을 돌리게 두지 않는다."""
    msgs = _run(_clutter_dxf(), wall_layers=["A-WALL"])
    assert _last(msgs)["blocked"] == "no_wall_layer"


# ── 제안은 확정이 아니다 (§3.6, §4) ─────────────────────────────────────

def test_인식기가_채운_실명은_결손으로_남는다():
    ents = _room_dxf() + [{"t": "T", "l": "A-ROOM", "p": [3000, 2500], "v": "사무실"}]
    msgs = _run(ents, wall_layers=["A-WALL"])
    draft = _draft(_last(msgs))
    room = draft.rooms[0]
    assert room.name == "사무실"
    assert room.provenance["name"] == PL.PROV_C180
    assert room.provenance["name"] not in (GA.PROV_GATE, GA.PROV_DEFAULT)
    assert room.confidence["name"] > 0.0


def test_용도_추정도_확정이_아니다():
    ents = _room_dxf() + [{"t": "T", "l": "A-ROOM", "p": [3000, 2500], "v": "사무실"}]
    draft = _draft(_last(_run(ents, wall_layers=["A-WALL"])))
    room = draft.rooms[0]
    assert room.use == "업무시설"
    assert room.provenance["use"] == PL.PROV_C180
    assert any(key.endswith(".use") for key in GA.unresolved(draft))


def test_코어는_확정_없이_나간다():
    ents = (_room_dxf() + _room_dxf(x=1000.0, y=1000.0, width=1500.0, height=1500.0)
            + [{"t": "T", "l": "A-ROOM", "p": [1750, 1750], "v": "PS"}])
    msgs = _run(ents, wall_layers=["A-WALL"])
    for core in _draft(_last(msgs)).cores:
        assert core.confirmed is None


# ── 산출물 ─────────────────────────────────────────────────────────────

def test_단계별_소요_시간이_리포트에_실린다():
    """부록 C.1 예산을 어디서 넘겼는지 알 수 없으면 예산이 아니다."""
    msg = _last(_run(_room_dxf(), wall_layers=["A-WALL"]))
    assert [s["name"] for s in msg["stages"]] == [
        "C130/C140", "C150", "C160", "C170", "C180", "C190"]
    assert msg["seconds"] >= 0.0


def test_버린_face_개수가_리포트에_남는다():
    msg = _last(_run(_room_dxf(), wall_layers=["A-WALL"]))
    assert set(msg["counts"]["dropped"]) >= {"outer", "too_small", "degenerate"}


def test_엔티티가_없어도_터지지_않는다():
    msgs = _run([], bbox=None)
    assert _last(msgs)["blocked"] == "no_wall_layer"
    assert _last(msgs)["draft"]["scale"]["confidence"] == 0.0


def _draft(building_msg: dict):
    from core.design.schema import BuildingDraft
    return BuildingDraft.from_dict(building_msg["draft"])
