# -*- coding: utf-8 -*-
"""지시서 §3.6 — C180 실명 텍스트 귀속.

실명은 용도로 이어지고 용도는 NFTC 기준개수를 바꾼다. 그래서 여기 테스트가
지키는 것은 "이름을 많이 붙였나" 가 아니라 **엉뚱한 이름을 안 붙이는가**,
못 붙였을 때 조용히 넘어가지 않고 `needs_input` 으로 드러내는가다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.recognize import params as P  # noqa: E402
from core.design.recognize import room_faces as R  # noqa: E402
from core.design.recognize import room_label as L  # noqa: E402
from core.design.recognize import spatial as S  # noqa: E402


def _face(x0, y0, x1, y1):
    polygon = [(float(x0), float(y0)), (float(x1), float(y0)),
               (float(x1), float(y1)), (float(x0), float(y1))]
    return R.RoomFace(
        polygon=polygon, node_cycle=[0, 1, 2, 3],
        area_m2=abs((x1 - x0) * (y1 - y0)) / 1.0e6,
        perimeter_mm=2 * (abs(x1 - x0) + abs(y1 - y0)),
        virtual_ratio=0.0, unpaired_ratio=0.0, confidence=P.FACE_CONF_BASE,
        center=S.centroid(polygon))


def _text(x, y, value):
    return {"t": "T", "l": "A-ROOM", "p": [x, y], "v": value}


def _line(x1, y1, x2, y2):
    return {"t": "L", "l": "A-ROOM", "p": [x1, y1, x2, y2]}


def _one(faces, texts, lines=()):
    return L.assign_labels(faces, texts, lines).labels[0]


# ── 직접 귀속 ───────────────────────────────────────────────────────────

def test_폴리곤_안의_텍스트는_그_실의_이름이다():
    label = _one([_face(0, 0, 5000, 5000)], [_text(2500, 2500, "사무실")])
    assert label.name == "사무실"
    assert label.source == L.INSIDE
    assert label.confidence == P.CONF_LABEL_INSIDE
    assert label.needs_input is False


def test_폴리곤_밖의_텍스트는_그_실의_이름이_아니다():
    label = _one([_face(0, 0, 5000, 5000)], [_text(9000, 9000, "사무실")])
    assert label.name is None
    assert label.needs_input is True


def test_실_안의_실이_있으면_작은_쪽이_가져간다():
    faces = [_face(0, 0, 10000, 10000), _face(1000, 1000, 3000, 3000)]
    labels = L.assign_labels(faces, [_text(2000, 2000, "창고")]).labels
    assert labels[1].name == "창고"
    assert labels[0].name is None


# ── 면적 표기 제외 ──────────────────────────────────────────────────────

def test_면적_표기는_실명이_아니다():
    """§3.6 3항. 이걸 안 빼면 방 이름이 '32.50㎡' 가 된다."""
    label = _one([_face(0, 0, 5000, 5000)],
                 [_text(2500, 2500, "32.50㎡"), _text(2500, 2000, "회의실")])
    assert label.name == "회의실"


def test_숫자만_있는_텍스트도_실명이_아니다():
    """치수·실번호가 폴리곤 안에 들어온다. 실 이름이 '3,600' 인 방은 없다."""
    label = _one([_face(0, 0, 5000, 5000)], [_text(2500, 2500, "3,600")])
    assert label.name is None
    assert label.candidates == ["3,600"]


def test_면적_표기뿐이면_이름을_지어내지_않는다():
    label = _one([_face(0, 0, 5000, 5000)], [_text(2500, 2500, "32.50 m2")])
    assert label.name is None
    assert label.needs_input is True


def test_여럿이면_가장_긴_것을_고른다():
    label = _one([_face(0, 0, 5000, 5000)],
                 [_text(2500, 2500, "A"), _text(2500, 2000, "제1회의실")])
    assert label.name == "제1회의실"


# ── 지시선 추적 ─────────────────────────────────────────────────────────

def test_지시선_끝이_폴리곤_안이면_그_실의_이름이다():
    label = _one([_face(0, 0, 5000, 5000)], [_text(8000, 2500, "주차장")],
                 [_line(7900, 2500, 2500, 2500)])
    assert label.name == "주차장"
    assert label.source == L.LEADER
    assert label.confidence == P.CONF_LABEL_LEADER


def test_텍스트에_안_붙은_선은_지시선이_아니다():
    """500mm 밖에서 시작하는 선까지 받으면 옆 실 이름을 끌어온다."""
    label = _one([_face(0, 0, 5000, 5000)], [_text(8000, 2500, "주차장")],
                 [_line(7000, 2500, 2500, 2500)])
    assert label.name is None


def test_지시선이_어느_실에도_안_닿으면_귀속하지_않는다():
    out = L.assign_labels([_face(0, 0, 5000, 5000)], [_text(8000, 2500, "주차장")],
                          [_line(7900, 2500, 9000, 9000)])
    assert out.labels[0].name is None
    assert out.unassigned_texts == 1


def test_직접_귀속이_지시선보다_앞선다():
    faces = [_face(0, 0, 5000, 5000)]
    texts = [_text(2500, 2500, "사무"), _text(8000, 2500, "지시선으로_들어온_긴_이름")]
    label = _one(faces, texts, [_line(7900, 2500, 2500, 2500)])
    assert label.name == "사무"
    assert label.source == L.INSIDE


# ── 용도 추정 ───────────────────────────────────────────────────────────

def test_실명에서_용도를_추정한다():
    label = _one([_face(0, 0, 5000, 5000)], [_text(2500, 2500, "지하주차장")])
    assert label.use_hint == "주차장"
    assert label.use_confidence == P.CONF_USE_HINT


def test_힌트에_없는_실명은_용도를_비워_둔다():
    label = _one([_face(0, 0, 5000, 5000)], [_text(2500, 2500, "전실")])
    assert label.use_hint is None
    assert label.use_confidence == 0.0


def test_용도는_신뢰도와_무관하게_사람이_확정한다():
    """§3.6 — 신뢰도 0.95 이상이어도 자동 확정 금지."""
    dumped = _one([_face(0, 0, 5000, 5000)], [_text(2500, 2500, "사무실")]).to_dict()
    assert dumped["needs_confirm"] is True


# ── 운영 ────────────────────────────────────────────────────────────────

def test_텍스트가_없으면_전부_needs_input_이다():
    out = L.assign_labels([_face(0, 0, 5000, 5000), _face(9000, 0, 12000, 3000)], [])
    assert [lb.needs_input for lb in out.labels] == [True, True]


def test_실이_없어도_터지지_않는다():
    out = L.assign_labels([], [_text(0, 0, "사무실")])
    assert out.labels == []
    assert out.unassigned_texts == 1


def test_미터_단위_도면도_mm_로_환산된다():
    out = L.assign_labels([_face(0, 0, 5000, 5000)], [{"t": "T", "p": [2.5, 2.5],
                                                      "v": "사무실"}],
                          unit_to_mm=1000.0)
    assert out.labels[0].name == "사무실"


def test_결과는_직렬화된다():
    dumped = L.assign_labels([_face(0, 0, 5000, 5000)],
                             [_text(2500, 2500, "사무실")]).to_dict()
    assert dumped["labels"][0]["use_hint"] == "업무시설"
    assert dumped["labels"][0]["provenance"]


def test_임계값이_코드에_박혀_있지_않다():
    src = (_ROOT / "core" / "design" / "recognize" / "room_label.py").read_text(encoding="utf-8")
    for name in ("LEADER_SEARCH_RADIUS_MM", "LEADER_TEXT_ATTACH_MM",
                 "CONF_LABEL_INSIDE", "CONF_LABEL_LEADER", "CONF_USE_HINT"):
        assert f"P.{name}" in src, f"{name} 을 params 에서 읽지 않는다"


def test_USE_HINTS_는_지시서_표_그대로다():
    assert set(L.USE_HINTS) == {"업무시설", "공동주택", "판매시설", "숙박시설",
                                "의료시설", "노유자시설", "주차장", "창고시설"}
    assert L.use_of("매장") == "판매시설"
    assert L.use_of("PARKING LOT") == "주차장"
