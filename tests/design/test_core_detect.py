# -*- coding: utf-8 -*-
"""지시서 §3.7 — C190 코어 판별.

§3.7 의 첫 문장이 **단층만으로는 확정하지 마라** 다. 그래서 여기 테스트가 지키는
것은 "코어를 찾았나" 가 아니라 **반복 증거 없이 확정하지 않는가**다. 샤프트를
실로 잘못 보면 수직 관통부에 헤드가 달리고, 그 반대면 실 하나가 통째로 빠진다.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.recognize import core_detect as C  # noqa: E402
from core.design.recognize import params as P  # noqa: E402
from core.design.recognize import room_faces as R  # noqa: E402
from core.design.recognize import spatial as S  # noqa: E402


def _face(cx, cy, side):
    half = side / 2.0
    polygon = [(cx - half, cy - half), (cx + half, cy - half),
               (cx + half, cy + half), (cx - half, cy + half)]
    return R.RoomFace(
        polygon=polygon, node_cycle=[0, 1, 2, 3], area_m2=side * side / 1.0e6,
        perimeter_mm=4 * side, virtual_ratio=0.0, unpaired_ratio=0.0,
        confidence=P.FACE_CONF_BASE, center=S.centroid(polygon))


_SHAFT = 1500.0        # 2.25㎡ — §3.2 소형 폐합 범위 안
_ROOM = 5000.0         # 25㎡ — 소형이 아니다


# ── 다층 반복 ───────────────────────────────────────────────────────────

def test_여러_층에_같은_자리_같은_크기면_샤프트다():
    floors = [[_face(0, 0, _SHAFT)] for _ in range(3)]
    out = C.detect_cores(floors)
    assert len(out.candidates) == 3
    assert out.candidates[0].kind == C.SHAFT
    assert out.candidates[0].confidence == P.CONF_SHAFT_MULTIFLOOR


def test_한_층에만_있으면_샤프트가_아니다():
    floors = [[_face(0, 0, _SHAFT)], [_face(50000, 0, _SHAFT)],
              [_face(90000, 0, _SHAFT)]]
    assert C.detect_cores(floors).candidates == []


def test_중심이_멀면_같은_수직_관통부가_아니다():
    off = P.CORE_CENTER_DIST_MAX_MM * 2
    floors = [[_face(0, 0, _SHAFT)], [_face(off, 0, _SHAFT)]]
    assert C.detect_cores(floors).candidates == []


def test_면적비가_대역_밖이면_같은_관통부가_아니다():
    """둘 다 소형 범위 안이지만 면적비 1.5 는 §3.7 대역(0.7~1.4) 밖이다."""
    bigger = _SHAFT * math.sqrt(1.5)
    floors = [[_face(0, 0, _SHAFT)], [_face(0, 0, bigger)]]
    assert floors[1][0].area_m2 < P.SMALL_CLOSED_AREA_MAX_M2
    assert C.detect_cores(floors).candidates == []


def test_큰_실은_코어_후보가_아니다():
    floors = [[_face(0, 0, _ROOM)] for _ in range(3)]
    assert C.detect_cores(floors).candidates == []


def test_한_층에서_여러_개가_맞아도_그_층은_한_표다():
    """§3.7 의 문턱이 층수로 쓰여 있다. 개수를 세면 2층 도면이 문턱을 넘는다."""
    d = P.CORE_CENTER_DIST_MAX_MM * 0.5
    floors = [[_face(0, 0, _SHAFT)],
              [_face(-d, 0, _SHAFT), _face(d, 0, _SHAFT)],
              [_face(80000, 0, _SHAFT)]]
    assert [c.floor for c in C.detect_cores(floors).candidates] == []


def test_증거가_결과에_실린다():
    out = C.detect_cores([[_face(0, 0, _SHAFT)] for _ in range(3)])
    assert any("층에서 같은 자리" in text for text in out.candidates[0].evidence)


# ── 단층 · 이름 힌트 ────────────────────────────────────────────────────

def test_단층이면_이름_힌트로만_후보를_낸다():
    out = C.detect_cores([[_face(0, 0, _SHAFT)]], names=[["PS"]])
    assert len(out.candidates) == 1
    assert out.candidates[0].kind == C.SHAFT
    assert out.candidates[0].confidence == P.CONF_SHAFT_SINGLE
    assert any("층 도면이 하나뿐" in line for line in out.provenance)


def test_단층에_이름이_없으면_후보를_내지_않는다():
    assert C.detect_cores([[_face(0, 0, _SHAFT)]]).candidates == []


def test_계단과_승강기도_이름으로_구분한다():
    out = C.detect_cores([[_face(0, 0, _SHAFT), _face(9000, 0, _SHAFT)]],
                         names=[["계단실", "E/V"]])
    assert [c.kind for c in out.candidates] == [C.STAIR, C.ELEVATOR]


def test_반복으로_확정한_것에_이름_후보를_또_붙이지_않는다():
    floors = [[_face(0, 0, _SHAFT)] for _ in range(3)]
    out = C.detect_cores(floors, names=[["PS"]] * 3)
    assert len(out.candidates) == 3
    assert all(c.confidence == P.CONF_SHAFT_MULTIFLOOR for c in out.candidates)


def test_반복_증거가_없는_실은_이름으로_후보가_된다():
    floors = [[_face(0, 0, _SHAFT), _face(60000, 0, _SHAFT)],
              [_face(0, 0, _SHAFT)]]
    out = C.detect_cores(floors, names=[[None, "EPS"], [None]])
    weak = [c for c in out.candidates if c.confidence == P.CONF_SHAFT_SINGLE]
    assert [c.face_index for c in weak] == [1]


# ── 운영 ────────────────────────────────────────────────────────────────

def test_확정은_언제나_사람이_한다():
    """§3.7 — 다층 반복으로 0.85 를 줘도 GATE 확정을 건너뛰지 않는다."""
    out = C.detect_cores([[_face(0, 0, _SHAFT)] for _ in range(3)])
    assert out.to_dict()["candidates"][0]["needs_confirm"] is True


def test_층이_비어도_터지지_않는다():
    out = C.detect_cores([[], []])
    assert out.candidates == []
    assert out.floor_count == 2


def test_임계값이_코드에_박혀_있지_않다():
    src = (_ROOT / "core" / "design" / "recognize" / "core_detect.py").read_text(encoding="utf-8")
    for name in ("CORE_CENTER_DIST_MAX_MM", "CORE_AREA_RATIO_MIN",
                 "CORE_AREA_RATIO_MAX", "CORE_FLOOR_SHARE_MIN",
                 "CONF_SHAFT_MULTIFLOOR", "CONF_SHAFT_SINGLE"):
        assert f"P.{name}" in src, f"{name} 을 params 에서 읽지 않는다"
