# -*- coding: utf-8 -*-
"""지시서 §13.5 — 차원 무결성 검사.

여기서 지키는 것은 "검사가 초록이면 잰 것" 이라는 규약이다. 재지 못한 검사가
`pass` 로 나오면 검사표 전체가 장식이 된다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.checks import dimensional as D  # noqa: E402
from core.design.deterministic.constraints import build_constraints  # noqa: E402
from core.design.schema import Room  # noqa: E402

_M = 1000.0


def _constraints():
    return build_constraints({
        "building": {"floors_total": 8, "structure": "내화구조", "use": "업무시설"},
        "rooms": [{"use": "사무실", "ambient_temp_max_c": 30.0,
                   "ceiling": {"has_finish": True, "finish_height_mm": 2700,
                               "slab_height_mm": 3200}}],
    })


def _rect(rid: str, floor: str, w_m: float, h_m: float, **kwargs) -> Room:
    return Room(id=rid, floor=floor, area_m2=w_m * h_m, polygon=[
        (0.0, 0.0), (w_m * _M, 0.0), (w_m * _M, h_m * _M), (0.0, h_m * _M)],
        **kwargs)


class _Layout:
    def __init__(self, room_id: str, count: int):
        self.room_id = room_id
        self.heads = [object()] * count


@pytest.fixture(scope="module")
def c():
    return _constraints()


# ────────────────────────────────────────────────────────────────────────────
# 기준층 교차검증 — 최우선
# ────────────────────────────────────────────────────────────────────────────

def _두층(counts: tuple[int, int]):
    rooms = [_rect("R-1", "1F", 10.0, 10.0), _rect("R-2", "2F", 10.0, 10.0)]
    return rooms, {"R-1": counts[0], "R-2": counts[1]}


def test_같은_평면인데_헤드_수가_다르면_플래그():
    rooms, counts = _두층((12, 13))
    (rec,) = D.check_typical_floor_consistency(rooms, counts)
    assert rec["code"] == "TYPICAL_FLOOR_MISMATCH"
    assert rec["status"] == "flag"
    assert rec["floors"] == ["1F", "2F"]
    assert rec["counts"] == [12, 13]


def test_같은_평면에_헤드_수가_같으면_통과():
    rooms, counts = _두층((12, 12))
    (rec,) = D.check_typical_floor_consistency(rooms, counts)
    assert rec["status"] == "pass"


def test_평면이_다르면_통과가_아니라_미검증():
    """0.95 아래로 갈리면 비교한 것이 아니다. 통과로 적으면 거짓 초록이다."""
    rooms = [_rect("R-1", "1F", 10.0, 10.0), _rect("R-2", "2F", 4.0, 4.0)]
    (rec,) = D.check_typical_floor_consistency(rooms, {"R-1": 12, "R-2": 2})
    assert rec["status"] == "unverified"


def test_층이_하나면_교차검증할_것이_없다():
    rooms = [_rect("R-1", "1F", 10.0, 10.0)]
    (rec,) = D.check_typical_floor_consistency(rooms, {"R-1": 12})
    assert rec["status"] == "unverified"


def test_평면_묶기는_거의_같은_층을_같이_본다():
    """5cm 어긋난 층까지 다른 평면으로 갈리면 실 도면에서 아무 층도 안 묶인다."""
    rooms = [_rect("R-1", "1F", 10.0, 10.0), _rect("R-2", "2F", 10.0, 10.0)]
    rooms[1].polygon = [(x + 50.0, y) for x, y in rooms[1].polygon]
    assert D.group_floors_by_similarity(rooms) == [["1F", "2F"]]


def test_격자가_커져도_층마다_다른_격자를_쓰지_않는다():
    """예산 초과로 셀을 키울 때 층별로 따로 키우면 IoU 가 뜻을 잃는다."""
    rooms = [_rect("R-1", "1F", 400.0, 400.0), _rect("R-2", "2F", 400.0, 400.0)]
    assert D.group_floors_by_similarity(rooms) == [["1F", "2F"]]


# ────────────────────────────────────────────────────────────────────────────
# 면적·개수
# ────────────────────────────────────────────────────────────────────────────

def test_연면적이_없으면_면적_대조는_미검증():
    rooms = [_rect("R-1", "1F", 10.0, 10.0)]
    rec = D.check_room_area(rooms, None)
    assert rec["status"] == "unverified"
    assert rec["room_area_m2"] == pytest.approx(100.0)


def test_연면적과_5퍼센트_넘게_벌어지면_플래그():
    rooms = [_rect("R-1", "1F", 10.0, 10.0)]
    assert D.check_room_area(rooms, 120.0)["status"] == "flag"
    assert D.check_room_area(rooms, 103.0)["status"] == "pass"


def test_면적은_폴리곤에서도_구한다():
    """`area_m2` 가 비어 있어도 폴리곤이 있으면 잰다. 0 으로 접으면 검사가 꺼진다."""
    room = _rect("R-1", "1F", 10.0, 10.0)
    room.area_m2 = 0.0
    assert D.check_room_area([room], 100.0)["status"] == "pass"


def test_표본이_적으면_헤드수_검사는_미검증(c):
    """10m 각 실 하나면 이론값이 9.5개다. 하나만 달라도 10% 를 넘는다."""
    rooms = [_rect("R-1", "1F", 10.0, 10.0)]
    rec = D.check_head_count(rooms, {"R-1": 12}, c)
    assert rec["status"] == "unverified"


def test_헤드수가_이론값_10퍼센트_안이면_통과(c):
    rooms = [_rect("R-1", "1F", 40.0, 40.0)]
    theoretical = 1600.0 / c.head_spacing_square_m ** 2
    assert D.check_head_count(rooms, {"R-1": round(theoretical)}, c)["status"] == "pass"
    assert D.check_head_count(rooms, {"R-1": round(theoretical * 1.3)},
                              c)["status"] == "flag"


def test_설치제외_실은_분자와_분모에서_함께_빠진다(c):
    """분모에만 남기면 제외가 많은 건물이 무조건 미달로 읽힌다."""
    rooms = [_rect("R-1", "1F", 40.0, 40.0),
             _rect("R-x", "1F", 40.0, 40.0, head_exempt=True)]
    theoretical = 1600.0 / c.head_spacing_square_m ** 2
    rec = D.check_head_count(rooms, {"R-1": round(theoretical), "R-x": 0}, c)
    assert rec["status"] == "pass"
    assert rec["theoretical"] == pytest.approx(theoretical, rel=1e-3)


def test_설치제외가_30퍼센트를_넘으면_플래그():
    rooms = [_rect("R-1", "1F", 10.0, 10.0),
             _rect("R-x", "1F", 10.0, 10.0, head_exempt=True)]
    rec = D.check_exempt_area(rooms, None)
    assert rec["status"] == "flag"
    assert rec["basis"] == "room_area_sum"
    assert rec["ratio"] == pytest.approx(0.5)


def test_연면적이_있으면_그것을_분모로_쓴다():
    rooms = [_rect("R-x", "1F", 10.0, 10.0, head_exempt=True)]
    rec = D.check_exempt_area(rooms, 1000.0)
    assert rec["status"] == "pass"
    assert rec["basis"] == "gross_floor_area"


# ────────────────────────────────────────────────────────────────────────────

def test_헤드수는_객체든_dict든_같게_센다():
    assert D.head_counts_of([_Layout("R-1", 3)]) == {"R-1": 3}
    assert D.head_counts_of([{"room_id": "R-1", "heads": [1, 2, 3]}]) == {"R-1": 3}
    assert D.head_counts_of([{"room_id": "R-1"}]) == {"R-1": 0}


def test_검사표는_전부_돌고_flags_는_부분집합이다(c):
    rooms = [_rect("R-1", "1F", 10.0, 10.0), _rect("R-2", "2F", 10.0, 10.0)]
    out = D.run_checks(rooms, [_Layout("R-1", 12), _Layout("R-2", 13)], c)
    codes = [r["code"] for r in out["checks"]]
    assert codes == ["TYPICAL_FLOOR_MISMATCH", "ROOM_AREA_MISMATCH",
                     "HEAD_COUNT_DEVIATION", "EXEMPT_AREA_EXCESS",
                     "PIPE_LENGTH_PER_HEAD"]
    assert "TYPICAL_FLOOR_MISMATCH" in [r["code"] for r in out["flags"]]
    assert all(r in out["checks"] for r in out["flags"])
    assert all(r["status"] in ("pass", "flag", "unverified") for r in out["checks"])


def test_배관_총연장은_지어내지_않는다(c):
    """C5 전이고 회귀 범위도 없다. 통과로 적으면 없는 근거가 생긴다."""
    out = D.run_checks([_rect("R-1", "1F", 10.0, 10.0)], [_Layout("R-1", 12)], c)
    (rec,) = [r for r in out["checks"] if r["code"] == "PIPE_LENGTH_PER_HEAD"]
    assert rec["status"] == "unverified"
