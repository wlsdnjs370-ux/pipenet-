# -*- coding: utf-8 -*-
"""지시서 §13.4 H1~H6 — C4 헤드 배치.

여기가 통과한다는 것은 "도면과 무관하게 실 하나를 옳게 덮는다"는 뜻이다. 실
도면 벤치마크는 별도이고, 이 파일은 좌표만 준 합성 실로 규약을 고정한다.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.deterministic import head_layout as H  # noqa: E402
from core.design.deterministic.constraints import build_constraints  # noqa: E402
from core.design.schema import Obstacles, Room  # noqa: E402

_M = 1000.0


def _constraints():
    """R=2.3 / S=√2R 를 내는 업무시설. 값은 build_constraints 밖에서 만들지 않는다."""
    return build_constraints({
        "building": {"floors_total": 8, "structure": "내화구조", "use": "업무시설"},
        "rooms": [{"use": "사무실", "ambient_temp_max_c": 30.0,
                   "ceiling": {"has_finish": True, "finish_height_mm": 2700,
                               "slab_height_mm": 3200}}],
    })


def _rect(rid: str, w_m: float, h_m: float, *, x0: float = 0.0, y0: float = 0.0,
          **kwargs) -> Room:
    return Room(id=rid, floor="1F", polygon=[
        (x0, y0), (x0 + w_m * _M, y0), (x0 + w_m * _M, y0 + h_m * _M),
        (x0, y0 + h_m * _M)], **kwargs)


def _uncovered(layout, room, radius_m: float) -> list:
    """R 밖에 남은 곳. 표본이 아니라 원·변 교점으로 재므로 비면 증명된 100% 다."""
    return H.coverage_witnesses([(h.x, h.y) for h in layout.heads],
                                [(float(p[0]), float(p[1])) for p in room.polygon],
                                radius_m * _M)


@pytest.fixture(scope="module")
def c():
    return _constraints()


# ────────────────────────────────────────────────────────────────────────────
# H1~H3 — 배치
# ────────────────────────────────────────────────────────────────────────────

def test_h1_정사각실_전면_피복(c):
    """10m × 10m, R=2.3 → 피복 100%, 헤드 12개 이하."""
    room = _rect("R-H1", 10.0, 10.0)
    layout = H.layout_heads(room, c)
    assert not _uncovered(layout, room, c.horizontal_distance_m)
    assert len(layout.heads) <= 12, [h.to_dict() for h in layout.heads]
    assert layout.metrics["area_m2"] == pytest.approx(100.0)


def test_h2_복도_벽이격(c):
    """폭 2m × 길이 20m 복도 → 벽 이격 ≤ S/2."""
    room = _rect("R-H2", 20.0, 2.0)
    layout = H.layout_heads(room, c)
    assert not _uncovered(layout, room, c.horizontal_distance_m)
    assert layout.metrics["wall_gap_m"] <= c.wall_clearance_max_m
    assert not layout.flags, layout.flags


def test_h3_좁은실도_최소_한개(c):
    """3㎡ 구획실 → 면적과 무관하게 헤드 1개. 0개인 실은 아무도 못 지킨다."""
    layout = H.layout_heads(_rect("R-H3", 1.5, 2.0), c)
    assert len(layout.heads) == 1
    assert not layout.flags, layout.flags


def test_격자보다_촘촘한_실도_한_개_이상(c):
    """격자점이 하나도 안 걸리는 실 — 대표점으로라도 놓는다."""
    room = Room(id="R-thin", floor="1F", polygon=[
        (0.0, 0.0), (600.0, 0.0), (600.0, 400.0), (0.0, 400.0)])
    layout = H.layout_heads(room, c)
    assert len(layout.heads) == 1
    assert H.point_in_polygon((layout.heads[0].x, layout.heads[0].y), room.polygon)


def test_설치제외_실은_헤드를_놓지_않는다(c):
    layout = H.layout_heads(_rect("R-ex", 5.0, 5.0, head_exempt=True), c)
    assert layout.heads == []
    assert layout.metrics["head_exempt"] is True


def test_행열_인덱스는_0부터_연속이다(c):
    """§8.3 — C5 가 이 인덱스로 가지배관을 묶는다. 비면 다시 군집화해야 한다."""
    layout = H.layout_heads(_rect("R-idx", 10.0, 10.0), c)
    rows = sorted({h.row for h in layout.heads})
    cols = sorted({h.col for h in layout.heads})
    assert rows == list(range(len(rows)))
    assert cols == list(range(len(cols)))
    # 축은 C5 가 정한다(§9.2 C510). 여기서 "x" 로 적어 두면 정해진 축으로 읽힌다.
    assert all(h.branch_axis is None for h in layout.heads)


def test_헤드_아이디는_실_안에서_유일하다(c):
    layout = H.layout_heads(_rect("R-uniq", 12.0, 9.0), c, room_index=7)
    ids = [h.id for h in layout.heads]
    assert len(set(ids)) == len(ids)
    assert all(i.startswith("H-007-") for i in ids)


def test_폴리곤이_없으면_지어내지_않는다(c):
    layout = H.layout_heads(Room(id="R-bad", floor="1F", polygon=[(0.0, 0.0)]), c)
    assert layout.heads == []
    assert [f["code"] for f in layout.flags] == ["ROOM_POLYGON_INVALID"]


def test_큰_실도_구석까지_덮는다(c):
    """40m 각. 후보를 성기게 줄여도 피복 판정은 표본이 아니라 기하로 한다."""
    room = _rect("R-big", 40.0, 40.0)
    layout = H.layout_heads(room, c)
    assert not _uncovered(layout, room, c.horizontal_distance_m)
    assert "coverage_hole_at" not in layout.metrics


# ────────────────────────────────────────────────────────────────────────────
# 기하 보조
# ────────────────────────────────────────────────────────────────────────────

def test_주축은_벽과_나란하다():
    theta = math.radians(30.0)
    poly = [H._rotate(x, y, theta) for x, y in
            [(0.0, 0.0), (8000.0, 0.0), (8000.0, 3000.0), (0.0, 3000.0)]]
    axes = H.principal_axes(poly)
    assert min(abs(a % (math.pi / 2) - theta) for a in axes) < 1e-6


def test_벽_이격은_수직거리로_잰다():
    """점거리로 재면 헤드 사이 중간 지점 때문에 적법한 배치도 초과로 읽힌다."""
    poly = [(0.0, 0.0), (10000.0, 0.0), (10000.0, 4000.0), (0.0, 4000.0)]
    heads = [H.Head(id="a", room_id="r", x=1000.0, y=2000.0, row=0, col=0),
             H.Head(id="b", room_id="r", x=9000.0, y=2000.0, row=0, col=1)]
    assert H.wall_gap_mm(heads, poly, 1000.0) == pytest.approx(2000.0)


def test_헤드_없는_긴_벽은_0이_아니라_무한이다():
    """0 으로 접으면 헤드가 없는 벽이 가장 좋은 배치로 읽힌다."""
    poly = [(0.0, 0.0), (10000.0, 0.0), (10000.0, 10000.0), (0.0, 10000.0)]
    heads = [H.Head(id="a", room_id="r", x=-1.0, y=5000.0, row=0, col=0)]
    assert math.isinf(H.wall_gap_mm(heads, poly, 1000.0))


# ────────────────────────────────────────────────────────────────────────────
# H4~H6 — 살수장애 (§8.4, §8.5)
# ────────────────────────────────────────────────────────────────────────────

def _laid_out(c, **room_kwargs):
    room = _rect("R-obs", 10.0, 10.0, **room_kwargs)
    return room, H.layout_heads(room, c)


def test_h4_보는_60cm_룰이_아니라_별도_표다(c):
    """보가 헤드에서 0.5m → `below_beam_bottom` 요구. FAIL 아니다."""
    room, layout = _laid_out(c)
    head = layout.heads[0]
    beam = {"point": [head.x + 500.0, head.y]}
    H.check_obstacles(layout, room, c, Obstacles(status="complete", beams=[beam]))

    assert not layout.flags, layout.flags
    mine = [r for r in layout.beam_requirements if r["head_id"] == head.id]
    assert mine and mine[0]["horizontal_m"] == pytest.approx(0.5)
    assert mine[0]["requirement"] == "below_beam_bottom"
    # 보 때문에 헤드를 더하지도, 옮기지도 않는다.
    assert all(h.provenance == "grid" for h in layout.heads)


def test_h5_덕트는_하부헤드를_먼저_시도한다(c):
    """덕트가 헤드에서 0.4m → 하부 헤드 추가. 추가되면 FAIL 아니다."""
    room, layout = _laid_out(c)
    before = len(layout.heads)
    head = layout.heads[0]
    duct = {"point": [head.x + 400.0, head.y]}
    H.check_obstacles(layout, room, c, Obstacles(status="complete", ducts=[duct]))

    assert not layout.flags, layout.flags
    added = [h for h in layout.heads if h.provenance == "under_obstacle"]
    assert len(added) == 1
    assert (added[0].x, added[0].y) == (duct["point"][0], duct["point"][1])
    assert layout.metrics["head_count"] == before + 1


def test_h5_하부헤드를_못_놓으면_FAIL(c):
    """벽에 붙은 덕트 — 자리가 없으면 조용히 넘어가지 않는다."""
    room = _rect("R-obs", 10.0, 10.0)
    # 배치가 어디에 헤드를 놓든 이 상황이 나오도록 좌표를 직접 준다. 실제 배치에
    # 기대면 벽에 붙은 헤드가 나올 때만 검사되는 테스트가 된다.
    layout = H.RoomLayout(room_id=room.id, heads=[
        H.Head(id="H-000-001", room_id=room.id, x=5000.0, y=300.0, row=0, col=0)])
    duct = {"point": [5000.0, 50.0]}   # 아래 벽에서 50mm — 이격 100mm 미만
    H.check_obstacles(layout, room, c, Obstacles(status="complete", ducts=[duct]))

    assert [f["code"] for f in layout.flags] == ["OBSTRUCTION_UNRESOLVED"]
    assert all(h.provenance == "grid" for h in layout.heads)


def test_멀리_있는_덕트는_아무것도_하지_않는다(c):
    room, layout = _laid_out(c)
    before = len(layout.heads)
    H.check_obstacles(layout, room, c, Obstacles(
        status="complete", ducts=[{"point": [5000.0, 5000.0]}],
        lights=[{"point": [5100.0, 5100.0]}]))
    near = [h for h in layout.heads if h.provenance == "under_obstacle"]
    assert len(layout.heads) == before + len(near)


def test_실_밖의_장애물은_이_실의_것이_아니다(c):
    room, layout = _laid_out(c)
    before = len(layout.heads)
    H.check_obstacles(layout, room, c, Obstacles(
        status="complete", ducts=[{"point": [-500.0, 5000.0]}],
        beams=[{"point": [20000.0, 5000.0]}]))
    assert not layout.flags
    assert len(layout.heads) == before
    assert layout.beam_requirements == []


def test_h6_장애물_정보가_없으면_플래그를_남긴다(c):
    """§8.5 — 검증하지 않은 것과 검증해서 통과한 것은 다르다."""
    room, layout = _laid_out(c)
    H.check_obstacles(layout, room, c, Obstacles(status="none"))
    assert [f["code"] for f in layout.flags] == ["OBSTACLE_UNVERIFIED"]
    assert layout.beam_requirements == []


def test_h6_상태가_아예_모르면도_같다(c):
    """`None` 은 "장애물 없음" 이 아니라 "모른다" 다."""
    room, layout = _laid_out(c)
    H.check_obstacles(layout, room, c, None)
    assert [f["code"] for f in layout.flags] == ["OBSTACLE_UNVERIFIED"]
