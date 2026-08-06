# -*- coding: utf-8 -*-
"""지시서 §3.4 — C160 개구부 간극 가상 폐합. 최대 위험 지점.

여기서 간극을 잘못 이으면 두 실이 하나가 되고 면적이 두 배가 되고 헤드 개수가
틀린다. 그래서 이 테스트가 지키는 것은 "많이 닫았는가" 가 아니라 **닫지 말아야 할
것을 닫지 않는가**, 그리고 닫은 것마다 근거가 붙어 있는가다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.recognize import opening_close as O  # noqa: E402
from core.design.recognize import params as P  # noqa: E402
from core.design.recognize import wall_centerline as W  # noqa: E402


def _result(*segments, wall_repr=W.WALL_REPR_DOUBLE, thickness=200.0):
    lines = [W.Centerline(p1=(s[0], s[1]), p2=(s[2], s[3]),
                          thickness_mm=thickness, source_pair=(i, i),
                          unpaired=False, confidence=P.CONF_CENTERLINE_PAIRED)
             for i, s in enumerate(segments)]
    nodes, degree = W.snap_endpoints(lines)
    return W.CenterlineResult(centerlines=lines, nodes=nodes,
                              node_degree=degree, wall_repr=wall_repr,
                              paired_ratio=1.0)


def _arc(cx, cy, r, start, end):
    return {"t": "A", "l": "A-DOOR", "c": [cx, cy], "r": r, "a": [start, end]}


def _only(closure):
    assert len(closure.virtual_edges) == 1
    return closure.virtual_edges[0]


# ── 증거 3종 ────────────────────────────────────────────────────────────

def test_문_호가_간극을_가로지르면_door_다():
    """경첩(중심)이 한쪽 문설주, 닫힌 위치 끝점이 반대쪽 문설주."""
    result = _result((0, 0, 4000, 0), (4900, 0, 9000, 0))
    closure = O.close_openings(result, [_arc(4000, 0, 900, 0, 90)])
    edge = _only(closure)
    assert edge.kind == O.DOOR
    assert edge.confidence == P.CONF_VE_DOOR
    assert edge.gap_mm == pytest.approx(900.0)
    assert len(edge.evidence) == 2


def test_반경이_간극_폭과_안_맞으면_문이_아니다():
    """0.8~1.3배 밖. 옆방 문 호가 우연히 걸리는 것을 막는다."""
    result = _result((0, 0, 4000, 0), (4900, 0, 9000, 0))
    closure = O.close_openings(result, [_arc(4000, 0, 400, 0, 90)])
    assert _only(closure).kind == O.OPENING


def test_같은_앵커_하나로는_문_근거가_되지_않는다():
    """짧은 간극에서는 앵커 하나가 양끝에 다 걸린다 — 그건 근거가 아니다."""
    result = _result((0, 0, 4000, 0), (4200, 0, 9000, 0))
    closure = O.close_openings(result, [_arc(4100, 0, 200, 0, 90)])
    assert all(e.kind != O.DOOR for e in closure.virtual_edges)


def test_공선이고_폭이_개구부_범위면_opening_이다():
    result = _result((0, 0, 4000, 0), (5000, 0, 9000, 0))
    edge = _only(O.close_openings(result))
    assert edge.kind == O.OPENING
    assert edge.confidence == P.CONF_VE_OPENING


def test_코너_미접합은_inferred_다():
    result = _result((0, 0, 4000, 0), (4500, 500, 4500, 5000))
    edge = _only(O.close_openings(result))
    assert edge.kind == O.INFERRED
    assert edge.confidence == P.CONF_VE_INFERRED


def test_나란하지만_한_직선_위가_아니면_공선이_아니다():
    """각도차만 보면 복도 양쪽 벽이 개구부가 된다."""
    result = _result((0, 0, 4000, 0), (5000, 1000, 9000, 1000))
    edges = O.close_openings(result).virtual_edges
    assert [e.kind for e in edges] == [O.INFERRED]


# ── 닫지 않는 경우 ──────────────────────────────────────────────────────

def test_공선인데_폭이_개구부_범위를_넘으면_잇지_않는다():
    """§3.4 표에 없는 조합이다. 지어내 채우면 없던 실이 생긴다."""
    result = _result((0, 0, 4000, 0), (6500, 0, 9000, 0))
    assert O.close_openings(result).virtual_edges == []


def test_스냅이_처리할_간극은_건너뛴다():
    result = _result((0, 0, 4000, 0), (4080, 0, 9000, 0))
    assert O.close_openings(result).virtual_edges == []


def test_이미_이어진_끝점은_후보가_아니다():
    """§3.4 — 스냅 군집 크기 ≥ 2 면 건드리지 않는다."""
    result = _result((0, 0, 3000, 0), (3000, 0, 3000, 3000),
                     (3800, 0, 6000, 0))
    corner = result.nodes.index((3000.0, 0.0))
    assert result.node_degree[corner] == 2

    closure = O.close_openings(result)
    assert closure.open_endpoints == 4
    assert corner not in [n for e in closure.virtual_edges for n in (e.n1, e.n2)]


def test_한_중심선의_두_끝을_서로_잇지_않는다():
    """벽 하나가 통째로 실 하나가 되는 것을 막는다."""
    result = _result((0, 0, 2500, 0))
    assert O.close_openings(result).virtual_edges == []


def test_탐색_반경_밖은_보지_않는다():
    result = _result((0, 0, 4000, 0), (12000, 0, 16000, 0))
    assert O.close_openings(result).virtual_edges == []


# ── 중복 방지 ───────────────────────────────────────────────────────────

def test_한_끝점은_가상_간선을_하나만_갖는다():
    """§3.4 중복 방지. 둘을 다 이으면 한 벽 끝에서 실이 두 갈래로 샌다."""
    result = _result((0, 0, 4000, 0), (5000, 0, 9000, 0),
                     (4000, -900, 4000, -5000))
    closure = O.close_openings(result)
    ends = [n for e in closure.virtual_edges for n in (e.n1, e.n2)]
    assert len(ends) == len(set(ends))


def test_근거가_강한_쪽을_먼저_고른다():
    """같은 끝점에 door 와 inferred 가 걸리면 door 가 이긴다."""
    result = _result((0, 0, 4000, 0), (4900, 0, 9000, 0),
                     (4000, -500, 4000, -5000))
    closure = O.close_openings(result, [_arc(4000, 0, 900, 0, 90)])
    assert closure.virtual_edges[0].kind == O.DOOR


# ── 절대 규칙 / 운영 ────────────────────────────────────────────────────

def test_모든_가상_간선은_is_virtual_로_나간다():
    """§3.4 절대 규칙 — 캔버스에서 점선 + 다른 색으로 그려야 한다."""
    result = _result((0, 0, 4000, 0), (5000, 0, 9000, 0))
    dumped = O.close_openings(result).to_dict()
    assert dumped["virtual_edges"][0]["is_virtual"] is True
    assert dumped["virtual_edges"][0]["evidence"]


def test_완화_모드는_공차만_넓히고_간극_폭_범위는_그대로다():
    """단선 표기 도면이라고 간극 범위를 넓히면 없던 실이 생긴다."""
    tilted = _result((0, 0, 4000, 0), (5000, 70, 9000, 70),
                     wall_repr=W.WALL_REPR_SINGLE)
    assert _only(O.close_openings(tilted)).kind == O.OPENING

    wide = _result((0, 0, 4000, 0), (6500, 0, 9000, 0),
                   wall_repr=W.WALL_REPR_SINGLE)
    assert O.close_openings(wide).virtual_edges == []


def test_엄격_모드에서는_같은_배치가_공선이_아니다():
    tilted = _result((0, 0, 4000, 0), (5000, 70, 9000, 70))
    assert _only(O.close_openings(tilted)).kind == O.INFERRED


def test_완화_모드는_provenance_에_남는다():
    result = _result((0, 0, 4000, 0), (5000, 0, 9000, 0),
                     wall_repr=W.WALL_REPR_SINGLE)
    assert any("완화" in line for line in O.close_openings(result).provenance)


def test_문_호도_단위가_환산된다():
    arcs = O.door_arcs([_arc(4, 0, 0.9, 0, 90)], unit_to_mm=1000.0)
    assert arcs[0].radius_mm == pytest.approx(900.0)
    assert arcs[0].center == pytest.approx((4000.0, 0.0))
    assert arcs[0].p1 == pytest.approx((4900.0, 0.0))


def test_중심선이_없어도_터지지_않는다():
    result = W.build_centerlines([], offset_peaks_mm=[150.0])
    closure = O.close_openings(result)
    assert closure.virtual_edges == []
    assert closure.open_endpoints == 0


def test_임계값이_코드에_박혀_있지_않다():
    src = (_ROOT / "core" / "design" / "recognize" / "opening_close.py").read_text(encoding="utf-8")
    for name in ("GAP_SEARCH_RADIUS_MM", "GAP_MIN_MM", "DOOR_GAP_ENDPOINT_TOL_MM",
                 "OPENING_GAP_MIN_MM", "INFERRED_GAP_MAX_MM", "GAP_CANDIDATE_MAX"):
        assert f"P.{name}" in src, f"{name} 을 params 에서 읽지 않는다"
