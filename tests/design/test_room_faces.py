# -*- coding: utf-8 -*-
"""지시서 §3.5 — C170 폐합 영역 → 실 폴리곤.

실 면적이 헤드 개수를 정한다. 그래서 여기 테스트가 지키는 것은 "face 를 몇 개
뽑았나" 가 아니라 **건물 바깥이 방이 되지 않는가**, 두 실이 하나로 합쳐지지
않는가, 그리고 추정으로 이은 변이 신뢰도에 그대로 드러나는가다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.recognize import opening_close as O  # noqa: E402
from core.design.recognize import params as P  # noqa: E402
from core.design.recognize import room_faces as R  # noqa: E402
from core.design.recognize import wall_centerline as W  # noqa: E402

# 100m x 100m. 실 면적 상한(bbox의 0.5)에 걸리지 않게 넉넉히 둔다.
_BBOX = 1.0e10


def _result(*segments, unpaired=()):
    lines = [W.Centerline(p1=(s[0], s[1]), p2=(s[2], s[3]),
                          thickness_mm=None if i in unpaired else 200.0,
                          source_pair=(i, i), unpaired=i in unpaired,
                          confidence=P.CONF_CENTERLINE_PAIRED)
             for i, s in enumerate(segments)]
    nodes, degree = W.snap_endpoints(lines)
    return W.CenterlineResult(centerlines=lines, nodes=nodes,
                              node_degree=degree,
                              wall_repr=W.WALL_REPR_DOUBLE, paired_ratio=1.0)


def _closure(result, *gaps):
    edges = []
    for (x1, y1), (x2, y2) in gaps:
        n1 = result.nodes.index((float(x1), float(y1)))
        n2 = result.nodes.index((float(x2), float(y2)))
        edges.append(O.VirtualEdge(
            p1=result.nodes[n1], p2=result.nodes[n2], n1=n1, n2=n2,
            kind=O.OPENING, confidence=P.CONF_VE_OPENING,
            gap_mm=0.0, evidence=["테스트"]))
    return O.ClosureResult(virtual_edges=edges, open_endpoints=0, relaxed=False)


def _square(x0, y0, x1, y1):
    return ((x0, y0, x1, y0), (x1, y0, x1, y1), (x1, y1, x0, y1), (x0, y1, x0, y0))


def _build(result, closure=None, bbox=_BBOX):
    return R.build_faces(result, closure, bbox_area_mm2=bbox)


# ── face 추출 ───────────────────────────────────────────────────────────

def test_닫힌_사각형_하나는_실_하나가_된다():
    faces = _build(_result(*_square(0, 0, 5000, 5000))).faces
    assert len(faces) == 1
    assert faces[0].area_m2 == pytest.approx(25.0)
    assert faces[0].edge_count == 4


def test_외곽_face_는_실이_아니다():
    """§3.5 5항. 이걸 안 버리면 건물 바깥이 방이 된다."""
    out = _build(_result(*_square(0, 0, 5000, 5000)))
    assert out.dropped["outer"] == 1


def test_연결_요소마다_생기는_외곽을_전부_버린다():
    """떨어져 있는 두 실은 외곽 face 도 두 개다."""
    out = _build(_result(*_square(0, 0, 5000, 5000),
                         *_square(20000, 0, 25000, 5000)))
    assert len(out.faces) == 2
    assert out.dropped["outer"] == 2


def test_벽을_공유한_두_실은_따로_나온다():
    """하나로 합쳐지면 면적이 두 배가 되고 헤드 개수가 틀린다."""
    faces = _build(_result(
        (0, 0, 2500, 0), (2500, 0, 5000, 0), (5000, 0, 5000, 5000),
        (5000, 5000, 2500, 5000), (2500, 5000, 0, 5000), (0, 5000, 0, 0),
        (2500, 0, 2500, 5000))).faces
    assert [round(f.area_m2, 2) for f in faces] == [12.5, 12.5]


def test_벽을_가로지르는_선은_교차점에서_잘린다():
    """§3.5 1항. 안 자르면 두 실을 가른 벽이 아무것도 가르지 못한다."""
    out = _build(_result(*_square(0, 0, 5000, 5000), (0, 2500, 5000, 2500)))
    assert [round(f.area_m2, 2) for f in out.faces] == [12.5, 12.5]


def test_십자로_교차하면_실이_넷이_된다():
    faces = _build(_result(*_square(0, 0, 5000, 5000),
                           (0, 2500, 5000, 2500),
                           (2500, 0, 2500, 5000))).faces
    assert [round(f.area_m2, 2) for f in faces] == [6.25] * 4


# ── face 필터 ───────────────────────────────────────────────────────────

def test_면적이_작으면_실이_아니라_벽_사이_틈이다():
    out = _build(_result(*_square(0, 0, 500, 500)))
    assert out.faces == []
    assert out.dropped["too_small"] == 1


def test_bbox_의_절반을_넘으면_외곽_오검출로_본다():
    out = _build(_result(*_square(0, 0, 5000, 5000)), bbox=30.0e6)
    assert out.faces == []
    assert out.dropped["too_large"] == 1


def test_열린_ㄷ자는_실이_되지_않는다():
    """C160 이 못 닫은 간극은 여기서 지어내 채우지 않는다."""
    assert _build(_result((0, 0, 5000, 0), (5000, 0, 5000, 5000),
                          (5000, 5000, 0, 5000))).faces == []


# ── 가상 간선 · 신뢰도 ──────────────────────────────────────────────────

def test_가상_간선으로_닫힌_실은_신뢰도가_깎인다():
    result = _result((0, 0, 5000, 0), (5000, 0, 5000, 5000),
                     (5000, 5000, 0, 5000), (0, 5000, 0, 1000))
    face = _build(result, _closure(result, ((0, 1000), (0, 0)))).faces[0]
    assert face.virtual_ratio == pytest.approx(1000.0 / 20000.0)
    assert face.confidence == pytest.approx(
        P.FACE_CONF_BASE - P.FACE_CONF_VIRTUAL_PENALTY * face.virtual_ratio)


def test_대부분_추정으로_닫힌_실은_플래그가_붙는다():
    """§3.5 — 검수 우선순위 상위. 여기가 틀리면 없던 실이 생긴 것이다."""
    result = _result((0, 0, 2000, 0), (2000, 5000, 0, 5000))
    closure = _closure(result, ((2000, 0), (2000, 5000)), ((0, 5000), (0, 0)))
    face = _build(result, closure).faces[0]
    assert R.MOSTLY_VIRTUAL in face.flags
    assert face.virtual_ratio > P.FACE_VIRTUAL_RATIO_SUSPICIOUS


def test_두께_미상_중심선은_신뢰도를_깎는다():
    result = _result(*_square(0, 0, 5000, 5000), unpaired=(0,))
    face = _build(result).faces[0]
    assert face.unpaired_ratio == pytest.approx(0.25)
    assert face.confidence == pytest.approx(
        P.FACE_CONF_BASE - P.FACE_CONF_UNPAIRED_PENALTY * 0.25)


def test_변이_많으면_신뢰도를_한_번_더_깎는다():
    n = P.FACE_EDGE_COUNT_PENALTY_MIN + 4
    step = 20000.0 / n
    pts = [(0.0, 0.0)]
    for i in range(1, n):
        pts.append((step * i, 0.0) if i * step <= 10000
                   else (10000.0, step * i - 10000.0))
    segs = [(*pts[i], *pts[(i + 1) % n]) for i in range(n)]
    face = _build(_result(*segs)).faces[0]
    assert face.edge_count == n
    assert face.confidence == pytest.approx(
        P.FACE_CONF_BASE - P.FACE_CONF_MANY_EDGES_PENALTY)


# ── 운영 ────────────────────────────────────────────────────────────────

def test_간선이_없어도_터지지_않는다():
    out = R.build_faces(W.build_centerlines([], offset_peaks_mm=[150.0]))
    assert out.faces == []
    assert any("간선이 없" in line for line in out.provenance)


def test_결과는_직렬화된다():
    dumped = _build(_result(*_square(0, 0, 5000, 5000))).to_dict()
    face = dumped["faces"][0]
    assert face["area_m2"] == pytest.approx(25.0)
    assert face["provenance"]
    assert dumped["dropped"]["outer"] == 1


def test_버린_개수가_provenance_에_남는다():
    """실이 안 나온 이유를 검수자가 알 수 있어야 한다."""
    out = _build(_result(*_square(0, 0, 500, 500)))
    assert any("면적 미달 1" in line for line in out.provenance)


def test_임계값이_코드에_박혀_있지_않다():
    src = (_ROOT / "core" / "design" / "recognize" / "room_faces.py").read_text(encoding="utf-8")
    for name in ("FACE_AREA_MIN_M2", "FACE_AREA_MAX_BBOX_RATIO",
                 "FACE_EDGE_COUNT_SUSPICIOUS", "FACE_VIRTUAL_RATIO_SUSPICIOUS",
                 "FACE_CONF_BASE", "FACE_SPLIT_TOL_MM"):
        assert f"P.{name}" in src, f"{name} 을 params 에서 읽지 않는다"
