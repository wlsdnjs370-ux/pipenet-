# -*- coding: utf-8 -*-
"""지시서 §3.3 — C150 벽 중심선화.

여기 테스트가 지키는 것은 "몇 개나 짝지었나" 가 아니라 **짝의 조건**이다. 두께가
C130 peak 과 맞는지, 한 선이 두 쌍에 들어가지 않는지, 짝을 못 찾은 선이 조용히
사라지지 않는지. 정답률은 실 도면 벤치마크(PR-4e)의 몫이다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.recognize import params as P  # noqa: E402
from core.design.recognize import wall_centerline as W  # noqa: E402


def _one(lines, peaks=(150.0,), **kwargs):
    result = W.build_centerlines(lines, offset_peaks_mm=peaks, **kwargs)
    paired = [c for c in result.centerlines if not c.unpaired]
    assert len(paired) == 1
    return paired[0]


# ── 쌍 판정 ─────────────────────────────────────────────────────────────

def test_나란한_두_선은_한_중심선이_되고_두께가_실린다():
    c = _one([(0, 0, 5000, 0), (0, 150, 5000, 150)])
    assert c.thickness_mm == pytest.approx(150.0)
    assert c.p1 == pytest.approx((0.0, 75.0))
    assert c.p2 == pytest.approx((5000.0, 75.0))
    assert c.source_pair == (0, 1)


def test_C130_peak_과_어긋난_오프셋은_짝이_아니다():
    """두께 후보 없이 짝지으면 나란한 가구 선 두 개도 벽이 된다."""
    result = W.build_centerlines([(0, 0, 5000, 0), (0, 300, 5000, 300)],
                                 offset_peaks_mm=[150.0])
    assert all(c.unpaired for c in result.centerlines)


def test_peak_이_하나도_없으면_아무_쌍도_성립하지_않는다():
    result = W.build_centerlines([(0, 0, 5000, 0), (0, 150, 5000, 150)],
                                 offset_peaks_mm=[])
    assert all(c.unpaired for c in result.centerlines)
    assert any("peak" in line for line in result.provenance)


def test_각도가_틀어지면_짝이_아니다():
    result = W.build_centerlines([(0, 0, 5000, 0), (0, 150, 5000, 550)],
                                 offset_peaks_mm=[150.0])
    assert all(c.unpaired for c in result.centerlines)


def test_스쳐_지나가면_짝이_아니다():
    result = W.build_centerlines([(0, 0, 1000, 0), (1200, 150, 2200, 150)],
                                 offset_peaks_mm=[150.0])
    assert all(c.unpaired for c in result.centerlines)


def test_중심선은_겹치는_구간만_낸다():
    """긴 선 옆에 짧은 선이 붙으면 짧은 쪽이 끝나는 데서 중심선도 끝난다."""
    c = _one([(0, 0, 5000, 0), (1000, 150, 3000, 150)])
    assert c.p1 == pytest.approx((1000.0, 75.0))
    assert c.p2 == pytest.approx((3000.0, 75.0))


def test_한_선은_한_쌍에만_들어간다():
    """벽선 옆 해치선. 전부 받으면 같은 벽에서 중심선이 둘 나온다."""
    result = W.build_centerlines(
        [(0, 0, 5000, 0), (0, 150, 5000, 150), (0, 300, 5000, 300)],
        offset_peaks_mm=[150.0])
    assert sum(1 for c in result.centerlines if not c.unpaired) == 1
    assert sum(1 for c in result.centerlines if c.unpaired) == 1


# ── 미짝 보존 ───────────────────────────────────────────────────────────

def test_짝을_못_찾은_선은_버리지_않고_두께_미상으로_보존한다():
    """§3.3 5항 — 조적벽을 단선으로 그린 도면이 있다."""
    result = W.build_centerlines([(0, 0, 5000, 0)], offset_peaks_mm=[150.0])
    assert len(result.centerlines) == 1
    lone = result.centerlines[0]
    assert lone.unpaired is True
    assert lone.thickness_mm is None
    assert lone.confidence < P.CONF_CENTERLINE_PAIRED


def test_평행쌍이_드물면_단선_표기_도면으로_보고_완화_모드가_켜진다():
    """§3.3 실패 신호. 조용히 빈 결과를 내면 C170 이 실을 못 찾은 이유를 모른다."""
    lines = [(0, i * 3000, 5000, i * 3000) for i in range(10)]
    result = W.build_centerlines(lines, offset_peaks_mm=[150.0])
    assert result.wall_repr == W.WALL_REPR_SINGLE
    assert result.relaxed is True


def test_대부분_짝지으면_이중선_표기다():
    lines = []
    for i in range(5):
        lines += [(0, i * 3000, 5000, i * 3000),
                  (0, i * 3000 + 150, 5000, i * 3000 + 150)]
    result = W.build_centerlines(lines, offset_peaks_mm=[150.0])
    assert result.wall_repr == W.WALL_REPR_DOUBLE
    assert result.relaxed is False


# ── 끝점 스냅 ───────────────────────────────────────────────────────────

def test_공차_안의_끝점은_같은_노드로_묶인다():
    lines = [(0, 0, 3000, 0), (0, 150, 3000, 150),
             (3010, 0, 6000, 0), (3010, 150, 6000, 150)]
    result = W.build_centerlines(lines, offset_peaks_mm=[150.0])
    assert sorted(result.node_degree) == [1, 1, 2]


def test_공차_밖의_끝점은_따로_남는다():
    lines = [(0, 0, 3000, 0), (0, 150, 3000, 150),
             (3200, 0, 6000, 0), (3200, 150, 6000, 150)]
    result = W.build_centerlines(lines, offset_peaks_mm=[150.0])
    assert sorted(result.node_degree) == [1, 1, 1, 1]


def test_스냅_공차는_벽_두께에_걸린다():
    """§3.3 6항 — 두꺼운 벽일수록 접합부가 벌어져 있다."""
    assert W.snap_tol(None) == P.CENTERLINE_SNAP_TOL_MIN_MM
    assert W.snap_tol(100.0) == P.CENTERLINE_SNAP_TOL_MIN_MM
    assert W.snap_tol(400.0) == pytest.approx(120.0)


# ── 운영 ────────────────────────────────────────────────────────────────

def test_미터_단위_도면도_mm_로_환산된다():
    c = _one([(0, 0, 5, 0), (0, 0.15, 5, 0.15)], unit_to_mm=1000.0)
    assert c.thickness_mm == pytest.approx(150.0)
    assert c.p2 == pytest.approx((5000.0, 75.0))


def test_빈_입력에도_터지지_않는다():
    result = W.build_centerlines([], offset_peaks_mm=[150.0])
    assert result.centerlines == []
    assert result.nodes == []


def test_길이가_0인_선은_중심선이_되지_않는다():
    result = W.build_centerlines([(100, 100, 100, 100)], offset_peaks_mm=[150.0])
    assert result.centerlines == []


def test_결과는_직렬화된다():
    """NDJSON(§11.1)으로 나가야 하므로 dict 로 떨어져야 한다."""
    result = W.build_centerlines([(0, 0, 5000, 0), (0, 150, 5000, 150)],
                                 offset_peaks_mm=[150.0])
    dumped = result.to_dict()
    assert dumped["wall_repr"] == W.WALL_REPR_DOUBLE
    assert dumped["centerlines"][0]["thickness_mm"] == 150.0
    assert dumped["centerlines"][0]["provenance"]


def test_임계값이_코드에_박혀_있지_않다():
    """§3.1 — 하드코딩 금지. 튜닝 이력이 남지 않으면 벤치마크가 무의미하다."""
    src = (_ROOT / "core" / "design" / "recognize" / "wall_centerline.py").read_text(encoding="utf-8")
    for name in ("CENTERLINE_ANGLE_TOL_DEG", "CENTERLINE_OFFSET_MATCH_TOL_MM",
                 "CENTERLINE_OVERLAP_MIN_RATIO", "CENTERLINE_SNAP_TOL_MIN_MM",
                 "SINGLE_LINE_PARALLEL_MAX_RATIO"):
        assert f"P.{name}" in src, f"{name} 을 params 에서 읽지 않는다"
