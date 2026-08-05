# -*- coding: utf-8 -*-
"""지시서 §3.1 — C130 기하 지문.

인식 셸은 도면이 바뀌면 깨지고 **조용히 틀린다**(부록 B). 여기 단위 테스트가
지키는 것은 정답률이 아니라 지문의 **의미**다. 평행쌍이 각도·겹침·오프셋 세
조건을 모두 요구하는지, 벽 두께가 bin 중심으로 반올림되지 않는지, 단위가 mm 로
환산되는지. 정답률은 실 도면 벤치마크(PR-4e)의 몫이다.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.recognize import geom_stats as G  # noqa: E402
from core.design.recognize import params as P  # noqa: E402


def _line(x1, y1, x2, y2, layer="A-WALL"):
    return {"t": "L", "l": layer, "p": [x1, y1, x2, y2]}


def _fp(entities, **kwargs):
    result = G.fingerprints(entities, **kwargs)
    assert len(result) == 1
    return result[0]


# ── 평행쌍 ──────────────────────────────────────────────────────────────

def test_나란한_두_선은_평행쌍이고_두께는_bin_중심이_아니다():
    fp = _fp([_line(0, 0, 5000, 0), _line(0, 150, 5000, 150)])
    assert fp.parallel_pair_ratio == 1.0
    assert fp.offset_peaks_mm == [150.0]


def test_각도가_틀어지면_평행쌍이_아니다():
    """허용 2°. 이걸 넓히면 비스듬한 잡선이 전부 벽이 된다."""
    fp = _fp([_line(0, 0, 5000, 0), _line(0, 150, 5000, 150 + 400)])  # 약 4.6°
    assert fp.parallel_pair_ratio == 0.0


def test_스쳐_지나가면_평행쌍이_아니다():
    """겹침 30% 미만 — 서로 다른 벽의 조각이 짝지어지는 것을 막는다."""
    fp = _fp([_line(0, 0, 1000, 0), _line(1200, 150, 2200, 150)])
    assert fp.parallel_pair_ratio == 0.0


@pytest.mark.parametrize("offset", [30, 600])
def test_벽_두께_범위_밖은_평행쌍이_아니다(offset):
    fp = _fp([_line(0, 0, 5000, 0), _line(0, offset, 5000, offset)])
    assert fp.parallel_pair_ratio == 0.0


def test_소수_두께는_peak_에서_탈락한다():
    """8% 미만은 노이즈다. 남기면 C150 이 없는 벽 두께를 찾아다닌다."""
    ents = []
    for i in range(20):
        ents += [_line(0, i * 3000, 5000, i * 3000),
                 _line(0, i * 3000 + 200, 5000, i * 3000 + 200)]
    ents += [_line(0, 90000, 5000, 90000), _line(0, 90060, 5000, 90060)]
    fp = _fp(ents)
    assert fp.offset_peaks_mm == [200.0]


def test_미터_단위_도면도_mm_로_환산된다():
    """임계값이 전부 mm 라 unit_to_mm 이 틀리면 판정 전체가 함께 틀린다."""
    fp = _fp([_line(0, 0, 5, 0), _line(0, 0.15, 5, 0.15)], unit_to_mm=1000.0)
    assert fp.parallel_pair_ratio == 1.0
    assert fp.offset_peaks_mm == [150.0]
    assert fp.len_median_mm == pytest.approx(5000.0)


# ── 문 / 초장선 ─────────────────────────────────────────────────────────

def test_문호는_line_끝점에_붙고_반경_범위에_든다():
    arc = {"t": "A", "l": "A-DOOR", "c": [0, 0], "r": 900, "a": [0.0, 90.0]}
    fp = _fp([arc, _line(0, 0, 900, 0, layer="A-DOOR")])
    assert fp.arc_attach_ratio == 1.0
    assert fp.door_radius_ratio == 1.0


def test_떨어진_호는_붙지_않은_것으로_센다():
    arc = {"t": "A", "l": "A-DOOR", "c": [0, 0], "r": 900, "a": [0.0, 90.0]}
    fp = _fp([arc, _line(5000, 5000, 5900, 5000, layer="A-DOOR")])
    assert fp.arc_attach_ratio == 0.0
    assert fp.door_radius_ratio == 1.0


def test_통심선은_도면_대각_대비로_판정한다():
    bbox = {"minx": 0, "miny": 0, "maxx": 30000, "maxy": 0}
    fp = _fp([_line(0, 0, 30000, 0), _line(0, 1000, 500, 1000)], bbox=bbox)
    assert fp.long_line_ratio == 0.5


# ── 폐합 도형 ───────────────────────────────────────────────────────────

def test_폐합_폴리라인은_면적으로_센다():
    pl = {"t": "PL", "l": "S-COL", "p": [[0, 0], [800, 0], [800, 800], [0, 800], [0, 0]]}
    fp = _fp([pl])
    assert fp.closed_shape_count == 1
    assert fp.closed_area_median_m2 == pytest.approx(0.64)
    assert fp.small_closed_count == 1


def test_열린_폴리라인은_폐합이_아니다():
    pl = {"t": "PL", "l": "S-COL", "p": [[0, 0], [800, 0], [800, 800], [0, 800]]}
    assert _fp([pl]).closed_shape_count == 0


def test_네_개의_line_이_이루는_사각형도_폐합으로_센다():
    """캔버스 엔티티에 닫힘 플래그가 없어 놓치는 몫을 여기서 메운다."""
    ents = [_line(0, 0, 800, 0, "S-COL"), _line(800, 0, 800, 800, "S-COL"),
            _line(800, 800, 0, 800, "S-COL"), _line(0, 800, 0, 0, "S-COL")]
    fp = _fp(ents)
    assert fp.closed_shape_count == 1
    assert fp.closed_area_median_m2 == pytest.approx(0.64)


def test_같은_크기_기둥이_반복되면_반복성과_격자정렬이_높다():
    ents = []
    for row in range(3):
        for col in range(3):
            x, y = col * 6000, row * 6000
            ents.append({"t": "PL", "l": "S-COL",
                         "p": [[x, y], [x + 600, y], [x + 600, y + 600], [x, y + 600], [x, y]]})
    fp = _fp(ents)
    assert fp.closed_shape_count == 9
    assert fp.closed_repeat_score == pytest.approx(1.0)
    assert fp.grid_alignment_score == pytest.approx(1.0)


def test_크기가_제각각이면_반복성이_떨어진다():
    ents = [{"t": "PL", "l": "X",
             "p": [[0, i * 9000], [s, i * 9000], [s, i * 9000 + s], [0, i * 9000 + s], [0, i * 9000]]}
            for i, s in enumerate((500, 2000, 6000))]
    assert _fp(ents).closed_repeat_score < 0.6


# ── 계단 / 텍스트 ───────────────────────────────────────────────────────

def test_등간격_평행_단선_다발을_찾는다():
    ents = [_line(0, i * 300, 1200, i * 300, "A-STRS") for i in range(8)]
    fp = _fp(ents)
    assert fp.stair_bundle_max == 8
    assert fp.stair_gap_cv == pytest.approx(0.0)


def test_간격이_들쭉날쭉하면_다발이_끊긴다():
    ys = [0, 300, 600, 1700, 2000, 2300, 2600, 2900]
    fp = _fp([_line(0, y, 1200, y, "A-STRS") for y in ys])
    assert fp.stair_bundle_max < len(ys)


def test_치수_텍스트와_실명_텍스트를_비율로_가른다():
    ents = [{"t": "T", "l": "A-DIMS", "v": v} for v in ("3,600", "1200", "Ø100", "사무실")]
    assert _fp(ents).text_numeric_ratio == pytest.approx(0.75)


# ── 운영 ────────────────────────────────────────────────────────────────

def test_거대_레이어는_표본만_보되_전체_개수는_보고한다():
    """부록 C.1 — 표본을 썼다는 사실을 숨기면 지문을 믿을 수 없다."""
    n = P.LAYER_SAMPLE_LIMIT + 5
    fp = _fp([{"t": "T", "l": "BIG", "v": "1"} for _ in range(n)])
    assert fp.n_entities == n
    assert fp.sampled is True
    assert sum(fp.type_hist.values()) == P.LAYER_SAMPLE_LIMIT


def test_레이어별로_따로_센다():
    result = G.fingerprints([_line(0, 0, 100, 0, "B"), _line(0, 0, 100, 0, "A")])
    assert [f.name for f in result] == ["A", "B"]


def test_빈_입력에도_터지지_않는다():
    assert G.fingerprints([]) == []


@pytest.mark.parametrize("diag,expected", [
    (30_000, 1.0),        # 30m — mm 도면
    (30, 1000.0),         # 30 — m 도면
])
def test_단위는_제안만_하고_적용하지_않는다(diag, expected):
    got = G.suggest_unit_to_mm({"minx": 0, "miny": 0, "maxx": diag, "maxy": 0})
    assert got["unit_to_mm"] == expected
    assert got["confidence"] < 1.0


def test_어떤_단위로도_말이_안_되면_후보를_내지_않는다():
    got = G.suggest_unit_to_mm({"minx": 0, "miny": 0, "maxx": 1e12, "maxy": 0})
    assert got["unit_to_mm"] is None


def test_파라미터는_캐시키로_직렬화된다():
    """§11.2 — 이게 깨지면 파라미터를 튜닝해도 옛 결과가 그대로 나온다."""
    dumped = json.dumps(P.recognize_params(), sort_keys=True)
    assert "PARALLEL_ANGLE_TOL_DEG" in dumped
    assert "LAYER_SAMPLE_LIMIT" in dumped


def test_임계값이_코드에_박혀_있지_않다():
    """§3.1 — 하드코딩 금지. 튜닝 이력이 남지 않으면 벤치마크가 무의미하다."""
    src = (_ROOT / "core" / "design" / "recognize" / "geom_stats.py").read_text(encoding="utf-8")
    for name in ("PARALLEL_ANGLE_TOL_DEG", "PARALLEL_OVERLAP_MIN_RATIO",
                 "OFFSET_HIST_BIN_MM", "LAYER_SAMPLE_LIMIT"):
        assert f"P.{name}" in src, f"{name} 을 params 에서 읽지 않는다"
