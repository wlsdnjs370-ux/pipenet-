# -*- coding: utf-8 -*-
"""지시서 §3.2 — C140 카테고리 판정.

여기서 지키는 것은 정답률이 아니라 **판정의 위계**다. 지문이 이름을 이기는지,
이름만으로는 아무 카테고리도 나오지 않는지, 판정할 수 없는 것을 판정한 척하지
않는지. 정답률은 실 도면 벤치마크(PR-4e)의 몫이다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.recognize import arch_category as C  # noqa: E402
from core.design.recognize import params as P  # noqa: E402
from core.design.recognize.geom_stats import LayerFingerprint  # noqa: E402


def _wall_fp(name="L1", **kwargs):
    base = dict(parallel_pair_ratio=0.80, offset_peaks_mm=[150.0], len_median_mm=3000.0)
    base.update(kwargs)
    return LayerFingerprint(name=name, n_entities=10, type_hist={"L": 10}, **base)


def _door_fp(name="L1", **kwargs):
    base = dict(arc_attach_ratio=0.90, door_radius_ratio=0.90)
    base.update(kwargs)
    return LayerFingerprint(name=name, n_entities=10, type_hist={"A": 5, "L": 5}, **base)


# ── 지문이 1순위 ────────────────────────────────────────────────────────

def test_지문이_맞으면_이름이_없어도_판정한다():
    """이름 규약은 설계사마다 다르다. 이름 없이도 서야 한다."""
    assert C.arch_category(_wall_fp(name="0"), None) == (C.WALL, P.CONF_WALL)


def test_이름만으로는_어떤_카테고리도_나오지_않는다():
    """§3.2 — 이름은 가산점 전용, 판정 단독 근거 금지."""
    empty = LayerFingerprint(name="A-WALL", n_entities=3, type_hist={"L": 3})
    assert C.arch_category(empty, None)[0] == C.OTHER


def test_이름이_맞으면_신뢰도만_올라간다():
    plain = C.evaluate(_wall_fp(name="0"))
    hinted = C.evaluate(_wall_fp(name="A-WALL"))
    assert hinted.category == plain.category == C.WALL
    assert hinted.confidence == pytest.approx(plain.confidence + P.NAME_HINT_BONUS)
    assert hinted.name_bonus is True


def test_이름이_틀려도_지문을_뒤집지_못한다():
    """벽 지문에 DOOR 이름이 붙어도 WALL 이다. 가산점은 성립한 규칙에만 붙는다."""
    verdict = C.evaluate(_wall_fp(name="A-DOOR"))
    assert verdict.category == C.WALL
    assert verdict.name_bonus is False


def test_신뢰도는_1을_넘지_않는다():
    fp = _door_fp(name="A-DOOR")
    assert C.evaluate(fp).confidence <= 1.0


# ── 판정표 각 행 ────────────────────────────────────────────────────────

def test_문은_호가_선_끝점에_붙고_반경이_맞아야_한다():
    assert C.arch_category(_door_fp(), None) == (C.DOOR, P.CONF_DOOR)


def test_반경이_문_범위가_아니면_문이_아니다():
    assert C.arch_category(_door_fp(door_radius_ratio=0.1), None)[0] == C.OTHER


def test_기둥은_반복성과_격자정렬과_면적을_모두_본다():
    fp = LayerFingerprint(name="X", n_entities=9, type_hist={"PL": 9},
                          closed_repeat_score=0.95, grid_alignment_score=0.90,
                          closed_area_median_m2=0.64, small_closed_count=9)
    assert C.arch_category(fp, None) == (C.COLUMN, P.CONF_COLUMN)


def test_계단은_등간격_다발이_필요하다():
    fp = LayerFingerprint(name="X", n_entities=8, type_hist={"L": 8},
                          stair_bundle_max=8, stair_gap_cv=0.02)
    assert C.arch_category(fp, None) == (C.STAIR, P.CONF_STAIR)


def test_거대_레이어의_우연한_다발은_계단이_아니다():
    """선 2만 개짜리 레이어에서는 등간격 6선 다발이 우연히 거의 항상 나온다."""
    fp = LayerFingerprint(name="X", n_entities=20000, type_hist={"L": 20000},
                          stair_bundle_max=6, stair_gap_cv=0.02)
    assert C.arch_category(fp, None)[0] != C.STAIR


def test_치수는_숫자_텍스트_비율로_판정한다():
    fp = LayerFingerprint(name="X", n_entities=10, type_hist={"T": 10},
                          text_numeric_ratio=0.95, text_len_median=5)
    assert C.arch_category(fp, None) == (C.DIM, P.CONF_DIM)


def test_실명_텍스트는_숫자가_아니고_짧다():
    fp = LayerFingerprint(name="X", n_entities=10, type_hist={"T": 10},
                          text_numeric_ratio=0.05, text_len_median=4)
    assert C.arch_category(fp, None) == (C.ROOM_TEXT, P.CONF_ROOM_TEXT)


def test_실명_텍스트는_긴_문장을_배제한다():
    fp = LayerFingerprint(name="X", n_entities=10, type_hist={"T": 10},
                          text_numeric_ratio=0.05,
                          text_len_median=P.ROOM_TEXT_LEN_MAX + 1)
    assert C.arch_category(fp, None)[0] == C.OTHER


def test_통심은_초장선과_원형심볼을_함께_요구한다():
    fp = LayerFingerprint(name="X", n_entities=10, type_hist={"L": 8, "C": 2},
                          long_line_ratio=0.80)
    assert C.arch_category(fp, None) == (C.GRID, P.CONF_GRID)


def test_원형심볼이_없으면_통심이_아니다():
    """통심선과 부호가 다른 레이어로 갈라진 도면이 흔하다 — 그때는 서지 않는다."""
    fp = LayerFingerprint(name="A-GRID", n_entities=10, type_hist={"L": 10},
                          long_line_ratio=0.80)
    assert C.arch_category(fp, None)[0] != C.GRID


def test_가구는_아무_규칙도_안_걸린_뒤에만_본다():
    fp = LayerFingerprint(name="X", n_entities=10, type_hist={"I": 10})
    assert C.arch_category(fp, None) == (C.FURNITURE, P.CONF_FURNITURE)


def test_인서트가_밀집해도_다른_규칙이_이긴다():
    fp = _wall_fp(name="X")
    fp.type_hist = {"I": 100, "L": 10}
    assert C.arch_category(fp, None)[0] == C.WALL


# ── 샤프트: 모르는 것을 아는 척하지 않는다 ──────────────────────────────

def _shaft_fp():
    return LayerFingerprint(name="X", n_entities=4, type_hist={"PL": 4},
                            closed_area_median_m2=2.0, small_closed_count=4)


def test_다층_반복을_확인하지_못하면_신뢰도가_묶인다():
    verdict = C.evaluate(_shaft_fp())
    assert verdict.category == C.SHAFT
    assert verdict.confidence == pytest.approx(P.CONF_SHAFT_SINGLE)
    assert any("다층 반복을 확인하지 못해" in p for p in verdict.provenance)


def test_먼지같은_폐합은_샤프트가_아니다():
    """실 도면 가구 레이어의 0.001㎡ 폐합 수천 개가 통째로 SHAFT 가 됐었다."""
    fp = LayerFingerprint(name="X", n_entities=1000, type_hist={"PL": 1000},
                          closed_area_median_m2=0.001, small_closed_count=1000)
    assert C.arch_category(fp, None)[0] != C.SHAFT


def test_다층_반복이_확인되면_신뢰도가_올라간다():
    verdict = C.evaluate(_shaft_fp(), multi_floor_repeat=True)
    assert verdict.confidence == pytest.approx(P.CONF_SHAFT_MULTIFLOOR)


# ── 근거와 탈락 사유 ────────────────────────────────────────────────────

def test_한_조건만_어긋나면_탈락_사유가_남는다():
    """실 도면 ARCHI 가 이 경우다 — 평행쌍도 두께 peak 도 맞는데 길이에서 진다."""
    verdict = C.evaluate(_wall_fp(len_median_mm=636.0))
    assert verdict.category != C.WALL
    assert any(text.startswith(f"{C.WALL}: 선 길이 중앙값") for text in verdict.near_misses)


def test_두_조건짜리_규칙은_스친_것으로_치지_않는다():
    """반만 맞은 것은 스친 게 아니다. 다 실으면 진짜 근거가 묻힌다."""
    verdict = C.evaluate(_door_fp(door_radius_ratio=0.1))
    assert not any(t.startswith(f"{C.DOOR}:") for t in verdict.near_misses)


def test_함께_성립한_카테고리는_대안으로_남는다():
    """한 레이어에 벽과 문이 섞인 도면이 있다. 이긴 쪽만 보이면 못 알아챈다."""
    fp = _wall_fp()
    fp.arc_attach_ratio, fp.door_radius_ratio = 0.90, 0.90
    verdict = C.evaluate(fp)
    assert verdict.category == C.WALL
    assert C.DOOR in [c for c, _ in verdict.alternatives]


def test_판정한_근거가_반드시_남는다():
    """작업 규칙 9 — 인식 셸 결과에는 항상 confidence 와 provenance."""
    verdict = C.evaluate(_wall_fp())
    assert len(verdict.provenance) == 3
    assert C.OTHER not in verdict.provenance


def test_아무것도_안_걸리면_봤다는_사실을_남긴다():
    """미분류(안 봄)와 OTHER(보고 나서 그 밖)는 다른 상태다."""
    verdict = C.evaluate(LayerFingerprint(name="X", n_entities=3, type_hist={"L": 3}))
    assert verdict.category == C.OTHER
    assert verdict.confidence == pytest.approx(P.OTHER_CONFIDENCE)
    assert verdict.provenance


def test_판정할_수_없는_두_종은_그렇다고_말한다():
    """WINDOW 는 C150 이 있어야 하고, BEAM 은 표에 지문 조건이 없다."""
    verdict = C.evaluate(_wall_fp())
    assert any(t.startswith(f"{C.WINDOW}:") for t in verdict.undecidable)
    assert any(t.startswith(f"{C.BEAM}:") for t in verdict.undecidable)


@pytest.mark.parametrize("category", [C.WINDOW, C.BEAM])
def test_판정_불가한_카테고리는_결과로_나오지_않는다(category):
    """이름이 맞아도 지어내지 않는다."""
    names = C.NAME_HINTS[category]
    assert names
    for name in names:
        fp = LayerFingerprint(name=name, n_entities=5, type_hist={"L": 5})
        assert C.arch_category(fp, None)[0] != category


# ── 운영 ────────────────────────────────────────────────────────────────

def test_결과는_JSON_으로_나간다():
    payload = C.evaluate(_wall_fp(name="A-WALL")).to_dict()
    assert payload["category"] == C.WALL
    assert payload["layer"] == "A-WALL"
    assert isinstance(payload["provenance"], list)
    assert isinstance(payload["confidence"], float)


def test_이름_힌트를_따로_줄_수_있다():
    """레이어 이름과 판정에 쓸 이름이 다른 경우(외부참조 접두어 등)."""
    verdict = C.evaluate(_wall_fp(name="X_PLAN_O$0$A-WALL-1"), "A-WALL")
    assert verdict.name_bonus is True


def test_임계값이_코드에_박혀_있지_않다():
    src = (_ROOT / "core" / "design" / "recognize" / "arch_category.py").read_text(encoding="utf-8")
    for name in ("NAME_HINT_BONUS", "CONF_WALL", "WALL_PARALLEL_MIN_RATIO",
                 "DOOR_ARC_ATTACH_MIN_RATIO", "COLUMN_REPEAT_MIN", "OTHER_CONFIDENCE"):
        assert f"P.{name}" in src, f"{name} 을 params 에서 읽지 않는다"


def test_모듈A_레이어_사전과_독립이다():
    """§3.2 단서 — 모듈 A 가 그 사전에 의존하므로 공유하지 않고 따로 둔다."""
    import subprocess
    probe = ("import sys; import core.design.recognize.arch_category; "
             "sys.exit(1 if 'sprinkler_remote30_extractor' in sys.modules else 0)")
    assert subprocess.run([sys.executable, "-c", probe], cwd=_ROOT).returncode == 0
    assert C.NAME_HINTS[C.WALL], "인식 셸은 자기 사전을 가진다"
