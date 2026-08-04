# -*- coding: utf-8 -*-
"""표고 출처 태깅 — FNCADnet 작업지시서 모듈 A T6.

표고 0 은 "수평이다"라는 주장이고, 근거가 없다는 것과는 다른 말이다. 예전엔
둘이 같은 0 이라 도면에 없는 낙차가 조용히 사라졌다. 이제 관로마다 값이
어디서 왔는지(사람 확정 / 도면 추정 / 관례 기본 / 미확정) 함께 남긴다.

실행::

    python -m pytest tests/test_elev_sources.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pytest  # noqa: E402

import remote30_prototype as rp  # noqa: E402
from remote30_constants import (  # noqa: E402
    ELEV_SOURCE_DEFAULT, ELEV_SOURCE_DRAWING, ELEV_SOURCE_UNRESOLVED,
    ELEV_SOURCE_USER, LOCAL_RISE_RULES, ZONE_KIND_PARKING,
    ZONE_KIND_UNIT_DWELLING,
)

# ── 계통도(라이저) — 수직 경로 위 층 라벨 3개 ──────────────────────
RISER_PATH = [(0.0, 6000.0), (0.0, 3000.0), (0.0, 0.0)]
FLOOR_LABELS = [(500.0, 6000.0, 18, "18층"),
                (500.0, 3000.0, 17, "17층"),
                (500.0, 0.0, 16, "16층")]
PROFILE_ROWS = [{"floor_label": "18층", "head_drop_m": 13.8},
                {"floor_label": "17층", "head_drop_m": 16.7},
                {"floor_label": "16층", "head_drop_m": 19.6}]


def _riser(*, floor_labels=None, profile_rows=None) -> dict:
    return rp._system_path_to_riser_dict(
        RISER_PATH, {}, RISER_PATH[0], RISER_PATH[-1],
        floor_labels=floor_labels, floor_profile_rows=profile_rows)


def _sources(items) -> set[str]:
    return {it["elev_source"] for it in items}


def test_압력표가_있으면_라이저_표고는_사람_확정():
    riser = _riser(floor_labels=FLOOR_LABELS, profile_rows=PROFILE_ROWS)
    assert _sources(riser["nodes"]) == {ELEV_SOURCE_USER}
    assert _sources(riser["pipes"]) == {ELEV_SOURCE_USER}
    # AV(16층) 기준 상대표고 — 위로 갈수록 +
    assert riser["nodes"][0]["elevation"] == pytest.approx(19.6 - 13.8)


def test_압력표가_없으면_도면_추정으로_내려간다():
    riser = _riser(floor_labels=FLOOR_LABELS)
    assert _sources(riser["nodes"]) == {ELEV_SOURCE_DRAWING}
    assert riser["floor_matching"]["height_source"] == ELEV_SOURCE_DRAWING


def test_층_라벨조차_없으면_관례_기본값():
    riser = _riser()
    assert _sources(riser["nodes"]) == {ELEV_SOURCE_DEFAULT}


def test_출처_분포를_노드_관로_각각_보고한다():
    riser = _riser(floor_labels=FLOOR_LABELS, profile_rows=PROFILE_ROWS)
    assert riser["elev_sources"]["nodes"][ELEV_SOURCE_USER] == len(riser["nodes"])
    assert riser["elev_sources"]["pipes"][ELEV_SOURCE_USER] == len(riser["pipes"])


def test_관로_표고는_양_끝_중_약한_출처를_따른다():
    """한 끝이 도면 추정이면 그 관로 전체가 도면 추정이다 — 좋은 쪽으로 반올림하지 않는다."""
    assert rp._weaker_elev_source(ELEV_SOURCE_USER, ELEV_SOURCE_DRAWING) == ELEV_SOURCE_DRAWING
    assert rp._weaker_elev_source(ELEV_SOURCE_USER, ELEV_SOURCE_USER) == ELEV_SOURCE_USER
    assert rp._weaker_elev_source(ELEV_SOURCE_DEFAULT,
                                  ELEV_SOURCE_UNRESOLVED) == ELEV_SOURCE_UNRESOLVED


# ── 평면도 국소 표고 ───────────────────────────────────────────────
PARKING_RECT = (2500.0, -500.0, 4500.0, 500.0)
INSIDE_A = (3000.0, 0.0)
INSIDE_B = (4000.0, 0.0)
OUTSIDE = (1000.0, 0.0)
PARKING_ZONE = [{"rect": PARKING_RECT, "kind": ZONE_KIND_PARKING}]
BEAM_DROP = LOCAL_RISE_RULES["parking_beam_drop_m"]
NIPPLE = LOCAL_RISE_RULES["upright_riser_nipple_m"]


def test_주차장_안에서만_도는_관로는_수평():
    assert rp._local_rise(INSIDE_A, INSIDE_B, PARKING_ZONE) == (0.0, ELEV_SOURCE_DEFAULT)


def test_주차장_경계를_넘으면_보_하단만큼_내려간다():
    assert rp._local_rise(OUTSIDE, INSIDE_A, PARKING_ZONE) == (BEAM_DROP, ELEV_SOURCE_DEFAULT)
    assert rp._local_rise(INSIDE_A, OUTSIDE, PARKING_ZONE) == (-BEAM_DROP, ELEV_SOURCE_DEFAULT)


def test_근거_없는_구역_전환은_0_으로_때우되_미확정으로_센다():
    zone = [{"rect": PARKING_RECT, "kind": ZONE_KIND_UNIT_DWELLING}]
    assert rp._local_rise(OUTSIDE, INSIDE_A, zone) == (0.0, ELEV_SOURCE_UNRESOLVED)


def test_구역이_아예_없으면_수평():
    assert rp._local_rise(OUTSIDE, INSIDE_A, None) == (0.0, ELEV_SOURCE_DEFAULT)


# ── 평면도 헤드망 ─────────────────────────────────────────────────
AV = (0.0, 0.0)
MID = (2000.0, 0.0)
FAR = (4000.0, 0.0)


def _plane_tables(material_zones=None):
    heads = [rp.HeadCandidate(pos=FAR, raw=FAR, block_name="", layer="SP")]
    selection = rp.SelectionResult(
        source_pos=AV, source_kind="manual", heads=heads, distances=[4000.0],
        edges=[(AV, MID, 2000.0), (MID, FAR, 2000.0)],
        nodes_in_subgraph=[AV, MID, FAR],
    )
    return rp.build_input_tables(selection, material_zones=material_zones)


PLANE_ZONE = (-1000.0, -1000.0, 5000.0, 1000.0)


def _zoned(kind):
    return [{"rect": PLANE_ZONE, "kind": kind}]


def test_주차장_헤드는_촛대만큼_올라간다():
    """KS D 3507 상향식 니플은 수직이라 길이가 곧 상승분이다."""
    pipes = _plane_tables(material_zones=_zoned(ZONE_KIND_PARKING)).pipes
    head_pipe = next(p for p in pipes if p["elev"] != 0.0)
    assert head_pipe["elev"] == NIPPLE
    assert {p["elev_source"] for p in pipes} == {ELEV_SOURCE_DEFAULT}


def test_촛대가_평면_길이보다_길면_관_길이를_늘린다():
    """PIPENET 은 |표고차| > 길이 인 관을 거부한다 (피타고라스)."""
    pipes = _plane_tables(material_zones=_zoned(ZONE_KIND_PARKING)).pipes
    assert all(p["length"] >= abs(p["elev"]) for p in pipes)


def test_세대_안_헤드는_수평_낙차는_신축배관_몫():
    pipes = _plane_tables(material_zones=_zoned(ZONE_KIND_UNIT_DWELLING)).pipes
    assert {p["elev"] for p in pipes} == {0.0}
    assert {p["elev_source"] for p in pipes} == {ELEV_SOURCE_DEFAULT}


def test_구역을_안_지정하면_상향_하향을_모른다():
    """상향/하향은 도면에 없다. 구역이 없으면 0 을 쓰되 주장하지 않는다."""
    pipes = _plane_tables().pipes
    sources = [p["elev_source"] for p in pipes]
    assert sources.count(ELEV_SOURCE_UNRESOLVED) == 1   # 헤드가 달린 관로만
    assert {p["elev"] for p in pipes} == {0.0}


def test_평면도_노드_표고는_관례_기본값임을_밝힌다():
    assert {n["elev_source"] for n in _plane_tables().nodes} == {ELEV_SOURCE_DEFAULT}


def test_미확정_관로_수를_meta_에_보고한다():
    meta = dict(_plane_tables().meta)
    assert meta["표고 미확정 관로"] == "1"
    assert ELEV_SOURCE_UNRESOLVED in meta["표고 출처 분포 (관로)"]


# ── 기계실 ────────────────────────────────────────────────────────
MR_PATH = [(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)]


def _machine_room(ceiling_m=None) -> dict:
    return rp._machine_room_path_to_dict(MR_PATH, {}, MR_PATH[0], MR_PATH[-1],
                                         ceiling_m=ceiling_m)


def test_천장고를_주면_첫_구간에만_낙차가_붙는다():
    mr = _machine_room(4.5)
    assert [p["elev"] for p in mr["pipes"]] == [-4.5, 0.0]
    assert mr["pipes"][0]["elev_source"] == ELEV_SOURCE_USER
    assert mr["pipes"][1]["elev_source"] == ELEV_SOURCE_DEFAULT
    assert [n["elevation"] for n in mr["nodes"]] == [0.0, -4.5, -4.5]


def test_천장고가_없으면_0_이_아니라_미확정():
    mr = _machine_room()
    assert mr["pipes"][0]["elev"] == 0.0
    assert mr["pipes"][0]["elev_source"] == ELEV_SOURCE_UNRESOLVED
    assert mr["elev_sources"]["pipes"][ELEV_SOURCE_UNRESOLVED] == 1


def test_낙차가_평면_길이보다_길면_관_길이를_늘린다():
    """PIPENET 은 |표고차| > 길이 인 관을 거부한다 (피타고라스)."""
    short = [(0.0, 0.0), (100.0, 0.0), (3000.0, 0.0)]
    mr = rp._machine_room_path_to_dict(short, {}, short[0], short[-1], ceiling_m=4.5)
    assert mr["pipes"][0]["length"] == pytest.approx(4.5)
