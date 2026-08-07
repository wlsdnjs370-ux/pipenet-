# -*- coding: utf-8 -*-
"""모듈 D — D2 결과 결합.

여기가 통과한다는 것은 "PIPENET 이 계산해 둔 수치를 우리가 같은 값으로 읽는다"
는 뜻이다. 단위를 한 자리라도 잘못 잡으면 도면에도 계산서에도 틀린 수가 찍힌다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core.d_display_model import UNIT_FACTORS  # noqa: E402
from core.d_result_binder import (  # noqa: E402
    COMPOUND_KINDS,
    XML_UNIT_ALIASES,
    _cell,
    bind_results,
    normalize_label,
    read_xdset,
)

SUB = _ROOT / "routes" / "제출용[최종]"
HAND_SDF, HAND_XML = SUB / "2. Pipenet_hand.sdf", SUB / "2. Pipenet_hand.xml"
AUTO_SDF, AUTO_XML = SUB / "2. Pipenet_auto.sdf", SUB / "2. Pipenet_auto.xml"

pytestmark = pytest.mark.skipif(not HAND_XML.exists(), reason="제출용 결과 XML 없음")

_DEFINITIONAL = {f for f, _sym in UNIT_FACTORS.values()}


@pytest.fixture(scope="module")
def hand():
    return bind_results(HAND_SDF, HAND_XML)


@pytest.fixture(scope="module")
def auto():
    return bind_results(AUTO_SDF, AUTO_XML)


def test_normalize_label_drops_xml_suffix():
    assert normalize_label("m1/0") == "M1"
    assert normalize_label(" 52 ") == "52"


def test_cell_leaves_missing_blank():
    # 지시서 7-4 — 없는 값은 추정하지 않고 빈칸이다.
    for raw in ("", "   ", "NaN", "inf", "-inf"):
        assert _cell(raw, "double") is None
    assert _cell("1.5", "double") == 1.5
    assert _cell("0", "boolean") is False


def test_table_rows_are_nested_under_data(hand):
    # TR 은 TABLE 의 직접 자식이 아니다 (TABLE > DATA > TABLEDATA > TR).
    doc = hand.document
    assert doc.table("Pipes-input").rows
    assert doc.table("Pipes-results").rows
    assert doc.table("Node pressures").rows


def test_unit_aliases_all_resolve():
    for token in set(XML_UNIT_ALIASES.values()):
        assert token in UNIT_FACTORS, token


def test_declaration_lies_and_measurement_wins(hand):
    # 두 실물 파일은 SI(Pa) 인데 선언은 코퍼스 전량과 똑같이 'kg/cm2 G' 라고 쓴다.
    press = hand.scales["pressure"]
    assert press.declared_token == "kg/cm2 G"
    assert press.declared_factor == pytest.approx(1.0 / 98066.5)
    assert press.source == "measured"
    assert press.factor == 1.0


def test_flat_network_still_measures_pressure(auto):
    # auto 는 136 관로 중 134 개가 표고 0 이라 압력 앵커 표본이 라이저 하나뿐이다.
    # 여기에 최소 표본을 걸면 거짓말하는 선언으로 물러나 압력이 98066 배 부푼다.
    press = auto.scales["pressure"]
    assert press.samples == 1
    assert press.source == "measured"
    assert press.factor == 1.0


@pytest.mark.parametrize("name", ["hand", "auto"])
def test_adopted_factors_are_definitional(name, request):
    for kind, scale in request.getfixturevalue(name).scales.items():
        if scale.factor is not None:
            assert scale.factor in _DEFINITIONAL, (kind, scale.factor)


def test_compound_units_are_not_scaled(hand):
    for kind in COMPOUND_KINDS & set(hand.scales):
        assert hand.scales[kind].factor is None
    # K-factor 는 파일 표기 그대로 실린다.
    assert hand.nozzles["1"].k_factor == pytest.approx(80.0, abs=1e-5)


def test_node_pressure_is_gauged_to_pipe_results(hand):
    assert hand.datum_offset_pa == pytest.approx(101325.0)
    for pipe in hand.pipes.values():
        node = hand.nodes.get(normalize_label(pipe.input_node))
        if node and node.pressure_pa is not None and pipe.inlet_pressure_pa is not None:
            assert node.pressure_pa == pytest.approx(pipe.inlet_pressure_pa, abs=1.0)


@pytest.mark.parametrize("name", ["hand", "auto"])
def test_pressure_identity_holds(name, request):
    """입구압 − 출구압 = sign(유량)·마찰손실 + 정수두손실.

    마찰손실은 실제 흐름 방향 기준의 양수 크기이고 입·출구는 선언된 노드 순서를
    따르므로, 흐름이 거꾸로 흐르는 관로에서는 부호가 갈린다 (auto 136 관로 중 50 개).
    이 항등식이 성립한다는 것은 압력·유량 배율이 서로 어긋나지 않았다는 뜻이다.
    코퍼스 6699 관로 전량에서 성립하며, 부속이 달린 관로는 펌프가 압력을 더해 제외한다.
    """
    bound = request.getfixturevalue(name)
    equipped = {normalize_label(p.label) for p in bound.model.pipes if p.equipment}
    checked = 0
    for key, pipe in bound.pipes.items():
        values = (pipe.inlet_pressure_pa, pipe.outlet_pressure_pa,
                  pipe.friction_loss_pa, pipe.static_head_loss_pa, pipe.flow_m3s)
        if key in equipped or any(v is None for v in values):
            continue
        signed = values[2] if values[4] >= 0 else -values[2]
        assert values[0] - values[1] == pytest.approx(signed + values[3], abs=2.0)
        checked += 1
    assert checked > 50


def test_hand_join_reports_every_missing_label(hand):
    report = hand.report
    assert not report.matched
    # 승인된 정책 (B): 교집합만 값으로 싣고 결손 라벨은 전량 나열한다.
    assert len(hand.nodes) == 104 and len(hand.pipes) == 103 and len(hand.nozzles) == 30
    assert (report.sdf_only_nodes, report.sdf_only_pipes, report.sdf_only_nozzles) == ((), (), ())
    assert len(report.xml_only_nodes) == 13
    assert len(report.xml_only_pipes) == 13
    assert len(report.xml_only_nozzles) == 7
    assert set(report.xml_only_pipes) & set(hand.pipes) == set()


def test_hand_detects_pipe_that_only_shares_a_name(hand):
    # 라벨은 같은데 SDF 와 XML 의 내용이 다르다 — 결손이 아니라 어긋남이다.
    assert hand.report.divergent_pipes == ("52",)


def test_auto_matches_completely(auto):
    assert auto.report.matched
    assert auto.report.divergent_pipes == ()
    assert len(auto.nodes) == 137 and len(auto.pipes) == 136 and len(auto.nozzles) == 30


def test_values_are_si(hand):
    pipe = hand.pipes["1"]
    assert pipe.nominal_bore_m == pytest.approx(0.15)
    assert pipe.length_m == pytest.approx(20.0)
    assert pipe.velocity_ms == pytest.approx(2.6119676)
    assert hand.nodes["100"].elevation_m == pytest.approx(-75.0, abs=1e-4)
    assert hand.nodes["100"].pressure_pa == pytest.approx(669948.0)


def test_same_unit_name_gets_different_factors_per_table(hand):
    # 같은 unit="diameter" 를 달고도 표마다 표기가 다르다. 호칭경은 mm 로, 실지름은
    # 이미 m 로 적혀 있다 — 한 배율로 뭉뚱그리면 실지름이 1000 배 작아진다.
    # 기대값은 워드 계산서 원문에서 옮겨 적는다 (AVAILABLE PIPE SIZES: 15 → 16.4mm).
    sizes = hand.tables["Design information/Pipe-type/Sizes"]
    actual = {row["Nominal bore"]: row["Actual diameter"] for row in sizes.rows}
    assert actual[15.0] == pytest.approx(0.0164, abs=1e-6)
    assert actual[20.0] == pytest.approx(0.0219, abs=1e-6)
    assert actual[25.0] == pytest.approx(0.0275, abs=1e-6)

    assert hand.tables["Pipes-input/Pipes-input"].rows[0]["Nominal bore"] == pytest.approx(0.15)
    fittings = hand.tables["Materials/Fittings"]
    assert min(r["Fitting Nominal size"] for r in fittings.rows) == pytest.approx(0.025)
    assert not [w for w in hand.warnings if "호칭경으로 못 갈랐다" in w]


def test_title_comes_from_info(hand):
    assert hand.title == ("Officetell", "WATER TANK_PH2F", "3-2 type_MSP_26F")
    assert all("Ã" not in v for v in hand.document.info.values())


def test_read_xdset_does_not_convert(hand):
    # 원문 문서는 파일이 쓴 단위 그대로여야 한다 — 환산은 결합 단계에서만 한다.
    raw = read_xdset(HAND_XML).table("Pipes-input")
    row = next(r for r in raw.rows if normalize_label(r["Label"]) == "1")
    assert row["Nominal bore"] == pytest.approx(150.0)
    assert hand.pipes["1"].nominal_bore_m == pytest.approx(0.15)
