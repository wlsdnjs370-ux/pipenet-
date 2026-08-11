# -*- coding: utf-8 -*-
"""모듈 D — D1 표시 모델 로더.

여기가 통과한다는 것은 "PIPENET 이 그려 둔 화면을 우리가 손실 없이 다시 읽는다"
는 뜻이다. 좌표를 스케일하거나 한글을 깨뜨리면 뒤의 렌더러가 아무리 정확해도
틀린 그림을 그린다.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core.d_display_model import (  # noqa: E402
    UNIT_FACTORS,
    load_display_model,
    restore_cp949,
)

HAND = _ROOT / "routes" / "제출용[최종]" / "2. Pipenet_hand.sdf"
AUTO = _ROOT / "routes" / "제출용[최종]" / "2. Pipenet_auto.sdf"
# 제출용 2 개는 관로·노즐뿐이라 장치 링크를 검증할 수 없다. 펌프·감압밸브·고정손실이
# 한 파일에 다 들어 있는 참조 코퍼스 파일을 쓴다.
_LIB = _ROOT / "data" / "reference_library" / "2. 고가수조_양주옥정 중상1블럭"
DEVICES = (_LIB / "수리계산 원본" / "수리계산 옥내소화전" / "오피스텔" / "116동"
           / "01-1-1. 양주옥정 116동 펌프가압구간 49F_URER.sdf")
PUMP = (_LIB / "수리계산 원본" / "수리계산 옥내소화전" / "오피스텔" / "116동"
        / "01-2. 양주옥정 116동 펌프가압구간 최하층 32F_USER_펌프성능곡선 반영300.sdf")

pytestmark = pytest.mark.skipif(not HAND.exists(), reason="제출용 SDF 없음")


@pytest.fixture(scope="module")
def hand():
    return load_display_model(HAND)


def test_source_display_reads_what_pipenet_saved(hand):
    seen = hand.source_display
    assert (seen.link_labels, seen.node_labels) == (False, False)
    assert (seen.flow_arrows, seen.link_values, seen.node_values) == (True, True, False)
    assert seen.grid == "isometric"


def test_flow_arrows_come_from_the_label_block_not_the_results_block():
    # 코퍼스 334 세트 중 `Results-display/@flow-arrows="0"` 인 13 세트 전량이
    # PIPENET 원본 PDF 에 화살표를 그려 놨다. 그 속성은 화살표 스위치가 아니다.
    model = load_display_model(HAND)
    model.display_options["Results-display"]["flow-arrows"] = "0"
    assert model.source_display.flow_arrows is True
    model.display_options["Label-display"]["arrows"] = "0"
    assert model.source_display.flow_arrows is False


def test_label_all_overrides_the_individual_switches():
    # `label-all='1'` 은 PIPENET 의 "전부 표시" 다. 개별 스위치가 0 이어도 이긴다.
    # 제출용 파일은 둘 다 0 이라 실물로는 확인할 수 없어 여기서 확인한다.
    model = load_display_model(HAND)
    model.display_options["Label-display"]["label-all"] = "1"
    assert (model.source_display.link_labels, model.source_display.node_labels) == (True, True)


def test_missing_display_switch_stays_unknown():
    # 없는 항목을 켬/끔 어느 쪽으로도 지어내지 않는다 — 원본이 무엇을 시켰는지
    # 다시 알 수 없게 된다.
    model = load_display_model(HAND)
    del model.display_options["Label-display"]["link-labels"]
    assert model.source_display.link_labels is None


def test_cp949_leaves_ascii_untouched():
    for s in ("", "PH2. FL -136.0M", "4F  AV", "A/V", "Pipe velocity", "150.0"):
        assert restore_cp949(s) == s


def test_cp949_recovers_korean():
    assert restore_cp949("49Ãþ ÃµÀå") == "49층 천장"
    assert restore_cp949("¿ÁÅ¾1Ãþ") == "옥탑1층"


def test_cp949_keeps_already_correct_korean():
    # 정상 한글은 latin-1 인코딩이 안 되므로 손대지 않고 그대로 나와야 한다.
    assert restore_cp949("49층 천장") == "49층 천장"


def test_at_nodes_are_virtual_and_match_nozzles(hand):
    virtual = {n.label for n in hand.nodes if n.virtual}
    assert len(hand.nodes) == 134
    assert len(virtual) == 30
    assert len(hand.real_nodes) == 104
    # `@` 노드는 노즐의 대기 방출단이다 — 노즐 output 집합과 정확히 같다.
    assert virtual == {z.output_node for z in hand.nozzles}


def test_coordinates_are_not_scaled(hand):
    # parse_sdf 는 평균 절댓값 100 초과면 ÷1000 한다. 표시 경로는 원본을 그대로 쓴다.
    xs = [abs(n.x) for n in hand.nodes]
    assert max(xs) > 1000
    node = next(n for n in hand.nodes if n.label == "10")
    assert node.elevation_m == -75.0


def test_waypoints_preserved(hand):
    with_way = [p for p in hand.pipes if p.waypoints]
    assert len(with_way) == 36
    assert sum(len(p.waypoints) for p in hand.pipes) == 50
    assert all(isinstance(w[0], float) for p in with_way for w in p.waypoints)


def test_equipment_rel_position_preserved(hand):
    eq = [e for p in hand.pipes for e in p.equipment]
    assert len(eq) == 31
    av = next(e for e in eq if e.description == "A/V")
    assert 0.0 <= av.rel_position <= 1.0
    assert av.equivalent_length_m == pytest.approx(12.9)


def test_hand_has_no_device_links(hand):
    assert hand.devices == ()


@pytest.mark.skipif(not DEVICES.exists(), reason="참조 코퍼스 없음")
def test_device_links_are_read_and_connected():
    # 관로가 아닌 링크도 노드를 잇는 구간이다 — 빠뜨리면 도면에서 망이 끊긴다.
    model = load_display_model(DEVICES)
    kinds = sorted(d.kind for d in model.devices)
    assert kinds == ["Elastomeric-valve", "Pressure-loss", "Pressure-loss", "Pressure-loss"]
    labels = {n.label for n in model.nodes}
    for dev in model.devices:
        assert dev.input_node in labels and dev.output_node in labels
    valve = next(d for d in model.devices if d.kind == "Elastomeric-valve")
    assert valve.attributes["target-value"] == "689724"
    assert model.warnings == []


@pytest.mark.skipif(not PUMP.exists(), reason="참조 코퍼스 없음")
def test_pump_carries_library_name():
    pump = next(d for d in load_display_model(PUMP).devices if d.kind == "Pump-fan")
    assert pump.description == "TEST1" and pump.on


def test_graphics_schemes_and_grid(hand):
    assert hand.grid["grid"] == "isometric"
    assert hand.link_scheme == "Pipe velocity"
    assert hand.node_scheme == "None"
    # select-name 과 같은 표기로 후보가 나와야 D3 가 표시값을 고를 수 있다.
    assert "Pipe velocity" in hand.link_schemes
    assert "Pipe bore" in hand.link_schemes
    assert "Node pressure" in hand.node_schemes


def test_text_elements_carry_position_and_korean(hand):
    assert len(hand.texts) == 10
    t = hand.texts[0]
    assert (t.x, t.y) == (-11650.0, -766.0)
    assert t.typeface == "Arial" and t.typesize == 60.0
    assert all("Ã" not in x.text for x in hand.texts)


def test_units_read_from_file_not_hardcoded(hand):
    press = hand.units["Pressure"]
    assert press.token == "kg-f-cm2-g"
    assert press.precision == 1
    # 771273 Pa == 7.9 kg/cm2 (98066.5 Pa/kgf-cm2)
    assert press.format(771273.0) == "7.9"
    assert hand.units["Volumetric-flow"].format(1.0 / 60000.0) == "1.0"
    assert hand.units["Diameter"].format(0.15) == "150.000"


def test_missing_value_renders_blank_not_guessed(hand):
    assert hand.units["Pressure"].format(None) == ""


def test_unit_factors_are_definitional():
    assert UNIT_FACTORS["kg-f-cm2-g"][0] == pytest.approx(1.0 / 98066.5)
    assert UNIT_FACTORS["l-min"][0] == 60000.0


def test_no_warnings_on_real_files(hand):
    assert hand.warnings == []
    assert load_display_model(AUTO).warnings == []


def test_bounds_cover_waypoints_and_texts(hand):
    minx, miny, maxx, maxy = hand.bounds()
    for p in hand.pipes:
        for x, y in p.waypoints:
            assert minx <= x <= maxx and miny <= y <= maxy
    for t in hand.texts:
        assert minx <= t.x <= maxx and miny <= t.y <= maxy


def test_parse_sdf_contract_unchanged():
    # 수리 경로는 이 작업으로 바뀌지 않는다 — 좌표는 여전히 m 로 축소되고
    # `@` 노드는 여전히 빠진다. 모듈 A/B 가 이 계약에 의존한다.
    from kfp_sdf_converter import parse_sdf

    net = parse_sdf(HAND)
    assert not any("@" in nid for nid in net.nodes)

    disp = {n.label: n for n in load_display_model(HAND).nodes}
    checked = 0
    for nid, hyd in net.nodes.items():
        shown = disp.get(nid.lstrip("N"))
        if shown is None:
            continue
        assert shown.x == pytest.approx(hyd.x * 1000)
        assert shown.y == pytest.approx(hyd.y * 1000)
        checked += 1
    assert checked == 104
