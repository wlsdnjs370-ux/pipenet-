# -*- coding: utf-8 -*-
"""모듈 D — D3 ISO 벡터 PDF 렌더러.

여기가 통과한다는 것은 "PIPENET 없이 만든 도면의 숫자가 결과 XML 원문 그대로"라는
뜻이다. 그래서 검사는 그림을 눈으로 보는 대신 **만든 PDF 에서 글자를 도로 뽑아**
XML 을 직접 파싱한 값과 1:1 로 맞춰 본다. 값이 한 자리라도 어긋나면 여기서 걸린다.
"""
from __future__ import annotations

import hashlib
import re
import sys
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core.d_iso_renderer import (  # noqa: E402
    BAND_COUNT,
    LINK_ITEMS,
    NODE_ITEMS,
    PRESETS,
    band_edges,
    render_iso,
)
from core.d_result_binder import bind_results  # noqa: E402

SUB = _ROOT / "routes" / "제출용[최종]"
HAND_SDF, HAND_XML = SUB / "2. Pipenet_hand.sdf", SUB / "2. Pipenet_hand.xml"
AUTO_SDF, AUTO_XML = SUB / "2. Pipenet_auto.sdf", SUB / "2. Pipenet_auto.xml"

pytestmark = pytest.mark.skipif(not HAND_XML.exists(), reason="제출용 파일 없음")

# 환산 계수를 렌더러에서 import 하지 않고 여기 다시 적는다 — 같은 상수를 돌려
# 쓰면 계수가 틀려도 양쪽이 똑같이 틀려서 검사가 통과해 버린다.
KGF_CM2 = 98066.5
L_MIN = 60000.0
ATM = 101325.0

# 표제란·범례가 놓이는 위쪽 띠(pt). 이 아래가 그림틀이다.
_FRAME_TOP_PT = (1.0 - (12.0 + 16.0 + 18.0) / 297.0) * 297.0 / 25.4 * 72.0


@pytest.fixture(scope="module")
def hand():
    return bind_results(HAND_SDF, HAND_XML)


@pytest.fixture(scope="module")
def auto():
    return bind_results(AUTO_SDF, AUTO_XML)


def _table(xml: Path, name: str) -> list[dict[str, str]]:
    for t in ET.parse(xml).getroot().iter("TABLE"):
        if t.get("name") == name:
            fields = [f.get("name") for f in t.iter("FIELD")]
            return [dict(zip(fields, [(td.text or "").strip() for td in tr]))
                    for tr in t.iter("TR")]
    return []


def _num(s):
    try:
        v = float(s)
    except (TypeError, ValueError):
        return None
    return None if v != v else v


def _key(label: str) -> str:
    return label.split("/")[0].strip().upper()


def _text_items(pdf: Path) -> list[str]:
    from pypdf import PdfReader

    got: list[str] = []
    PdfReader(str(pdf)).pages[0].extract_text(
        visitor_text=lambda t, *a: got.append(t.strip()) if t.strip() else None)
    return got


def _drawn(pdf: Path) -> tuple[dict[str, str], dict[str, str]]:
    """PDF 글자를 관로(라벨 두칸 값)·노드(라벨 한칸 값)로 되읽는다."""
    links, nodes = {}, {}
    for raw in _text_items(pdf):
        if "  " in raw:
            label, _, value = raw.partition("  ")
            links[_key(label)] = value.strip()
        elif " " in raw:
            label, _, value = raw.partition(" ")
            nodes[_key(label)] = value.strip()
    return links, nodes


def _shape_digest(pdf: Path) -> str:
    """그림틀 안 경로 좌표의 지문. 색·글자는 빠진다."""
    from pypdf import PdfReader

    data = PdfReader(str(pdf)).pages[0].get_contents().get_data().decode("latin-1")
    pts = [p for p in re.findall(r"([\d.\-]+) ([\d.\-]+) [mlc]\b", data)
           if float(p[1]) < _FRAME_TOP_PT]
    assert pts, "그림틀 안에 그려진 것이 없다"
    return hashlib.sha256(repr(pts).encode()).hexdigest()


# ── 값 정합 ─────────────────────────────────────────────────────────────────


def test_flow_labels_match_xml(hand, tmp_path):
    pdf = tmp_path / "flow.pdf"
    render_iso(hand, pdf, preset="유량본")
    drawn, _ = _drawn(pdf)
    checked = 0
    for row in _table(HAND_XML, "Pipes-results"):
        got = drawn.get(_key(row["Label"]))
        if got is None:                     # 도면에 형상이 없는 결과 라벨
            continue
        assert got == f"{_num(row['Flow']) * L_MIN:.1f}"
        checked += 1
    assert checked == 103


def test_pressure_labels_match_xml(hand, tmp_path):
    pdf = tmp_path / "press.pdf"
    render_iso(hand, pdf, preset="압력본")
    links, nodes = _drawn(pdf)
    for row in _table(HAND_XML, "Pipes-results"):
        got = links.get(_key(row["Label"]))
        if got is None:
            continue
        drop = _num(row["Inlet pressure"]) - _num(row["Outlet pressure"])
        assert got == f"{drop / KGF_CM2:.1f}"
    checked = 0
    for row in _table(HAND_XML, "Node pressures"):
        got = nodes.get(_key(row["Label"]))
        if got is None:
            continue
        # XDSET 노드 압력은 절대압이다. 게이지로 내려야 관로 결과와 같은 자를 쓴다.
        assert got == f"{(_num(row['Pressure']) - ATM) / KGF_CM2:.1f}"
        checked += 1
    assert checked == 104


def test_bore_labels_match_xml(auto, tmp_path):
    # 호칭경은 mm → m → mm 로 두 번 환산된다. 왕복해도 자리가 밀리면 안 된다.
    pdf = tmp_path / "bore.pdf"
    render_iso(auto, pdf, preset="압력본_옥내소화전")
    links, _ = _drawn(pdf)
    checked = 0
    for row in _table(AUTO_XML, "Pipes-input"):
        got = links.get(_key(row["Label"]))
        if got is None:
            continue
        assert got == f"{_num(row['Nominal bore']):.3f}"
        checked += 1
    assert checked == 136


# ── 수용 기준 (지시서 6) ────────────────────────────────────────────────────


def test_shape_identical_across_presets(hand, tmp_path):
    """표시 항목을 바꿔도 도형은 한 톨도 달라지지 않는다 — 숫자와 색만 바뀐다."""
    digests, texts = set(), []
    for name in PRESETS:
        pdf = tmp_path / f"{name}.pdf"
        render_iso(hand, pdf, preset=name)
        digests.add(_shape_digest(pdf))
        texts.append(_text_items(pdf))
    assert len(digests) == 1
    assert len({tuple(t) for t in texts}) == len(texts)


def test_single_a4_page(hand, tmp_path):
    from pypdf import PdfReader

    pdf = tmp_path / "page.pdf"
    render_iso(hand, pdf, preset="압력본")
    page = PdfReader(str(pdf)).pages[0]
    assert len(PdfReader(str(pdf)).pages) == 1
    width, height = (round(float(v) * 25.4 / 72, 1) for v in page.mediabox[2:])
    assert (width, height) == (210.0, 297.0)
    assert "Page 1 of 1" in _text_items(pdf)


def test_png_resolution_is_fixed(hand, tmp_path):
    # 시각 회귀는 픽셀을 맞대 보는 검사다. 해상도가 흔들리면 비교 자체가 성립하지 않는다.
    from PIL import Image

    png = tmp_path / "vis.png"
    render_iso(hand, png, preset="압력본")
    assert Image.open(png).size == (1240, 1753)


def test_output_is_vector_not_raster(hand, tmp_path):
    from pypdf import PdfReader

    pdf = tmp_path / "vec.pdf"
    render_iso(hand, pdf, preset="압력본")
    page = PdfReader(str(pdf)).pages[0]
    xobjects = page["/Resources"].get("/XObject", {})
    assert not [k for k, v in xobjects.items() if v.get_object().get("/Subtype") == "/Image"]
    assert page["/Resources"].get("/Font")


def test_korean_survives_the_round_trip(hand, tmp_path):
    pdf = tmp_path / "kor.pdf"
    render_iso(hand, pdf, preset="압력본")
    joined = "\n".join(_text_items(pdf))
    assert "FNCADnet 모듈 D 생성" in joined
    assert "계산하거나 보간하지 않는다" in joined
    # 지시서 7-2 — PIPENET 서식을 흉내 내지 않는다.
    assert "PIPENET Schematic" not in joined


# ── 리포트 (지시서 7-4) ─────────────────────────────────────────────────────


def test_report_lists_labels_it_could_not_draw(hand, tmp_path):
    report = render_iso(hand, tmp_path / "r.pdf", preset="압력본")
    assert report.pipes_drawn == 103 and report.pipes_unplaced == ()
    assert report.nozzles_drawn == 30
    # 결과에만 있는 라벨은 조용히 사라지지 않고 전량 실린다.
    assert len(report.undrawn_result_pipes) == 13
    assert len(report.undrawn_result_nodes) == 13
    assert report.blank_link_values == () and report.blank_node_values == ()


def test_unknown_display_item_is_refused(hand, tmp_path):
    with pytest.raises(ValueError, match="모르는 관로 표시 항목"):
        render_iso(hand, tmp_path / "x.pdf", link_item="Pipe colour")


def test_geometry_only_render_without_results(tmp_path):
    # 결과 XML 없이 SDF 만으로도 형상은 나온다. 값이 필요한 항목은 빈칸이 된다.
    from core.d_display_model import load_display_model

    report = render_iso(load_display_model(HAND_SDF), tmp_path / "g.pdf",
                        link_item="Pipe volumetric flow")
    assert report.pipes_drawn == 103
    assert len(report.blank_link_values) == 103
    assert report.link_bands == ()


# ── 분류 ────────────────────────────────────────────────────────────────────


def test_band_edges_are_inside_the_range():
    edges = band_edges([0.0, 1.0, 2.0, 9.7])
    assert len(edges) <= BAND_COUNT - 1
    assert all(0.0 <= e < 9.7 for e in edges)
    assert list(edges) == sorted(edges)


def test_band_edges_empty_when_every_value_is_the_same():
    assert band_edges([3.0, 3.0, 3.0]) == ()
    assert band_edges([]) == ()


def test_presets_only_name_known_items():
    known_link = {i.name for i in LINK_ITEMS}
    known_node = {i.name for i in NODE_ITEMS}
    for preset in PRESETS.values():
        assert preset.link in known_link and preset.node in known_node
