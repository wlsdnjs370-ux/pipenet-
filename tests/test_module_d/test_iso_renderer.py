# -*- coding: utf-8 -*-
"""모듈 D — D3 ISO 벡터 PDF 렌더러.

여기가 통과한다는 것은 "PIPENET 없이 만든 도면의 숫자가 결과 XML 원문 그대로"라는
뜻이다. 그래서 검사는 그림을 눈으로 보는 대신 **만든 PDF 에서 글자를 도로 뽑아**
XML 을 직접 파싱한 값과 1:1 로 맞춰 본다. 값이 한 자리라도 어긋나면 여기서 걸린다.
"""
from __future__ import annotations

import collections
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

# 그림틀이 종이에서 차지하는 세로 구간. 표제란과 범례는 그 아래에 함께 있다.
_A4_H_PT = 297.0 / 25.4 * 72.0
_FRAME_BOTTOM_PT = 131.33               # 종이 아래에서
_FRAME_TOP_PT = _A4_H_PT - 58.93        # 종이 위에서 58.93


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


# 스트림은 줄 길이에 맞춰 아무 데서나 접힌다 — 피연산자 사이를 빈칸 하나로 못 박으면
# 획 굵기 하나만 바뀌어도 색 연산자가 두 줄로 갈라져 안 보인다.
_STREAM_OP = re.compile(
    r"(?P<grey>[\d.]+)\s+G\b|(?P<rgb>(?:[\d.]+\s+){3})RG\b|(?P<stack>(?<![\w.])[qQ](?![\w.]))"
    r"|(?P<x>[\d.\-]+)\s+(?P<y>[\d.\-]+)\s+[mlc]\b")


def _shape_digest(pdf: Path) -> str:
    """그림틀 안 경로 좌표의 지문. 색·글자는 빠진다.

    지시선은 도형이 아니라 라벨의 부속이다 — 라벨이 어디로 밀려났는지에 따라
    프리셋마다 달라지므로 지문에서 뺀다. 회색 획으로만 그려지니 획 색을 따라가며
    거른다. 망이 지시선과 같은 색으로 바뀌면 이 필터가 망까지 지워 지문이 비고,
    아래 단언이 먼저 걸린다.
    """
    from pypdf import PdfReader

    from core.d_iso_renderer import _LEADER_COLOUR

    leader_grey = int(_LEADER_COLOUR[1:3], 16) / 255.0
    data = PdfReader(str(pdf)).pages[0].get_contents().get_data().decode("latin-1")
    pts, grey = [], None
    for m in _STREAM_OP.finditer(data):
        if m["x"] is None:
            grey = float(m["grey"]) if m["grey"] else None
            continue
        if grey is not None and abs(grey - leader_grey) < 1e-6:
            continue
        if _FRAME_BOTTOM_PT < float(m["y"]) < _FRAME_TOP_PT:
            pts.append((m["x"], m["y"]))
    assert pts, "그림틀 안에 그려진 것이 없다"
    return hashlib.sha256(repr(pts).encode()).hexdigest()


def _stroked_paths(pdf: Path) -> list[tuple[object, list[tuple[float, float]]]]:
    """획 색깔별 경로 좌표(pt). 회색 획은 float, 색깔 획은 3튜플로 돌려준다."""
    from pypdf import PdfReader

    data = PdfReader(str(pdf)).pages[0].get_contents().get_data().decode("latin-1")
    out: list[tuple[object, list[tuple[float, float]]]] = []
    colour: object = None
    started: object = None
    cur: list[tuple[float, float]] = []
    for m in _STREAM_OP.finditer(data):
        if m["x"] is None:
            if m["grey"]:
                colour = float(m["grey"])
            elif m["rgb"]:
                colour = tuple(round(float(v), 4) for v in m["rgb"].split())
            continue
        point = (float(m["x"]), float(m["y"]))
        if data[m.end() - 1] == "m":
            if cur:
                out.append((started, cur))
            cur, started = [point], colour
        else:
            cur.append(point)
    if cur:
        out.append((started, cur))
    return out


_WIDTH_OP = re.compile(
    r"(?P<grey>[\d.]+)\s+G\b|(?P<rgb>(?:[\d.]+\s+){3})RG\b|(?P<w>[\d.]+)\s+w\b"
    r"|(?P<x>[\d.\-]+)\s+(?P<y>[\d.\-]+)\s+[mlc]\b")


def _stroked_widths(pdf: Path) -> list[tuple[float, object, int]]:
    """획마다 (굵기 pt, 색, 꼭짓점 수). _stroked_paths 와 달리 굵기를 함께 본다."""
    from pypdf import PdfReader

    data = PdfReader(str(pdf)).pages[0].get_contents().get_data().decode("latin-1")
    out: list[tuple[float, object, int]] = []
    colour: object = None
    width = 1.0
    started: object = None
    started_w, count = 1.0, 0
    for m in _WIDTH_OP.finditer(data):
        if m["x"] is None:
            if m["grey"]:
                colour = float(m["grey"])
            elif m["rgb"]:
                colour = tuple(round(float(v), 4) for v in m["rgb"].split())
            else:
                width = float(m["w"])
            continue
        if data[m.end() - 1] == "m":
            if count:
                out.append((started_w, started, count))
            started, started_w, count = colour, width, 1
        else:
            count += 1
    if count:
        out.append((started_w, started, count))
    return out


_HAIRLINE_PT = 0.06
_RULE_OP = re.compile(
    r"(?P<w>[\d.]+)\s+w\b"
    r"|(?P<x0>[\d.\-]+)\s+(?P<y0>[\d.\-]+)\s+m\s+(?P<x1>[\d.\-]+)\s+(?P<y1>[\d.\-]+)\s+l\s+S\b")


def _rules(pdf: Path) -> tuple[list[tuple[float, float, float]], list[tuple[float, float, float]]]:
    """판면 괘선(0.06pt 직선)을 가로/세로로 갈라 (고정좌표, 시작, 끝) 로 돌려준다."""
    from pypdf import PdfReader

    data = PdfReader(str(pdf)).pages[0].get_contents().get_data().decode("latin-1")
    horizontal: list[tuple[float, float, float]] = []
    vertical: list[tuple[float, float, float]] = []
    width = 1.0
    for m in _RULE_OP.finditer(data):
        if m["w"]:
            width = float(m["w"])
        elif width == _HAIRLINE_PT:
            x0, y0, x1, y1 = (float(m[k]) for k in ("x0", "y0", "x1", "y1"))
            if y0 == y1:
                horizontal.append((y0, min(x0, x1), max(x0, x1)))
            else:
                vertical.append((x0, min(y0, y1), max(y0, y1)))
    return horizontal, vertical


_FILL_OP = re.compile(
    r"(?P<rgb>(?:[\d.]+\s+){3})rg\b|(?P<grey>[\d.]+)\s+g\b"
    r"|(?P<x>[\d.\-]+)\s+(?P<y>[\d.\-]+)\s+m\b|(?P<paint>[fS])\b")


def _legend_fills(pdf: Path) -> list[tuple[float, ...]]:
    """범례 칸을 읽는 순서대로 — 윗줄 왼쪽에서 아랫줄 오른쪽으로 — 채움색만 돌려준다."""
    from pypdf import PdfReader

    data = PdfReader(str(pdf)).pages[0].get_contents().get_data().decode("latin-1")
    found: list[tuple[float, float, tuple[float, ...]]] = []
    colour: tuple[float, ...] | None = None
    start: tuple[float, float] | None = None
    for m in _FILL_OP.finditer(data):
        if m["rgb"]:
            colour = tuple(round(float(v), 4) for v in m["rgb"].split())
        elif m["grey"]:
            # 세 성분이 같은 색은 rg 가 아니라 g 한 값으로 적힌다. 이걸 빠뜨리면
            # 무채색 칸이 앞 칸 색을 그대로 물려받아 검사가 거짓으로 통과한다.
            colour = (round(float(m["grey"]), 4),) * 3
        elif m["paint"]:
            # 표제란 괘선도 같은 자리를 지나지만 채우지 않고 S 로 긋고 끝난다.
            if m["paint"] == "f" and colour is not None and start is not None:
                found.append((round(start[1], 1), start[0], colour))
            start = None
        # y 0 은 종이 바탕이다 — 범례 칸이 아니다.
        elif 0.0 < float(m["y"]) < _FRAME_BOTTOM_PT:
            start = (float(m["x"]), float(m["y"]))
    # 칸마다 경로가 하나씩이다. 줄이 여럿이므로 x 만 보면 두 줄이 섞인다.
    return [c for _, _, c in sorted(found, key=lambda f: (-f[0], f[1]))]


def _data_to_pt(model):
    """데이터 좌표 → 페이지 pt. 종이·여백 수치를 렌더러에서 가져오지 않고 다시 적는다."""
    minx, miny, maxx, maxy = model.bounds()
    span_x, span_y = maxx - minx, maxy - miny
    page = (297.0, 210.0) if span_x / span_y > 297.0 / 210.0 else (210.0, 297.0)
    mm = 72.0 / 25.4
    box_w = (page[0] - 24.0) * mm
    box_h = _FRAME_TOP_PT - _FRAME_BOTTOM_PT
    box_cx = page[0] * mm / 2.0
    box_cy = (_FRAME_TOP_PT + _FRAME_BOTTOM_PT) / 2.0
    pad = 0.04 * max(span_x, span_y)
    # 축은 aspect equal 이라 가로세로 같은 배율로 줄고 칸 가운데에 놓인다.
    scale = min(box_w / (span_x + 2 * pad), box_h / (span_y + 2 * pad))
    cx, cy = (minx + maxx) / 2.0, (miny + maxy) / 2.0
    return lambda x, y: (box_cx + scale * (x - cx), box_cy + scale * (y - cy))


# ── 값 정합 ─────────────────────────────────────────────────────────────────
#
# 아래 세 검사는 이름표를 켠 채로 그린다. 기본값은 원본 SDF 를 따르는데 제출용
# 두 파일 모두 이름표가 꺼져 있어, 그대로 두면 PDF 에 값만 남아 어느 관로의
# 값인지 짝지을 수 없다. 이름표는 값을 식별하기 위한 검사 도구다.
_TAGGED = {"show_link_labels": True, "show_node_labels": True}


def test_flow_labels_match_xml(hand, tmp_path):
    pdf = tmp_path / "flow.pdf"
    render_iso(hand, pdf, preset="유량본", **_TAGGED)
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
    render_iso(hand, pdf, preset="압력본", **_TAGGED)
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
    render_iso(auto, pdf, preset="압력본_옥내소화전", **_TAGGED)
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


def _glyph_boxes(pdf: Path) -> list[tuple[str, float, float, float, float]]:
    """그림틀 안 글자가 실제로 차지한 자리(pt).

    배치기가 쓴 상자를 돌려 쓰지 않는다 — 자리는 PDF 텍스트 행렬에서, 크기는 글꼴
    윤곽에서 가져온다. 배치기가 글자를 작게 재고 있었다면 여기서 걸린다.
    """
    from matplotlib.textpath import TextPath
    from pypdf import PdfReader

    from core.d_iso_renderer import _korean_font

    prop = _korean_font()
    extents: dict[tuple[str, float], object] = {}
    got: list[tuple[str, float, float, float, float]] = []

    def visit(text, cm, tm, font_dict, font_size):
        s = text.strip()
        if not s:
            return
        # 회전과 위치는 tm 이 아니라 cm 에 실려 나온다 — tm 은 단위행렬 그대로다.
        a, b, c, d, e, f = (float(v) for v in cm)
        if not _FRAME_BOTTOM_PT < f < _FRAME_TOP_PT:
            return
        key = (s, font_size)
        ext = extents.get(key)
        if ext is None:
            ext = extents[key] = TextPath((0, 0), s, size=font_size, prop=prop).get_extents()
        # 중심을 되짚지 않는다. 윤곽의 네 귀퉁이를 PDF 행렬로 그대로 옮긴다 — 기준점이
        # 글자 중심인지 밑선인지 짐작할 필요가 없어진다.
        pts = [(a * px + c * py + e, b * px + d * py + f)
               for px, py in ((ext.x0, ext.y0), (ext.x1, ext.y0),
                              (ext.x1, ext.y1), (ext.x0, ext.y1))]
        xs, ys = [p[0] for p in pts], [p[1] for p in pts]
        got.append((s, min(xs), min(ys), max(xs), max(ys)))

    PdfReader(str(pdf)).pages[0].extract_text(visitor_text=visit)
    return got


def _overlaps(boxes) -> list[tuple[str, str]]:
    boxes = sorted(boxes, key=lambda b: b[1])
    bad = []
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            if b[1] >= a[3]:
                break
            if a[2] < b[4] and b[2] < a[4]:
                bad.append((a[0], b[0]))
    return bad


@pytest.mark.parametrize("preset", list(PRESETS))
def test_no_label_overlaps_in_the_drawing(hand, auto, preset, tmp_path):
    """첨부 실물 파일에서 라벨 겹침 0건 (지시서 D4 수용 기준)."""
    for name, model in (("hand", hand), ("auto", auto)):
        pdf = tmp_path / f"{name}_{preset}.pdf"
        report = render_iso(model, pdf, preset=preset)
        boxes = _glyph_boxes(pdf)
        assert len(boxes) > 100
        assert _overlaps(boxes) == []
        # 자리를 못 찾았다고 조용히 빼지 않는다 — 뺐다면 겹침 0 은 공짜로 나온다.
        assert report.labels_dropped == ()
        assert report.label_seconds < 3.0


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


# ── 헤드 배치 ───────────────────────────────────────────────────────────────

# 스텁 길이. 렌더러에서 import 하지 않고 다시 적는다 — 같은 상수를 돌려 쓰면 값이
# 틀려도 양쪽이 똑같이 틀려서 검사가 통과한다. 도면 크기가 아니라 모델 단위에 붙어
# 있고(변동계수 0.218 대 0.800), 이 값은 참조 코퍼스 노즐 3437 개의 최빈값이다.
_STUB_UNITS = 58.0

# 헤드 삼각형. 참조 PDF 60 장 / 기호 638 개 실측이다 — 꼭짓점이 `@` 노드에 앉고
# 밑변은 그로부터 17.96 모델 단위 뒤, 반폭 10.01 단위다. 스텁은 꼭짓점이 아니라
# 이 밑변에서 끝난다.
_HEAD_LENGTH = 17.96
_HEAD_HALF_WIDTH = 10.01
# 노드 점 지름. 같은 60 장 / 점 3524 개에서 모델 단위 변동계수 0.019, 종이 pt 0.420.
_NODE_DOT = 9.59
# 관로 획 굵기. 참조 80 장에서 모델 단위 변동계수 0.006, 종이 pt 0.405.
_PIPE_WIDTH = 1.0


def _head_back(base, tip):
    """스텁이 끝나는 자리 — 삼각형 밑변 가운데. 렌더러 헬퍼를 쓰지 않고 다시 구한다."""
    import math

    reach = math.dist(base, tip)
    length = min(_HEAD_LENGTH, reach)
    return (tip[0] - (tip[0] - base[0]) / reach * length,
            tip[1] - (tip[1] - base[1]) / reach * length)


def _head_tips(bound):
    import math
    import statistics
    from types import SimpleNamespace

    from core.d_iso_renderer import _nozzle_tips

    model = bound.model
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    rows = {z.label: SimpleNamespace(nozzle=z) for z in model.nozzles}
    # 유도 길이는 같은 도면의 방향 있는 헤드에서 온다. 없으면 코퍼스 최빈값이다.
    seen = [math.dist(coords[z.input_node], coords[z.output_node])
            for z in model.nozzles
            if z.input_node in coords and z.output_node in coords
            and coords[z.input_node] != coords[z.output_node]]
    stub = statistics.median(seen) if seen else _STUB_UNITS
    tips, derived, undirected = _nozzle_tips(rows, model.pipes, coords, stub)
    return tips, derived, undirected, stub, coords, model


def _outgoing_dirs(model, coords, base):
    """base 에서 관로가 뻗어 나가는 방향. 렌더러 헬퍼를 쓰지 않고 다시 구한다."""
    import math

    origin = coords[base]
    out = []
    for pipe in model.pipes:
        if pipe.input_node == base:
            ahead = [*pipe.waypoints, coords.get(pipe.output_node)]
        elif pipe.output_node == base:
            ahead = [*reversed(pipe.waypoints), coords.get(pipe.input_node)]
        else:
            continue
        for point in ahead:
            if point is None:
                continue
            dx, dy = point[0] - origin[0], point[1] - origin[1]
            dist = math.hypot(dx, dy)
            if dist > 1e-9:
                out.append((dx / dist, dy / dist))
                break
    return out


def test_head_positions_given_by_the_file_are_left_alone(hand):
    # 수작업본은 @/n 에 진짜 좌표가 있다. 원본이 정한 자리를 옮기지 않는다.
    tips, derived, undirected, _, coords, model = _head_tips(hand)
    assert derived == [] and undirected == []
    for z in model.nozzles:
        assert tips[z.label] == coords[z.output_node]


def test_heads_without_a_direction_are_derived_and_declared(auto, tmp_path):
    # 자동 SDF 는 @/n 좌표가 입력노드와 겹쳐 방향이 없다. 원본은 고칠 수 없으므로
    # (지시서 7-3) 그릴 때 유도하되, 유도했다는 사실을 리포트가 전량 실어야 한다.
    import math

    tips, derived, undirected, stub, coords, model = _head_tips(auto)
    assert len(derived) == 30 and undirected == []
    # 자동 SDF 는 방향 있는 헤드가 하나도 없다 — 잴 것이 없으니 코퍼스 값으로 떨어진다.
    assert stub == _STUB_UNITS

    for z in model.nozzles:
        base, tip = coords[z.input_node], tips[z.label]
        assert coords[z.output_node] == base       # 원본이 방향을 주지 않았다는 전제
        assert math.dist(base, tip) == pytest.approx(_STUB_UNITS, rel=1e-9)
        # 입사 관로의 반대쪽 — 망 바깥으로 뻗는다 (코퍼스 3115 개 중 98.8% 의 규칙).
        dirs = _outgoing_dirs(model, coords, z.input_node)
        assert dirs
        sx = sum(d[0] for d in dirs) / len(dirs)
        sy = sum(d[1] for d in dirs) / len(dirs)
        mag = math.hypot(sx, sy)
        ux, uy = (tip[0] - base[0]) / math.dist(base, tip), (tip[1] - base[1]) / math.dist(base, tip)
        assert (sx / mag) * ux + (sy / mag) * uy == pytest.approx(-1.0, abs=1e-9)

    # 헤드가 서로 다른 자리에 놓인다 — 겹쳐 찍히던 것이 이 검사에서 걸린다.
    assert len(set(tips.values())) == 30

    report = render_iso(auto, tmp_path / "heads.pdf", link_item="Pipe volumetric flow")
    assert len(report.nozzles_derived) == 30 and report.nozzles_undirected == ()
    assert any("유도해 그린 것 30개" in w for w in report.warnings)


def test_a_derived_head_borrows_the_length_the_drawing_already_uses():
    # 유도 길이는 도면마다 다르다 — 참조 코퍼스에서 값이 26 종뿐인데 55~62 와 86~88
    # 두 무리로 갈린다. 같은 도면에 이미 잰 길이가 있으면 코퍼스 최빈값보다 그쪽이 옳다.
    import dataclasses
    import math
    import statistics

    bound = bind_results(HAND_SDF, HAND_XML)      # 모듈 fixture 를 건드리지 않는다
    model = bound.model
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    blind, *rest = model.nozzles
    # 헤드 하나만 방향을 지운다 — 나머지가 길이의 근거로 남는 '섞인 도면'이 된다.
    model.nozzles = (dataclasses.replace(blind, output_node=blind.input_node), *rest)
    theirs = statistics.median(
        [math.dist(coords[z.input_node], coords[z.output_node]) for z in rest])
    assert theirs != _STUB_UNITS                  # 두 경로가 갈리는 표본이어야 의미가 있다

    tips, derived, undirected, stub, coords, _ = _head_tips(bound)
    assert derived == [blind.label] and undirected == []
    assert stub == theirs
    assert math.dist(coords[blind.input_node], tips[blind.label]) == pytest.approx(theirs)


def _stub_of(pdf: Path, colour: object) -> list[list[tuple[float, float]]]:
    # 회색조는 스트림에 자릿수가 잘린 십진수로 적힌다 (0x22/255 → "0.1333333333").
    def same(got: object) -> bool:
        if isinstance(got, float) and isinstance(colour, float):
            return abs(got - colour) < 5e-4
        return got == colour

    return [pts for got, pts in _stroked_paths(pdf) if same(got)]


def _same_segment(a, b) -> bool:
    return all(abs(p[0] - q[0]) < 0.5 and abs(p[1] - q[1]) < 0.5 for p, q in zip(a, b))


def test_every_head_is_joined_to_the_pipe_it_hangs_from(hand, tmp_path):
    # 노즐도 입력노드와 출력노드를 잇는 링크다. 그 선을 그리지 않으면 헤드가
    # 배관에서 떨어져 떠 있는 것처럼 보인다. PDF 안에 선분이 실제로 있는지 본다.
    pdf = tmp_path / "stubs.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity")   # 관로는 띠 색이라 회색과 갈린다

    tips, derived, undirected, _, coords, model = _head_tips(hand)
    assert derived == [] and undirected == []          # 수작업본은 전부 실측 자리
    to_pt = _data_to_pt(model)
    drawn = _stub_of(pdf, 0x33 / 255.0)
    assert len(drawn) == len(model.nozzles) == 30
    assert all(len(pts) == 2 for pts in drawn)

    for z in model.nozzles:
        base = coords[z.input_node]
        want = [to_pt(*base), to_pt(*_head_back(base, tips[z.label]))]
        assert any(_same_segment(want, got) or _same_segment(want, got[::-1])
                   for got in drawn), f"헤드 {z.label} 가 배관에 닿는 선이 없다"


def test_derived_head_lines_are_not_dressed_up_as_measured(auto, tmp_path):
    # 자동본은 30개 전부 방향을 유도한 자리다. 지어낸 선을 실측 선과 같은 색·같은
    # 실선으로 그리면 도면을 보는 사람이 둘을 구분할 길이 없다.
    pdf = tmp_path / "derived.pdf"
    render_iso(auto, pdf, link_item="Pipe velocity")
    assert _stub_of(pdf, 0x33 / 255.0) == []          # 실측 실선으로 새어 나간 것 없음

    orange = _stub_of(pdf, (0.7608, 0.2549, 0.0471))  # #c2410c
    stubs = [pts for pts in orange if len(pts) == 2]
    assert len(stubs) == 30
    assert len(orange) - len(stubs) == 30             # 헤드 표시도 같은 색 테두리

    tips, derived, _, _, coords, model = _head_tips(auto)
    assert len(derived) == 30
    to_pt = _data_to_pt(model)
    for z in model.nozzles:
        base = coords[z.input_node]
        want = [to_pt(*base), to_pt(*_head_back(base, tips[z.label]))]
        assert any(_same_segment(want, got) or _same_segment(want, got[::-1])
                   for got in stubs), f"유도한 헤드 {z.label} 의 선이 없다"


def test_head_triangle_is_an_outline_with_its_apex_on_the_at_node(hand, tmp_path):
    # PIPENET 은 채운 삼각형이 아니라 속 빈 윤곽을 그리고, 꼭짓점을 `@` 노드 좌표에
    # 그대로 앉힌다 — 참조 기호 638 개 중 채움 0 개, 그린 길이 ÷ 모델 길이 1.0004.
    import math

    pdf = tmp_path / "heads.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity")

    tips, derived, _, _, coords, model = _head_tips(hand)
    assert derived == []
    to_pt = _data_to_pt(model)
    # 윤곽선은 꼭짓점으로 돌아와 닫힌다 — 채웠다면 획이 아니라 채움으로 나온다.
    triangles = [pts for pts in _stub_of(pdf, 0x22 / 255.0)
                 if len(pts) == 4 and pts[0] == pts[3]]
    assert len(triangles) == len(model.nozzles) == 30

    for z in model.nozzles:
        base, tip = coords[z.input_node], tips[z.label]
        back = _head_back(base, tip)
        reach = math.dist(base, tip)
        half = _HEAD_HALF_WIDTH * min(_HEAD_LENGTH, reach) / _HEAD_LENGTH
        ux, uy = (tip[0] - base[0]) / reach, (tip[1] - base[1]) / reach
        apex = to_pt(*tip)
        want = [apex,
                to_pt(back[0] - uy * half, back[1] + ux * half),
                to_pt(back[0] + uy * half, back[1] - ux * half),
                apex]
        assert any(_same_segment(want, got) for got in triangles), \
            f"헤드 {z.label} 의 삼각형 윤곽이 없다"


def test_no_equipment_symbol_is_shaped_like_a_head(hand, tmp_path):
    # 헤드가 속 빈 삼각형이 된 뒤로 삼각형은 노즐 헤드만의 표시여야 한다. 신축배관
    # 기기가 2.6pt 짜리 검은 삼각이라 3.6pt 짜리 헤드와 구별되지 않았다.
    from pypdf import PdfReader

    pdf = tmp_path / "symbols.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity")

    forms = (PdfReader(str(pdf)).pages[0]["/Resources"].get("/XObject") or {})
    assert forms, "기호를 재사용 도형으로 찍지 않으면 이 검사가 아무것도 못 본다"
    for name, ref in forms.items():
        data = ref.get_object().get_data().decode("latin-1")
        corners = [ln for ln in data.splitlines() if ln.endswith((" m", " l", " c"))]
        assert len(corners) != 3, f"{name} 이 삼각형이라 노즐 헤드와 겹친다"


def test_node_dots_are_sized_in_model_units(hand, tmp_path):
    # 점 크기도 종이가 아니라 모델 좌표에 붙어 있다 — 도면이 커지면 같이 커진다.
    pdf = tmp_path / "dots.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity")

    to_pt = _data_to_pt(hand.model)
    want = _NODE_DOT * (to_pt(1.0, 0.0)[0] - to_pt(0.0, 0.0)[0])
    # 점은 베지어 여덟 도막으로 닫힌다 — 도면의 다른 경로와 꼭짓점 수가 겹치지 않는다.
    dots = [pts for _, pts in _stroked_paths(pdf) if len(pts) == 9 and pts[0] == pts[8]]
    assert len(dots) == len(hand.model.real_nodes) == 104
    for pts in dots:
        wide = max(x for x, _ in pts) - min(x for x, _ in pts)
        assert wide == pytest.approx(want, abs=1e-3)


def test_pipe_strokes_are_one_width_fixed_in_model_units(hand, tmp_path):
    # PIPENET 은 선 굵기로 관경을 나타내지 않는다 — 참조 80 장이 관경을 6~8 종류씩
    # 쓰면서도 획 굵기는 한 장에 한 종류다. 그 굵기도 종이가 아니라 모델 좌표에 있다.
    pdf = tmp_path / "widths.pdf"
    render_iso(hand, pdf)

    bores = {p.bore_m for p in hand.model.pipes if p.bore_m}
    assert len(bores) > 1, "관경이 한 종류면 굵기가 관경을 따라가는지 알 수 없다"

    to_pt = _data_to_pt(hand.model)
    want = _PIPE_WIDTH * (to_pt(1.0, 0.0)[0] - to_pt(0.0, 0.0)[0])
    pipe_grey = 0x33 / 255.0
    widths = {w for w, colour, count in _stroked_widths(pdf)
              if isinstance(colour, float) and abs(colour - pipe_grey) < 1e-6 and count > 1}
    assert widths, "관로 획을 찾지 못했다"
    assert len(widths) == 1, f"한 도면에 굵기가 여러 가지다: {sorted(widths)}"
    assert widths.pop() == pytest.approx(want, abs=1e-3)


def test_leader_lines_are_thinner_than_pipes(hand, tmp_path):
    # 지시선이 관로로 오독되면 안 된다. 종이 pt 로 묶여 있던 0.25pt 는 참조 배율
    # 폭(0.089~0.167 pt/단위) 어디에서도 배관보다 1.5~2.8 배 굵었다 — 굵기를 배관과
    # 같은 자(모델 단위)에 태워야 도면 크기와 무관하게 가늘다.
    pdf = tmp_path / "leaders.pdf"
    # 두 항목을 함께 켜야 라벨이 붐벼 지시선이 나온다.
    render_iso(hand, pdf, link_item="Pipe velocity", node_item="Node pressure")

    to_pt = _data_to_pt(hand.model)
    per_unit = to_pt(1.0, 0.0)[0] - to_pt(0.0, 0.0)[0]
    strokes = _stroked_widths(pdf)
    leader_grey = 0x7f / 255.0
    leaders = {w for w, colour, _ in strokes
               if isinstance(colour, float) and abs(colour - leader_grey) < 2e-3}
    assert leaders, "지시선을 찾지 못했다"
    assert max(leaders) < _PIPE_WIDTH * per_unit, \
        f"지시선이 배관({_PIPE_WIDTH * per_unit:.4f}pt)보다 굵다: {sorted(leaders)}"


def test_our_symbols_scale_with_the_network_not_the_page(hand, tmp_path):
    # 기기 기호는 지시서 7-2 에 따라 우리 표기라 PIPENET 실측이 없다. 그래도 크기를
    # 매다는 기준은 망과 같아야 한다 — 종이 pt 로 묶으면 같은 기호가 참조 배율 폭에서
    # 노드 점의 2.1 배부터 4.0 배까지 널뛴다. 모델을 통째로 키워 배율만 바꿔 본다.
    import copy
    import dataclasses

    pdf_a, pdf_b = tmp_path / "a.pdf", tmp_path / "b.pdf"
    render_iso(hand, pdf_a, link_item="Pipe velocity")

    k = 3.0
    big = copy.deepcopy(hand)
    m = big.model
    m.nodes = tuple(dataclasses.replace(n, x=n.x * k, y=n.y * k) for n in m.nodes)
    m.pipes = tuple(dataclasses.replace(
        p, waypoints=tuple((x * k, y * k) for x, y in p.waypoints)) for p in m.pipes)
    m.texts = tuple(dataclasses.replace(t, x=t.x * k, y=t.y * k) for t in m.texts)
    render_iso(big, pdf_b, link_item="Pipe velocity")

    # 좌표를 3 배로 늘리면 종이가 그대로이므로 pt/단위 배율은 1/3 이 된다. 모델에
    # 매달린 굵기는 함께 1/3 이 되고, 종이에 매달린 굵기는 꿈쩍도 하지 않는다.
    # 기기 테두리가 가장 굵은 획이라 이 값이 우리 기호를 대표한다.
    def widest(pdf):
        return max(w for w, colour, _ in _stroked_widths(pdf) if isinstance(colour, float))

    a, b = widest(pdf_a), widest(pdf_b)
    assert b == pytest.approx(a / k, rel=0.02), \
        f"배율이 {k:g} 배 달라졌는데 굵기가 {a:.4f} → {b:.4f} 로 따라오지 않았다"


def test_band_colours_are_the_ones_pipenet_uses(hand, tmp_path):
    # 참조 40 장 중 범례가 있는 35 장이 전부 이 여섯 색을 이 순서로 쓴다. 렌더러
    # 상수를 가져오지 않고 실측값을 다시 적는다.
    want = [(1.0, 0.0, 0.0), (1.0, 0x_ac / 255.0, 0.0), (0.0, 1.0, 0.0),
            (0.0, 1.0, 1.0), (0.0, 0.0, 1.0), (1.0, 0.0, 1.0)]
    want = [tuple(round(v, 4) for v in c) for c in want]

    pdf = tmp_path / "bands.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity")
    drawn = {colour for colour, _ in _stroked_paths(pdf) if isinstance(colour, tuple)}
    assert drawn, "띠 색으로 그린 획이 없다"
    assert drawn <= set(want), f"참조에 없는 색을 쓴다: {sorted(drawn - set(want))}"
    # 색만 맞고 순서가 뒤집히면 도면이 거짓말한다. 범례 칸을 왼쪽부터 읽어 확인한다.
    assert _legend_fills(pdf) == want


_LABEL_UNITS = 24.17     # 값 글씨 높이(모델 단위). 변동계수 0.029
_NOTE_UNITS = 22.56      # 도면 주기 높이(모델 단위). 변동계수 0.011
_NOTE_TYPESIZE = 30.0    # 그 주기를 잰 도면들의 SDF typesize (코퍼스 334개 전부)

_TF_OP = re.compile(r"/\S+\s+([\d.]+)\s+Tf\b")


def _font_sizes(pdf: Path) -> collections.Counter:
    """PDF 가 실제로 쓴 글자 크기(pt)를 몇 번씩 썼는지 함께 돌려준다."""
    from pypdf import PdfReader

    data = PdfReader(str(pdf)).pages[0].get_contents().get_data().decode("latin-1")
    return collections.Counter(round(float(v), 4) for v in _TF_OP.findall(data))


def test_label_height_is_fixed_in_model_units(hand, tmp_path):
    # 참조 40 장에서 값 글씨 높이는 모델 좌표에 고정이다 — 모델 단위 변동계수 0.029,
    # 종이 pt 로 보면 0.172, 종이 비율로 보면 0.286. 종이에 고정한 크기를 쓰면 큰
    # 도면에서 글씨만 커진다. 렌더러 상수를 가져오지 않고 실측값을 다시 적는다.
    pdf = tmp_path / "labels.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity")

    to_pt = _data_to_pt(hand.model)
    per_unit = to_pt(1.0, 0.0)[0] - to_pt(0.0, 0.0)[0]
    sizes = _font_sizes(pdf)
    # 값 라벨이 압도적으로 많다. 표제란·범례·각주는 종이에 고정이라 섞이면 안 된다.
    label_pt, count = sizes.most_common(1)[0]
    assert count > 20, f"값 라벨을 못 찾았다: {sizes.most_common()}"
    assert label_pt == pytest.approx(_LABEL_UNITS * per_unit, rel=1e-3)

    # 도면 주기도 같은 자에 걸린다. 이 도면의 typesize 는 60 이라 참조(30)의 두 배다.
    typesize = {t.typesize for t in hand.model.texts}
    assert typesize == {60.0}, f"주기 typesize 가 달라졌다: {typesize}"
    want_note = _NOTE_UNITS * 60.0 / _NOTE_TYPESIZE * per_unit
    assert any(s == pytest.approx(want_note, rel=1e-3) for s in sizes), \
        f"주기 글씨 {want_note:.3f}pt 를 못 찾았다: {sorted(sizes)}"


def test_sheet_form_matches_the_reference(hand, tmp_path):
    # 참조 PIPENET 지면 실측(595.22×842.00pt). 테두리는 네 변 22.68pt(8mm) 안쪽이고
    # 표제란은 그 오른쪽 아래 모서리에 붙는 362.64pt(128mm) 폭 상자다. 행은 위에서
    # 5/5/5/10/10mm 이고, 2 행은 절반 · 3 행은 1/4·3/4 에서 갈린다.
    inset, block_w = 22.68, 362.64
    w_pt = 210.0 / 25.4 * 72.0
    pdf = tmp_path / "form.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity", node_item="Node pressure")
    horizontal, vertical = _rules(pdf)

    assert min(x for x, _, _ in vertical) == pytest.approx(inset, abs=0.05)
    assert max(x for x, _, _ in vertical) == pytest.approx(w_pt - inset, abs=0.05)
    assert min(y for y, _, _ in horizontal) == pytest.approx(inset, abs=0.05)
    assert max(y for y, _, _ in horizontal) == pytest.approx(_A4_H_PT - inset, abs=0.05)

    left = w_pt - inset - block_w
    rows = sorted(y for y, x0, x1 in horizontal if x1 - x0 == pytest.approx(block_w, abs=0.05))
    assert [b - a for a, b in zip([inset] + rows, rows)][::-1] == pytest.approx(
        [14.22, 14.16, 14.16, 28.32, 28.26], abs=0.05)

    dividers = sorted(x - left for x, y0, y1 in vertical if y1 - y0 == pytest.approx(14.16, abs=0.05))
    assert dividers == pytest.approx([block_w / 4, block_w / 2, block_w * 3 / 4], abs=0.05)


def test_node_bands_are_grey_not_the_link_colours(hand, tmp_path):
    # 참조 압력 도면 40 장 중 39 장이 노드 범례에 이 무채색 계단을 쓴다 — 검정에서
    # 시작해 한 칸에 42/255 씩 밝아진다. 관로 띠와 같은 여섯 색을 쓰는 장은 없다.
    grey = [tuple(round(v / 255.0, 4) for _ in range(3))
            for v in (0x00, 0x2a, 0x54, 0x7e, 0xa9, 0xd4)]

    pdf = tmp_path / "both.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity", node_item="Node pressure")
    fills = _legend_fills(pdf)
    # 두 벌이 함께 나오면 노드가 위다. 위에서부터 읽으므로 노드 벌이 먼저 온다.
    assert fills[:6] == grey
    assert all(len(set(c)) == 1 for c in fills[:6]), "노드 범례에 색이 섞였다"


# ── 흐름 화살표 ─────────────────────────────────────────────────────────────

# PIPENET 이 직접 출력한 ISO PDF 255 장을 SDF 모델 좌표에 맞춰 실측한 값이다.
# 렌더러에서 import 하지 않고 다시 적는다 — 같은 상수를 돌려 쓰면 값이 틀려도
# 양쪽이 똑같이 틀려서 검사가 통과한다.
_WING_UNITS = 6.76           # 날개 길이(모델 단위). 변동계수 0.029
_WING_HALF_ANGLE = 26.565    # atan(1/2). 벌어진각 53.13°, 변동계수 0.025
_ARROW_FRACTION = 0.687      # 관로 호길이의 이 지점. 4분위 0.672/0.687/0.697


def _chevron_at(path, wing, reverse):
    """호길이 0.687 지점에 꼭짓점을 둔 갈매기표 두 획."""
    import math

    spans = [math.dist(a, b) for a, b in zip(path, path[1:])]
    want, acc = sum(spans) * _ARROW_FRACTION, 0.0
    for (a, b), span in zip(zip(path, path[1:]), spans):
        if span and acc + span >= want:
            t = (want - acc) / span
            tip = (a[0] + t * (b[0] - a[0]), a[1] + t * (b[1] - a[1]))
            ang = math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])) + (180 if reverse else 0)
            return [[tip, (tip[0] + wing * math.cos(math.radians(ang + 180 + side)),
                           tip[1] + wing * math.sin(math.radians(ang + 180 + side)))]
                    for side in (-_WING_HALF_ANGLE, _WING_HALF_ANGLE)]
        acc += span
    raise AssertionError("호길이 지점을 찾지 못했다")


def test_flow_arrows_are_pipe_coloured_chevrons(hand, tmp_path):
    # PIPENET 은 채운 검은 화살촉이 아니라 관로와 같은 색·같은 선폭의 열린
    # 갈매기표를 그린다. 크기는 도면 크기가 아니라 모델 좌표에 붙어 있다.
    from core.d_result_binder import normalize_label

    pdf = tmp_path / "arrows.pdf"
    render_iso(hand, pdf, link_item="Pipe velocity")   # 관로가 띠 색이라 회색과 갈린다

    model = hand.model
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    to_pt = _data_to_pt(model)
    origin, unit = to_pt(0.0, 0.0), to_pt(1.0, 0.0)
    wing_pt = _WING_UNITS * (unit[0] - origin[0])

    drawn = _stroked_paths(pdf)
    strokes = [(colour, pts) for colour, pts in drawn if len(pts) == 2]

    checked = 0
    for pipe in model.pipes:
        result = hand.pipes.get(normalize_label(pipe.label))
        flow = result.flow_m3s if result else None
        if not flow:
            continue
        path = [to_pt(*p) for p in
                (coords[pipe.input_node], *pipe.waypoints, coords[pipe.output_node])]
        colour = next(c for c, pts in drawn
                      if len(pts) == len(path) and _same_segment(path, pts))
        assert isinstance(colour, tuple), f"관로 {pipe.label} 가 띠 색으로 안 그려졌다"
        for want in _chevron_at(path, wing_pt, flow < 0):
            assert any(c == colour and _same_segment(want, pts) for c, pts in strokes), \
                f"관로 {pipe.label} 의 갈매기표 획이 없다"
        checked += 1
    assert checked == 103


def test_arrows_switched_off_leave_the_pipes_bare(hand, tmp_path):
    from core.d_result_binder import normalize_label

    on, off = tmp_path / "on.pdf", tmp_path / "off.pdf"
    render_iso(hand, on, link_item="Pipe velocity")
    render_iso(hand, off, link_item="Pipe velocity", show_arrows=False)

    model = hand.model
    coords = {n.label: (n.x, n.y) for n in model.nodes}
    to_pt = _data_to_pt(model)
    origin, unit = to_pt(0.0, 0.0), to_pt(1.0, 0.0)
    wing_pt = _WING_UNITS * (unit[0] - origin[0])

    pipe = next(p for p in model.pipes
                if (hand.pipes.get(normalize_label(p.label)) or None)
                and hand.pipes[normalize_label(p.label)].flow_m3s)
    flow = hand.pipes[normalize_label(pipe.label)].flow_m3s
    path = [to_pt(*p) for p in
            (coords[pipe.input_node], *pipe.waypoints, coords[pipe.output_node])]
    want = _chevron_at(path, wing_pt, flow < 0)[0]

    def has_wing(pdf):
        return any(len(pts) == 2 and _same_segment(want, pts)
                   for _, pts in _stroked_paths(pdf))

    assert has_wing(on) and not has_wing(off)


# ── 원본 표시 설정 따르기 ───────────────────────────────────────────────────


def test_default_follows_the_switches_the_file_was_saved_with(hand, tmp_path):
    # 제출용 두 파일 모두 PIPENET 이 이름표를 끄고 화살표를 켠 채로 저장했다.
    seen = hand.model.source_display
    assert (seen.link_labels, seen.node_labels, seen.flow_arrows) == (False, False, True)

    report = render_iso(hand, tmp_path / "asis.pdf", preset="유량본")
    assert (report.link_labels, report.node_labels, report.flow_arrows) == (False, False, True)

    # 이름표가 빠진 만큼 글자가 줄어야 한다 — 값 글자는 그대로 남는다.
    bare = len(_text_items(tmp_path / "asis.pdf"))
    render_iso(hand, tmp_path / "tagged.pdf", preset="유량본", **_TAGGED)
    assert bare < len(_text_items(tmp_path / "tagged.pdf"))


def test_caller_overrides_what_the_file_says(hand, tmp_path):
    report = render_iso(hand, tmp_path / "over.pdf", preset="유량본",
                        show_link_labels=True, show_arrows=False)
    assert report.link_labels is True and report.flow_arrows is False
    assert report.node_labels is False           # 시키지 않은 것은 원본을 따른다


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
