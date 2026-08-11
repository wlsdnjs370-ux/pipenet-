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

# 표제란·범례가 놓이는 위쪽 띠, 각주가 놓이는 아래쪽 띠(pt). 그 사이가 그림틀이다.
_FRAME_TOP_PT = (1.0 - (12.0 + 16.0 + 18.0) / 297.0) * 297.0 / 25.4 * 72.0
_FRAME_BOTTOM_PT = (12.0 + 6.0) / 25.4 * 72.0


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


_STREAM_OP = re.compile(
    r"(?P<grey>[\d.]+) G\b|(?P<rgb>(?:[\d.]+ ){3})RG\b|(?P<stack>(?<![\w.])[qQ](?![\w.]))"
    r"|(?P<x>[\d.\-]+) (?P<y>[\d.\-]+) [mlc]\b")


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
        if float(m["y"]) < _FRAME_TOP_PT:
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


def _data_to_pt(model):
    """데이터 좌표 → 페이지 pt. 종이·여백 수치를 렌더러에서 가져오지 않고 다시 적는다."""
    minx, miny, maxx, maxy = model.bounds()
    span_x, span_y = maxx - minx, maxy - miny
    page = (297.0, 210.0) if span_x / span_y > 297.0 / 210.0 else (210.0, 297.0)
    mm = 72.0 / 25.4
    box_w = (page[0] - 24.0) * mm
    box_h = (page[1] - (12.0 + 6.0) - (12.0 + 16.0 + 18.0)) * mm
    box_cx, box_cy = page[0] * mm / 2.0, (12.0 + 6.0) * mm + box_h / 2.0
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
