# -*- coding: utf-8 -*-
"""2앵커(수동 밸브 + 헤드 영역) anchored 추출 테스트 — ModuleA 작업지시서 §4.

fixture: samples/dxf/대명동201동 단위세대_layer정리.dxf
(지시서의 `1__입력도면_대명동_단위세대_평면도.dxf` 와 동일 도면 — L4 SPLINE 11,770개,
범례 블록 A$C60792707/(288201,−233417) · A$C3F157AFD/(288201,−234617) 실측 일치로 확정.)

실행::

    python -m pytest tests/test_anchored_extraction.py -q
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import ezdxf
import pytest

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import remote30_prototype as rp  # noqa: E402
from remote30_prototype import HeadDetection, HeadRegion  # noqa: E402

FIXTURE = _ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"

# 범례 표본 헤드 (지시서 §0 실측 근거) — 승인 목록에 절대 없어야 하는 좌표
LEGEND_XY = [(288201.0, -233417.2), (288201.0, -234617.2)]

# 서쪽 세대 1곳을 덮는 region 다각형 (하드코딩 — 지시서 W1 수용 기준)
WEST_UNIT_POLY = [
    (244500.0, -243500.0),
    (253500.0, -243500.0),
    (253500.0, -221500.0),
    (244500.0, -221500.0),
]


@pytest.fixture(scope="module")
def bundle():
    assert FIXTURE.exists(), f"fixture 없음: {FIXTURE}"
    return rp.parse_dxf_bundle(FIXTURE)


@pytest.fixture(scope="module")
def layer_categories(bundle):
    return {ly["name"]: ly["auto_category"] for ly in bundle.layers}


@pytest.fixture(scope="module")
def pipe_ents(bundle):
    return rp.filter_pipenet_only(bundle)


# ── §4 레이어 분류 ──────────────────────────────────────────────────────────

@pytest.mark.parametrize("layer,expected", [
    ("-소화(SP가지관)", "PIPE"),
    ("소화기CO2", "EXCLUDE"),
    ("L4", "OTHER"),
    ("SHEET-TEXT", "TEXT"),
    ("1.68℃하향식", "HEAD"),
])
def test_layer_classification(layer_categories, layer, expected):
    assert layer_categories[layer] == expected


# ── §4 헤드층 원시 CIRCLE(r=52.6mm) 9개 → R2 승인 ──────────────────────────

def test_r2_raw_circles_approved(pipe_ents, layer_categories):
    raw = [en for en in pipe_ents
           if en["t"] == "C"
           and layer_categories.get(en.get("l", ""), "OTHER") == "HEAD"
           and abs(float(en.get("r", 0)) - 52.6) < 0.5]
    assert len(raw) == 9
    heads = rp.detect_heads(pipe_ents, layer_categories)
    for en in raw:
        cx, cy = float(en["c"][0]), float(en["c"][1])
        near = [h for h in heads
                if math.hypot(h.pos[0] - cx, h.pos[1] - cy) <= 250.0
                and "circle_signature" in h.kind]
        assert near, f"R2 미승인 CIRCLE: ({cx}, {cy})"


# ── W1 수용 기준 ────────────────────────────────────────────────────────────

def test_w1_region_gate_excludes_legend(pipe_ents, layer_categories):
    region = HeadRegion.from_polygon(WEST_UNIT_POLY)
    heads = rp.detect_heads(pipe_ents, layer_categories, region=region)
    assert heads, "서쪽 세대 region 안에 승인 헤드가 있어야 함"
    for lx, ly in LEGEND_XY:
        hit = [h for h in heads if math.hypot(h.pos[0] - lx, h.pos[1] - ly) < 1.0]
        assert not hit, f"범례 표본 헤드가 승인됨: ({lx}, {ly})"
    for h in heads:
        assert region.contains(h.pos), f"region 밖 헤드 승인: {h.pos}"


def test_w1_region_none_regression(pipe_ents, layer_categories):
    base = rp.detect_heads(pipe_ents, layer_categories)
    none_kw = rp.detect_heads(pipe_ents, layer_categories, region=None)
    assert base == none_kw
    # 범례 표본은 region 미지정 시 여전히 검출됨 (게이트가 검출 자체를 바꾸지 않음)
    for lx, ly in LEGEND_XY:
        hit = [h for h in base if math.hypot(h.pos[0] - lx, h.pos[1] - ly) < 1.0]
        assert hit, f"region=None 인데 범례 표본 미검출: ({lx}, {ly})"


# ── W1.2 미도달 헤드 보고 헬퍼 (합성 그래프) ───────────────────────────────

def test_w1_find_unreachable_region_heads():
    a, b = (0.0, 0.0), (1000.0, 0.0)
    c, d = (50000.0, 0.0), (51000.0, 0.0)
    graph = {a: {b}, b: {a}, c: {d}, d: {c}}

    def _head(x, y):
        return HeadDetection(pos=(x, y), bbox=(x - 1, y - 1, x + 1, y + 1),
                             kind="t", confidence=0.8)

    reachable = _head(900.0, 100.0)          # source 컴포넌트에 부착
    foreign = _head(50900.0, 100.0)          # 다른 컴포넌트 최근접
    orphan = _head(200000.0, 0.0)            # HEAD_DROP_MAX_MM 밖
    out = rp.find_unreachable_region_heads(graph, a, [reachable, foreign, orphan])
    assert reachable.pos not in out
    assert foreign.pos in out
    assert orphan.pos in out
    # source 가 그래프에 없으면 전원 미도달 (조용한 drop 금지)
    assert rp.find_unreachable_region_heads({}, None, [reachable]) == [reachable.pos]


# ── W2 수용 기준: attach_source (합성 그래프) ──────────────────────────────

def _w2_graph():
    """102mm 무헤드 고립 조각 vs 606mm 승인 헤드 보유 본망 (지시서 §0 실측 재현)."""
    n1, n2 = (102.0, 0.0), (202.0, 0.0)                      # 고립 노이즈 조각
    m1, m2, m3 = (606.0, 0.0), (1606.0, 0.0), (2606.0, 0.0)  # 본망
    graph = {n1: {n2}, n2: {n1}, m1: {m2}, m2: {m1, m3}, m3: {m2}}
    edge_len = {(n1, n2): 100.0, (m1, m2): 1000.0, (m2, m3): 1000.0}
    comp_of = {n1: 0, n2: 0, m1: 1, m2: 1, m3: 1}
    head = HeadDetection(pos=(1600.0, 50.0), bbox=(1599, 49, 1601, 51),
                         kind="t", confidence=0.9)
    return graph, edge_len, comp_of, head


def test_w2_attach_source_prefers_head_component():
    graph, edge_len, comp_of, head = _w2_graph()
    audit = {}
    src, key = rp.attach_source((0.0, 0.0), graph, comp_of, [head], edge_len, audit)
    # blind nearest 였다면 (102,0) — 헤드 보유 본망 (606,0) 에 부착되어야 함
    assert src == (0.0, 0.0)
    assert (606.0, 0.0) in graph[src]
    assert (102.0, 0.0) not in graph[src]
    sa = audit["source_attach"]
    assert sa["escalation"] == 0 and sa["method"] == "head_component_most_heads"
    assert abs(sa["dist_mm"] - 606.0) < 1e-6 and sa["comp_head_count"] == 1
    assert key in edge_len and abs(edge_len[key] - 606.0) < 1e-6


def test_w2_attach_source_escalates_without_heads():
    graph, edge_len, comp_of, _head = _w2_graph()
    audit = {}
    comp_len = {0: 100.0, 1: 2000.0}
    src, _key = rp.attach_source((0.0, 0.0), graph, comp_of, [], edge_len, audit,
                                 comp_len=comp_len)
    # 승인 헤드가 없어도 최근접(102mm 노이즈 조각)이 아니라 연장이 긴 본망에 붙는다
    assert (606.0, 0.0) in graph[src]
    assert (102.0, 0.0) not in graph[src]
    assert audit["source_attach"]["escalation"] == 1


def test_w2_attach_source_splits_pipe_interior():
    """알람밸브가 배관 *중간* 옆에 서면 그 배관을 쪼개 붙는다 — 끝점까지 끌려가지 않음."""
    a, b = (0.0, 0.0), (10000.0, 0.0)
    graph = {a: {b}, b: {a}}
    edge_len = {(a, b): 10000.0}
    audit = {}
    src, key = rp.attach_source((5000.0, 800.0), graph, {a: 0, b: 0}, [],
                                edge_len, audit)
    foot = (5000.0, 0.0)
    assert foot in graph[src] and graph[foot] == {a, b, src}
    assert (a, b) not in edge_len
    assert audit["source_attach"]["attached_to"] == "pipe_interior"
    assert abs(audit["source_attach"]["dist_mm"] - 800.0) < 1e-6
    assert key in edge_len


def test_w2_attach_source_fails_beyond_cap():
    far = (10_000_000.0, 0.0)
    graph = {far: set()}
    audit = {}
    with pytest.raises(ValueError):
        rp.attach_source((0.0, 0.0), graph, {far: 0}, [], {}, audit)
    assert audit["source_attach"]["method"] == "failed"


def test_w3_anchored_fixture_excludes_other_units(pipe_ents, layer_categories):
    region = HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(pipe_ents, layer_categories, region=region)
    assert gated
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    audit = {}
    res = rp.select_worst30_heads_anchored(pipe_ents, layer_categories,
                                           alarm_xy=(cx, cy), head_region=region,
                                           audit_out=audit)
    assert res.heads, "anchored 선정 결과 헤드 없음"
    assert res.source_kind == "manual_anchored"
    # 범례(x≈288,000) 는 최종망에 없어야 한다. 서측(x<255,000) 으로 완전히 가둘 수는
    # 없다 — 추정 연결을 폐지한 뒤 서측 가지 일부는 공용 주배관(동측 우회)을 실제로
    # 타야 소스에 닿는다. 그 경유 노드는 실배관이므로 남는 게 맞다.
    for n in res.nodes_in_subgraph:
        assert n[0] < 270000.0, f"다른 세대 노드 포함: {n}"
    for h in res.heads:
        assert region.contains(h.pos), f"region 밖 헤드 선정: {h.pos}"
    assert audit["source_attach"]["method"] in (
        "head_component_most_heads", "any_component_nearest")
    assert "unreachable" in audit["heads"]


def test_w3_anchored_requires_both_anchors(pipe_ents, layer_categories):
    with pytest.raises(ValueError):
        rp.select_worst30_heads_anchored(pipe_ents, layer_categories,
                                         alarm_xy=None,
                                         head_region=HeadRegion.from_polygon(WEST_UNIT_POLY))
    with pytest.raises(ValueError):
        rp.select_worst30_heads_anchored(pipe_ents, layer_categories,
                                         alarm_xy=(250000.0, -232000.0),
                                         head_region=None)


# ── W4 수용 기준: HeadRegion 영역 표현 통일 ──────────────────────────────────

# L자형 다각형 — bbox(rect union)는 (0,0)~(10000,10000) 이지만 notch(우상단)는 밖
L_POLY = [(0.0, 0.0), (10000.0, 0.0), (10000.0, 4000.0),
          (4000.0, 4000.0), (4000.0, 10000.0), (0.0, 10000.0)]


def test_w4_headregion_rect_regression():
    # 기존 branch_zones in_region 판정과 동일: min/max 정규화 + 경계 포함(<=)
    r = HeadRegion.from_rects([(10.0, 10.0, 0.0, 0.0)])  # 역순 rect 도 정규화
    assert r.contains((5.0, 5.0))
    assert r.contains((0.0, 10.0))     # 경계 포함
    assert not r.contains((10.1, 5.0))
    assert not HeadRegion.from_rects([])   # 빈 region 은 falsy (no-op 게이트용)
    d = r.dilate(2.0)                  # margin 누적, 원본 불변
    assert d.contains((11.5, 5.0))
    assert not r.contains((11.5, 5.0))


def test_w4_headregion_lshape_excludes_rect_union_point():
    region = HeadRegion.from_polygon(L_POLY)
    assert region.contains((2000.0, 8000.0))   # 세로 다리 안
    assert region.contains((8000.0, 2000.0))   # 가로 다리 안
    # 사각형 union(bbox) 이었다면 물었을 notch 점 — 다각형에선 제외
    assert not region.contains((8000.0, 8000.0))
    assert HeadRegion.from_rects([(0.0, 0.0, 10000.0, 10000.0)]).contains((8000.0, 8000.0))
    # dilate 는 다각형 경계 팽창 (Minkowski 원판)
    assert region.dilate(500.0).contains((10400.0, 2000.0))
    assert not region.dilate(500.0).contains((8000.0, 8000.0))


def test_w4_restrict_lshape_excludes_neighbor_node():
    """L자형 region 을 _restrict_to_branch_region 에 직접 투입 —
    rect-union 이면 물었을 notch(이웃) 노드·가지가 제거되는지."""
    g: dict = {}
    el: dict = {}

    def add(u, v):
        g.setdefault(u, set()).add(v)
        g.setdefault(v, set()).add(u)
        el[(min(u, v), max(u, v))] = math.hypot(u[0] - v[0], u[1] - v[1])

    src = (-2000.0, 2000.0)     # 영역 밖 source
    a = (2000.0, 2000.0)        # L 안 (가로 다리)
    b = (8000.0, 2000.0)        # L 안
    notch = (8000.0, 8000.0)    # L 밖 (bbox 안) — 이웃 세대 노드
    add(src, a)
    add(a, b)
    add(b, notch)

    region_nodes = rp._restrict_to_branch_region(
        g, el, src, HeadRegion.from_polygon(L_POLY))
    assert a in region_nodes and b in region_nodes
    assert notch not in region_nodes
    assert (min(b, notch), max(b, notch)) not in el   # notch 가지 제거
    assert (min(a, b), max(a, b)) in el               # 영역 안 edge 보존
    assert (min(src, a), max(src, a)) in el           # corridor 보존
    # rect-list 입력 경로 회귀 — 동일 bbox rect 는 notch 를 물고 있어야 함(기존 의미론)
    g2: dict = {}
    el2: dict = {}

    def add2(u, v):
        g2.setdefault(u, set()).add(v)
        g2.setdefault(v, set()).add(u)
        el2[(min(u, v), max(u, v))] = math.hypot(u[0] - v[0], u[1] - v[1])

    add2(src, a)
    add2(a, b)
    add2(b, notch)
    rn2 = rp._restrict_to_branch_region(g2, el2, src, [(0.0, 0.0, 10000.0, 10000.0)])
    assert notch in rn2
    assert (min(b, notch), max(b, notch)) in el2


# ── W5 수용 기준: 공간한정 조건부 재선별 (플래그, 기본 off) ──────────────────
# 변조 fixture — 레이어명을 전부 무의미 문자열로 치환한 도면 (ezdxf 생성).
# 1차 명목 수집(filter_pipenet_only)으로는 아무 entity 도 못 얻는다.

TAMPER_REGION = [(-500.0, -500.0), (10500.0, -500.0),
                 (10500.0, 3500.0), (-500.0, 3500.0)]
TAMPER_HEADS = [(2000.0, 3000.0), (8000.0, 3000.0)]
TAMPER_ALARM = (0.0, 0.0)


def _make_tampered_dxf(path):
    doc = ezdxf.new()
    for name in ("XQZW1", "XQZW2", "XQZW3"):
        doc.layers.add(name)
    msp = doc.modelspace()
    # 배관 comb (주관 + 가지 2개) — 무의미 레이어명 → OTHER 로 분류됨
    msp.add_line((0, 0), (10000, 0), dxfattribs={"layer": "XQZW1"})
    msp.add_line((2000, 0), (2000, 3000), dxfattribs={"layer": "XQZW1"})
    msp.add_line((8000, 0), (8000, 3000), dxfattribs={"layer": "XQZW1"})
    # 열린 LWPOLYLINE — 승인 대상 유형
    msp.add_lwpolyline([(10000, 0), (10000, 2000)], dxfattribs={"layer": "XQZW2"})
    # SPLINE — W 안이라도 음성 유형 (승인 금지)
    msp.add_spline([(1000, 1000), (3000, 2500), (5000, 1200)],
                   dxfattribs={"layer": "XQZW2"})
    # 닫힌 LWPOLYLINE (박스) — 음성 유형
    msp.add_lwpolyline([(4000, 500), (4500, 500), (4500, 1000), (4000, 1000)],
                       close=True, dxfattribs={"layer": "XQZW3"})
    # W 밖 LINE — 공간 한정으로 제외돼야 함
    msp.add_line((50000, 0), (60000, 0), dxfattribs={"layer": "XQZW1"})
    doc.saveas(path)
    return path


@pytest.fixture()
def tampered(tmp_path):
    path = _make_tampered_dxf(tmp_path / "tampered.dxf")
    bundle = rp.parse_dxf_bundle(path)
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = rp.filter_pipenet_only(bundle)
    return path, layer_cat, pipe_ents


def test_w5_flag_off_fails(tampered):
    path, layer_cat, pipe_ents = tampered
    assert all(layer_cat.get(n) == "OTHER" for n in ("XQZW1", "XQZW2", "XQZW3"))
    assert not pipe_ents  # 1차 명목 수집 결과 없음
    # 플래그 off(기본) — 재선별 미발동, 명목 수집만으로는 소스 결합 실패
    with pytest.raises(ValueError):
        rp.select_worst30_heads_anchored(
            pipe_ents, layer_cat, alarm_xy=TAMPER_ALARM,
            head_region=HeadRegion.from_polygon(TAMPER_REGION),
            manual_heads=TAMPER_HEADS)


def test_w5_reselect_succeeds_without_spline(tampered):
    path, layer_cat, pipe_ents = tampered
    audit = {}
    res = rp.select_worst30_heads_anchored(
        pipe_ents, layer_cat, alarm_xy=TAMPER_ALARM,
        head_region=HeadRegion.from_polygon(TAMPER_REGION),
        manual_heads=TAMPER_HEADS,
        spatial_reselect=True, dxf_path=path, audit_out=audit)
    assert len(res.heads) == 2, "재선별로 추출이 성공해야 함"
    nn = audit["nonnominal"]
    assert nn["edge_count"] >= 4 and nn["len_mm"] > 0 and nn["ratio"] > 0
    # 공간 한정: W 밖 LINE(x 50k~60k) 유입 금지
    assert all(n[0] <= 20000.0 for n in res.nodes_in_subgraph)


def test_w5_candidates_exclude_negative_types(tampered):
    path, layer_cat, _ = tampered
    win = rp._AnchorWindow(TAMPER_REGION, TAMPER_ALARM)
    cands = rp.collect_spatial_reselect_segments(path, layer_cat, win)
    kinds = sorted(c["t"] for c in cands)
    # LINE 3 + 열린 LWPOLYLINE 1 — SPLINE 0개, 닫힌 PL 0개, W 밖 LINE 0개
    assert kinds == ["L", "L", "L", "PL"]


# ── W6 수용 기준: 관경 텍스트 위생 ──────────────────────────────────────────

def test_w6_extractor_dia_text_constants():
    from sprinkler_remote30_extractor import _match_dia_text
    assert _match_dia_text("NO.20") is None    # 구 naive \b 정규식이면 20 으로 오인
    assert _match_dia_text("25A") == 25
    assert _match_dia_text("Ø65") == 65
    assert _match_dia_text("50") == 50         # 순수 숫자
    assert _match_dia_text("옥내소화전 50") is None  # 노이즈 키워드


def test_w6_anchor_window_excludes_legend_dia_text():
    ents = [
        {"t": "L", "l": "SP", "p": [0.0, 0.0, 10000.0, 0.0]},
        {"t": "L", "l": "SP", "p": [2000.0, 0.0, 2000.0, 3000.0]},
        {"t": "L", "l": "SP", "p": [8000.0, 0.0, 8000.0, 3000.0]},
    ]
    layer_cat = {"SP": "PIPE"}
    audit = {}
    res = rp.select_worst30_heads_anchored(
        ents, layer_cat, alarm_xy=TAMPER_ALARM,
        head_region=HeadRegion.from_polygon(TAMPER_REGION),
        manual_heads=TAMPER_HEADS, audit_out=audit)
    texts = [
        {"t": "T", "v": "50", "p": [2100.0, 1500.0]},        # W 안 — 유효 후보
        {"t": "T", "v": "25A", "p": [288201.0, -233417.0]},  # 범례 좌표대 — W 밖
    ]

    def _cand_count(tables):
        return int(dict(tables.meta)["Diameter 텍스트 후보 수 (도면)"])

    t_plain = rp.build_input_tables(res, pipe_entities=ents + texts)
    t_win = rp.build_input_tables(res, pipe_entities=ents + texts,
                                  anchor_window=audit["anchor_window"])
    assert _cand_count(t_plain) == 2   # 제한 없음 — 범례 관경 문자 유입(오염)
    assert _cand_count(t_win) == 1     # W 제한 — 범례 관경 문자 제외


# ── W7 수용 기준: ExtractionAudit 리포트 ────────────────────────────────────

def test_w7_audit_schema_on_fixture(pipe_ents, layer_categories):
    import json as _json
    region = HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(pipe_ents, layer_categories, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    res = rp.select_worst30_heads_anchored(pipe_ents, layer_categories,
                                           alarm_xy=(cx, cy), head_region=region)
    aud = res.audit
    assert isinstance(aud, rp.ExtractionAudit)
    j = _json.loads(_json.dumps(aud.to_json_dict()))   # JSON 직렬화 계약
    assert set(j) == {"heads", "snap_eps", "fragments", "head_drops", "nonnominal",
                      "corridor", "source_attach", "water", "tee_splits",
                      "covered_drops", "head_gap_joins", "crossing_tees",
                      "load", "pruned", "residual_cycles",
                      "unreachable_heads", "fallback", "ortho"}
    assert j["heads"]["detected_in_region"] == len(gated)
    assert j["heads"]["attached"] + len(j["heads"]["unreachable"]) == len(gated)
    assert {"dist_mm", "method", "escalation"} <= set(j["source_attach"])
    # 이음매 간격은 도면에서 실측한다 — 채택값과 후보별 근거가 남아야 함
    assert j["snap_eps"]["chosen_mm"] in rp.SNAP_EPS_CANDIDATES_MM
    assert any(t["kept"] for t in j["snap_eps"]["trials"])
    for t in j["snap_eps"]["trials"]:
        assert {"eps_mm", "largest_component", "components", "len_mm",
                "kept"} == set(t)
    # 안 붙은 조각은 봉합하지 않는다 — 몇 조각이 얼마나 남았는지 정직하게 보고
    assert {"count", "attached_nodes", "detached_nodes",
            "detached_len_mm"} == set(j["fragments"])
    assert j["fragments"]["count"] >= 1
    assert j["fragments"]["attached_nodes"] > 0
    for e in j["head_drops"]:
        assert {"p1", "p2", "len_mm"} <= set(e)
        assert e["len_mm"] > 0
    assert j["head_drops"], "헤드 스냅(head-drop) 기록이 있어야 함"
    assert {"edge_count", "len_mm", "ratio"} <= set(j["nonnominal"])
    assert {"node_count", "len_mm"} <= set(j["corridor"])
    # 재선별 미발동 → nonnominal 0, 분기영역 미지정 → corridor 0
    assert j["nonnominal"]["edge_count"] == 0
    assert j["corridor"]["node_count"] == 0


def test_w7_nonanchored_audit_none(pipe_ents, layer_categories):
    sel = rp.select_worst30_heads(pipe_ents, layer_categories)
    assert sel.audit is None   # 비-anchored 응답에 extraction_audit 키 미노출 전제


# ── V4 수용 기준: water_load_audit 급수 감사 ────────────────────────────────

def _v4_graph():
    """급수 간선 2개(s→a→b, 헤드 부착) + 비급수 드레인 가지 1개(a→c)."""
    s, a, b, c = (0.0, 0.0), (1000.0, 0.0), (2000.0, 0.0), (1000.0, 3000.0)
    graph = {s: {a}, a: {s, b, c}, b: {a}, c: {a}}
    edge_len = {(s, a): 1000.0, (a, b): 1000.0, (a, c): 3000.0}
    head = HeadDetection(pos=(2000.0, 50.0), bbox=(1999, 49, 2001, 51),
                         kind="t", confidence=0.9)
    return graph, edge_len, s, [head]


def test_v4_water_marks_unwatered_branch():
    graph, edge_len, s, heads = _v4_graph()
    before = ({k: set(v) for k, v in graph.items()}, dict(edge_len))
    w = rp.water_load_audit(graph, edge_len, s, heads)
    assert w["heads_routed"] == 1
    assert w["watered_edge_count"] == 2          # s→a, a→b
    assert w["dead_edge_count"] == 1             # a→c (드레인)
    assert abs(w["dead_len_mm"] - 3000.0) < 1e-6
    assert abs(w["dead_ratio"] - 3000.0 / 5000.0) < 1e-9
    # 계측 전용 — 그래프를 수정하면 안 됨
    assert ({k: set(v) for k, v in graph.items()}, dict(edge_len)) == before


def test_v4_water_loop_alternate_path_counted_dead():
    """루프의 대체 경로는 부하 0 으로 잡힌다 — 삭제 금지 근거(계측만 하는 이유)."""
    graph, edge_len, s, heads = _v4_graph()
    b, c = (2000.0, 0.0), (1000.0, 3000.0)
    graph[b].add(c); graph[c].add(b)             # a-b-c 루프 폐합
    edge_len[(b, c)] = 4000.0
    w = rp.water_load_audit(graph, edge_len, s, heads)
    assert w["heads_routed"] == 1
    assert w["dead_edge_count"] == 2             # a→c, b→c 둘 다 부하 0
    assert 0.0 < w["dead_ratio"] < 1.0


def test_v4_water_no_source():
    graph, edge_len, _s, heads = _v4_graph()
    w = rp.water_load_audit(graph, edge_len, None, heads)
    assert w["heads_routed"] == 0 and w["dead_ratio"] == 0.0


def test_v4_water_in_anchored_audit(pipe_ents, layer_categories):
    region = HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(pipe_ents, layer_categories, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    res = rp.select_worst30_heads_anchored(pipe_ents, layer_categories,
                                           alarm_xy=(cx, cy), head_region=region)
    w = res.audit.to_json_dict()["water"]
    assert w["heads_routed"] == res.audit.heads["attached"]
    assert w["watered_edge_count"] > 0
    assert 0.0 <= w["dead_ratio"] <= 1.0


# ── P1 수용 기준: T분기 edge-split ──────────────────────────────────────────

TEE_GAP = 5.0  # 가지관 끝점이 주배관에 못미치는 거리. TEE_SPLIT_MAX_MM 미만이어야 한다.


def _tee_graph(*branch_feet):
    """수평 주배관 a-b 와, 그 위 지정 위치에 TEE_GAP 만큼 못미쳐 끝나는 가지관들."""
    a, b = (0.0, 0.0), (4000.0, 0.0)
    graph = {a: {b}, b: {a}}
    edge_len = {(a, b): 4000.0}
    for x, sign in branch_feet:
        u, v = (x, TEE_GAP * sign), (x, 1000.0 * sign)
        graph[u] = {v}; graph[v] = {u}
        edge_len[(min(u, v), max(u, v))] = 1000.0 - TEE_GAP
    return graph, edge_len, a, b


def test_p1_tee_split_at_midspan():
    graph, edge_len, a, b = _tee_graph((2000.0, 1))
    u = (2000.0, TEE_GAP)
    rec: list = []
    assert rp._split_tee_branches(graph, edge_len, splits_out=rec) == 1
    assert (a, b) not in edge_len and b not in graph[a]
    assert graph[u] == {(2000.0, 1000.0), a, b}
    assert (a, u) in edge_len and (u, b) in edge_len
    # 분기점은 수선발이 아니라 끝점 u 자신 — 새 좌표를 만들지 않는다
    assert abs(edge_len[(a, u)] - math.hypot(2000.0, TEE_GAP)) < 1e-9
    assert rec[0]["p"] == [2000.0, TEE_GAP] and abs(rec[0]["gap_mm"] - TEE_GAP) < 1e-9


def test_p1_tee_split_skips_endpoint_proximity():
    """수선발이 주배관 끝점 코앞이면 T분기가 아니라 끝점 접속 — weld 에 맡긴다."""
    graph, edge_len, a, b = _tee_graph((30.0, 1))
    assert rp._split_tee_branches(graph, edge_len) == 0
    assert (a, b) in edge_len


def test_p1_tee_split_far_gap_untouched():
    graph, edge_len, a, b = _tee_graph((2000.0, 1))
    assert rp._split_tee_branches(graph, edge_len, max_gap_mm=TEE_GAP / 2.0) == 0
    assert (a, b) in edge_len


def test_p1_tee_split_two_branches_same_trunk():
    """한 주배관에 가지 둘 — 앞선 split 으로 생긴 조각에 다시 붙어 체인이 된다."""
    graph, edge_len, a, b = _tee_graph((1000.0, 1), (3000.0, -1))
    u1, u2 = (1000.0, TEE_GAP), (3000.0, -TEE_GAP)
    assert rp._split_tee_branches(graph, edge_len) == 2
    assert (a, b) not in edge_len and (u1, b) not in edge_len
    assert graph[u1] == {(1000.0, 1000.0), a, u2}
    assert graph[u2] == {(3000.0, -1000.0), u1, b}


def test_p1_tee_split_always_on_in_anchored(pipe_ents, layer_categories):
    """추정 연결이 없는 지금 T분기가 유일한 실배관 접속 복원 수단 — 끌 수 없다."""
    region = HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(pipe_ents, layer_categories, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    aud = rp.select_worst30_heads_anchored(
        pipe_ents, layer_categories, alarm_xy=(cx, cy), head_region=region).audit
    assert aud.tee_splits, "대명동 fixture 에 T분기 후보가 있어야 함"
    for r in aud.tee_splits:
        assert set(r) == {"p", "edge", "gap_mm", "sym_mm"}
        assert 0.0 <= r["gap_mm"] <= rp.TEE_SPLIT_MAX_MM


# ── R7: 헤드 기호가 끊어놓은 동일선상 배관 접속 ────────────────────────────
HEAD_GAP = 300.0


def _head_gap_graph(gap: float = HEAD_GAP, offset: float = 0.0):
    """세로 런이 gap 만큼 끊긴 두 조각. offset 은 위 조각의 x 어긋남."""
    lo = ((0.0, 0.0), (0.0, 1000.0))
    hi = ((offset, 1000.0 + gap), (offset, 2500.0))
    graph: dict = {}
    edge_len: dict = {}
    for a, b in (lo, hi):
        graph.setdefault(a, set()).add(b)
        graph.setdefault(b, set()).add(a)
        edge_len[(min(a, b), max(a, b))] = math.hypot(b[0] - a[0], b[1] - a[1])
    return graph, edge_len, lo[1], hi[0]


def test_r7_head_gap_join_collinear_with_head():
    graph, edge_len, u, v = _head_gap_graph()
    rec: list = []
    assert rp._join_head_gap_endpoints(graph, edge_len, [(0.0, 1150.0)],
                                       joins_out=rec) == 1
    assert v in graph[u]
    assert edge_len[(u, v)] == pytest.approx(HEAD_GAP)
    assert rec[0]["gap_mm"] == pytest.approx(HEAD_GAP)


def test_r7_head_gap_join_requires_head_between():
    """③ 이 없으면 그냥 가까운 두 끝점 — 추정 연결이라 잇지 않는다."""
    graph, edge_len, u, v = _head_gap_graph()
    assert rp._join_head_gap_endpoints(graph, edge_len, [(0.0, 3000.0)]) == 0
    assert v not in graph[u]


def test_r7_head_gap_join_requires_collinear():
    graph, edge_len, u, v = _head_gap_graph(offset=200.0)
    assert rp._join_head_gap_endpoints(graph, edge_len, [(100.0, 1150.0)]) == 0
    assert v not in graph[u]


def test_r7_head_gap_join_far_gap_untouched():
    gap = rp.HEAD_GAP_JOIN_MAX_MM + 100.0
    graph, edge_len, u, v = _head_gap_graph(gap=gap)
    assert rp._join_head_gap_endpoints(graph, edge_len,
                                       [(0.0, 1000.0 + gap / 2.0)]) == 0
    assert v not in graph[u]


def test_r7_head_gap_join_rejects_perpendicular_neighbor():
    """① 인접 간선이 다른 축이면 런의 연속이 아니라 모서리 — 잇지 않는다."""
    graph, edge_len, u, _v = _head_gap_graph()
    v, w = (0.0, 1300.0), (1500.0, 1300.0)
    graph[v] = {w}
    graph[w] = {v}
    del edge_len[((0.0, 1300.0), (0.0, 2500.0))]
    graph.pop((0.0, 2500.0))
    edge_len[(v, w)] = 1500.0
    assert rp._join_head_gap_endpoints(graph, edge_len, [(0.0, 1150.0)]) == 0
    assert v not in graph[u]


def test_r7_head_gap_join_skips_already_connected():
    """이미 이어진 조각끼리는 안 잇는다 — 루프를 만들지 않는다."""
    graph, edge_len, u, v = _head_gap_graph()
    side = (900.0, 1250.0)
    for n in ((0.0, 0.0), (0.0, 2500.0)):
        graph.setdefault(side, set()).add(n)
        graph[n].add(side)
        edge_len[(min(side, n), max(side, n))] = math.hypot(side[0] - n[0],
                                                            side[1] - n[1])
    assert rp._join_head_gap_endpoints(graph, edge_len, [(0.0, 1150.0)]) == 0
    assert v not in graph[u]


def test_r7_head_gap_join_chain_three_pieces():
    """헤드 2개로 3조각 난 런 — 한 번에 전부 이어져 단일 컴포넌트가 된다."""
    graph: dict = {}
    edge_len: dict = {}
    for y0, y1 in ((0.0, 1000.0), (1300.0, 2300.0), (2600.0, 3600.0)):
        a, b = (0.0, y0), (0.0, y1)
        graph.setdefault(a, set()).add(b)
        graph.setdefault(b, set()).add(a)
        edge_len[(a, b)] = y1 - y0
    assert rp._join_head_gap_endpoints(graph, edge_len,
                                       [(0.0, 1150.0), (0.0, 2450.0)]) == 2
    assert len(rp._connected_components(graph)) == 1


# ── R8: 관통 교차 티 복원 ───────────────────────────────────────────────────

def _add_edge(graph: dict, edge_len: dict, a, b) -> None:
    graph.setdefault(a, set()).add(b)
    graph.setdefault(b, set()).add(a)
    edge_len[(min(a, b), max(a, b))] = math.hypot(b[0] - a[0], b[1] - a[1])


def _cross_graph(with_stub: bool = True):
    """세로 가지관이 가로 주배관 중간을 관통. 양쪽 다 3노드 이상(비고립)."""
    graph: dict = {}
    edge_len: dict = {}
    _add_edge(graph, edge_len, (0.0, 1000.0), (5000.0, 1000.0))       # 가로 주배관
    _add_edge(graph, edge_len, (5000.0, 1000.0), (9000.0, 1000.0))
    _add_edge(graph, edge_len, (2000.0, 0.0), (2000.0, 3000.0))       # 세로 가지관
    if with_stub:
        _add_edge(graph, edge_len, (2000.0, 3000.0), (2000.0, 6000.0))
    return graph, edge_len


def test_r8_crossing_tee_merges_two_components():
    graph, edge_len = _cross_graph()
    assert len(rp._connected_components(graph)) == 2
    cuts: list = []
    assert rp._split_crossing_tees(graph, edge_len, cuts_out=cuts) == 1
    assert len(rp._connected_components(graph)) == 1
    assert cuts[0]["p"] == [2000.0, 1000.0]
    # 교차점에서 네 갈래 — 가로·세로 둘 다 쪼개졌다
    assert len(graph[(2000.0, 1000.0)]) == 4


def test_r8_crossing_tee_never_creates_loop():
    """이미 이어진 두 배관의 교차는 스쳐 지나감과 구분 불가 — 손대지 않는다."""
    graph, edge_len = _cross_graph()
    _add_edge(graph, edge_len, (2000.0, 6000.0), (9000.0, 1000.0))  # 우회로로 연결
    assert len(rp._connected_components(graph)) == 1
    before = dict(edge_len)
    assert rp._split_crossing_tees(graph, edge_len) == 0
    assert edge_len == before


def test_r8_crossing_tee_skips_lone_segment():
    """고립 선분(노드 2개)은 배관이 아니라 표시 획 — 절단 대상 아님."""
    graph, edge_len = _cross_graph(with_stub=False)
    graph.pop((2000.0, 3000.0))
    graph[(2000.0, 0.0)].clear()
    edge_len.clear()
    _add_edge(graph, edge_len, (0.0, 1000.0), (5000.0, 1000.0))
    _add_edge(graph, edge_len, (5000.0, 1000.0), (9000.0, 1000.0))
    _add_edge(graph, edge_len, (2000.0, 900.0), (2000.0, 1100.0))  # 십자 기호 획
    assert rp._split_crossing_tees(graph, edge_len) == 0


def test_r8_crossing_tee_skips_near_endpoint():
    """끝점 근처 교차는 _split_tee_branches 담당 — 최소길이 미만 토막 금지."""
    graph, edge_len = _cross_graph()
    graph.clear(); edge_len.clear()
    _add_edge(graph, edge_len, (0.0, 1000.0), (5000.0, 1000.0))
    _add_edge(graph, edge_len, (5000.0, 1000.0), (9000.0, 1000.0))
    _add_edge(graph, edge_len, (10.0, 0.0), (10.0, 3000.0))   # 가로 끝에서 10mm
    _add_edge(graph, edge_len, (10.0, 3000.0), (10.0, 6000.0))
    assert rp._split_crossing_tees(graph, edge_len) == 0


def test_r8_crossing_tee_multiple_cuts_on_one_edge():
    """한 주배관을 두 가지관이 각각 관통 — 둘 다 접속된다."""
    graph: dict = {}
    edge_len: dict = {}
    _add_edge(graph, edge_len, (0.0, 1000.0), (9000.0, 1000.0))
    for x in (2000.0, 6000.0):
        _add_edge(graph, edge_len, (x, 0.0), (x, 3000.0))
        _add_edge(graph, edge_len, (x, 3000.0), (x, 6000.0))
    assert rp._split_crossing_tees(graph, edge_len) == 2
    assert len(rp._connected_components(graph)) == 1
    assert len(edge_len) - len(graph) + 1 == 0  # 트리 유지 (루프 0)


# ── R9: 겹쳐 그린 중복 선분 제거 ────────────────────────────────────────────

def _overlap_graph():
    """긴 세로 런의 위쪽 300mm 에 짧은 선분이 통째로 포개져 있다."""
    graph: dict = {}
    edge_len: dict = {}
    _add_edge(graph, edge_len, (0.0, 0.0), (0.0, 2400.0))
    _add_edge(graph, edge_len, (0.0, 2100.0), (0.0, 2400.0))
    return graph, edge_len


def test_r9_drops_covered_duplicate():
    graph, edge_len = _overlap_graph()
    drops: list = []
    assert rp._drop_covered_edges(graph, edge_len, drops_out=drops) == 1
    assert list(edge_len) == [((0.0, 0.0), (0.0, 2400.0))]
    assert drops[0]["len_mm"] == 300.0
    assert (0.0, 2100.0) not in graph
    assert len(graph[(0.0, 2400.0)]) == 1  # 런의 끝 — 헤드틈 접속 후보로 복귀


def test_r9_keeps_covered_edge_carrying_a_branch():
    """포개진 선분의 자유 끝점에 다른 배관이 걸려 있으면 지우지 않는다."""
    graph, edge_len = _overlap_graph()
    _add_edge(graph, edge_len, (0.0, 2100.0), (1500.0, 2100.0))
    assert rp._drop_covered_edges(graph, edge_len) == 0


def test_r9_keeps_merely_touching_edges():
    """끝점만 잇닿은 연속 런은 덮인 게 아니다."""
    graph: dict = {}
    edge_len: dict = {}
    _add_edge(graph, edge_len, (0.0, 0.0), (0.0, 2400.0))
    _add_edge(graph, edge_len, (0.0, 2400.0), (0.0, 2700.0))
    assert rp._drop_covered_edges(graph, edge_len) == 0


def test_r9_keeps_parallel_offset_run():
    """같은 축이라도 오프셋이 허용치를 넘으면 다른 배관이다."""
    graph: dict = {}
    edge_len: dict = {}
    _add_edge(graph, edge_len, (0.0, 0.0), (0.0, 2400.0))
    _add_edge(graph, edge_len, (500.0, 2100.0), (500.0, 2400.0))
    assert rp._drop_covered_edges(graph, edge_len) == 0


def test_r9_unblocks_head_gap_join():
    """중복 선분이 막고 있던 헤드틈 접속이 제거 후 성립한다."""
    graph, edge_len = _overlap_graph()
    _add_edge(graph, edge_len, (0.0, 2700.0), (0.0, 5100.0))
    heads = [(0.0, 2550.0)]
    assert rp._join_head_gap_endpoints(graph, edge_len, heads) == 0
    assert rp._drop_covered_edges(graph, edge_len) == 1
    assert rp._join_head_gap_endpoints(graph, edge_len, heads) == 1
    assert len(rp._connected_components(graph)) == 1


# ── R10: 헤드 틈 지문 배관 레이어 승격 ──────────────────────────────────────
PROMO_LAYER = "현장조사#셔터"


def _headgap_bundle(pieces: int = 4, head_x: float = 0.0,
                    category: str = "OTHER") -> rp.ParsedDxfBundle:
    """세로 런이 헤드마다 300mm 끊긴 도면. pieces-1 개의 헤드 틈이 생긴다."""
    entities = []
    for i in range(pieces):
        y0 = i * 1300.0
        entities.append({"t": "L", "l": PROMO_LAYER,
                         "p": [0.0, y0, 0.0, y0 + 1000.0]})
    for i in range(pieces - 1):
        entities.append({"t": "I", "l": "SP-헤드", "n": "SP_HEAD",
                         "p": [head_x, i * 1300.0 + 1150.0]})
    return rp.ParsedDxfBundle(
        entities=entities,
        layers=[{"name": "SP-헤드", "auto_category": "HEAD"},
                {"name": PROMO_LAYER, "auto_category": category}],
    )


def test_r10_promotes_layer_with_enough_head_gaps():
    bundle = _headgap_bundle()
    recs = rp._promote_headgap_pipe_layers(bundle)
    assert [r["layer"] for r in recs] == [PROMO_LAYER]
    assert recs[0]["prev_category"] == "OTHER"
    assert recs[0]["headgap_count"] == 3
    cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    assert cats[PROMO_LAYER] == "PIPE"


def test_r10_below_threshold_stays_unpromoted():
    """지문 2건 — 우연히 맞아떨어진 한두 쌍으로는 승격하지 않는다."""
    bundle = _headgap_bundle(pieces=3)
    assert rp._promote_headgap_pipe_layers(bundle) == []
    cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    assert cats[PROMO_LAYER] == "OTHER"


def test_r10_never_promotes_excluded_layer():
    """EXCLUDE 는 사용자가 키워드로 명시해 뺀 레이어 — 지오메트리로 뒤집지 않는다."""
    bundle = _headgap_bundle(category="EXCLUDE")
    assert rp._promote_headgap_pipe_layers(bundle) == []
    cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    assert cats[PROMO_LAYER] == "EXCLUDE"


def test_r10_requires_head_inside_the_gap():
    """헤드가 런에서 벗어난 자리에 있으면 그냥 끊긴 선 — 지문이 아니다."""
    bundle = _headgap_bundle(head_x=rp.HEAD_GAP_JOIN_TOL_MM * 3)
    assert rp._promote_headgap_pipe_layers(bundle) == []


def test_r10_no_heads_no_promotion():
    """헤드가 없는 도면에선 지문 자체가 성립하지 않는다."""
    bundle = _headgap_bundle()
    bundle.entities = [en for en in bundle.entities if en["t"] != "I"]
    assert rp._promote_headgap_pipe_layers(bundle) == []
