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

import pytest

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import remote30_prototype as rp  # noqa: E402
from remote30_prototype import HeadDetection  # noqa: E402

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


class PolyRegion:
    """테스트용 최소 region — contains((x, y)) 프로토콜 (→W4 HeadRegion 로 대체)."""

    def __init__(self, pts: list[tuple[float, float]]):
        self.pts = pts

    def contains(self, pt) -> bool:
        x, y = float(pt[0]), float(pt[1])
        inside = False
        n = len(self.pts)
        for i in range(n):
            x1, y1 = self.pts[i]
            x2, y2 = self.pts[(i + 1) % n]
            if (y1 > y) != (y2 > y):
                xin = (x2 - x1) * (y - y1) / (y2 - y1) + x1
                if x < xin:
                    inside = not inside
        return inside


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
    region = PolyRegion(WEST_UNIT_POLY)
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
    orphan = _head(200000.0, 0.0)            # HEAD_BRIDGE_MAX_MM 밖
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
    assert sa["escalation"] == 0 and sa["method"] == "head_component_nearest"
    assert abs(sa["dist_mm"] - 606.0) < 1e-6 and sa["comp_head_count"] == 1
    assert key in edge_len and abs(edge_len[key] - 606.0) < 1e-6


def test_w2_attach_source_escalates_without_heads():
    graph, edge_len, comp_of, _head = _w2_graph()
    audit = {}
    src, _key = rp.attach_source((0.0, 0.0), graph, comp_of, [], edge_len, audit)
    # 승인 헤드가 없으면 1단계 완화 — 거리 상한 내 최근접(고립 조각 102mm)
    assert (102.0, 0.0) in graph[src]
    assert audit["source_attach"]["escalation"] == 1


def test_w2_attach_source_fails_beyond_cap():
    far = (10_000_000.0, 0.0)
    graph = {far: set()}
    audit = {}
    with pytest.raises(ValueError):
        rp.attach_source((0.0, 0.0), graph, {far: 0}, [], {}, audit)
    assert audit["source_attach"]["method"] == "failed"


# ── W3 수용 기준: 표적 브릿지 ──────────────────────────────────────────────

def test_w3_bridge_targeted_only_head_components():
    s1, s2 = (0.0, 0.0), (1000.0, 0.0)              # comp(source)
    n1, n2 = (1100.0, 20.0), (1200.0, 20.0)         # 무헤드 노이즈 조각 (102mm)
    h1, h2 = (1300.0, 0.0), (2300.0, 0.0)           # 승인 헤드 보유 컴포넌트 (300mm)
    graph = {s1: {s2}, s2: {s1}, n1: {n2}, n2: {n1}, h1: {h2}, h2: {h1}}
    edge_len = {(s1, s2): 1000.0, (n1, n2): 100.0, (h1, h2): 1000.0}
    head = HeadDetection(pos=(2290.0, 10.0), bbox=(2289, 9, 2291, 11),
                         kind="t", confidence=0.9)
    audit = {}
    bridges = set()
    total = rp.bridge_targeted(graph, edge_len, s1, [head], (200.0, 500.0),
                               bridge_edges_out=bridges, audit=audit)
    # 헤드 보유 컴포넌트만 봉합 — 더 가까운(102mm) 무헤드 조각은 어떤 tol 에서도 금지
    assert total == 1
    assert h1 in graph[s2]
    assert graph[n1] == {n2} and graph[n2] == {n1}
    assert audit["bridges"][0]["tol"] == 500.0
    assert audit["bridges"][0]["p1_in_source_comp"] is True
    assert len(bridges) == 1


def test_w3_anchored_fixture_excludes_other_units(pipe_ents, layer_categories):
    region = PolyRegion(WEST_UNIT_POLY)
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
    # 최종망(SPT 후)에 다른 세대의 노드 0개 — 동측 세대(x≥255,000)·범례(x≈288,000) 배제
    for n in res.nodes_in_subgraph:
        assert n[0] < 255000.0, f"다른 세대 노드 포함: {n}"
    for h in res.heads:
        assert region.contains(h.pos), f"region 밖 헤드 선정: {h.pos}"
    # 모든 bridge 양단이 comp(source) 성장 이력에 속함 — audit 로 검증
    for b in audit.get("bridges", []):
        assert b["p1_in_source_comp"] is True
    assert audit["source_attach"]["method"] in (
        "head_component_nearest", "any_component_nearest")
    assert "unreachable" in audit["heads"]


def test_w3_anchored_requires_both_anchors(pipe_ents, layer_categories):
    with pytest.raises(ValueError):
        rp.select_worst30_heads_anchored(pipe_ents, layer_categories,
                                         alarm_xy=None,
                                         head_region=PolyRegion(WEST_UNIT_POLY))
    with pytest.raises(ValueError):
        rp.select_worst30_heads_anchored(pipe_ents, layer_categories,
                                         alarm_xy=(250000.0, -232000.0),
                                         head_region=None)
