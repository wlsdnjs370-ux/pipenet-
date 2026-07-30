# -*- coding: utf-8 -*-
"""직교화(O2) + 루프 금지 게이트(O1) 테스트.

실행::

    python -m pytest tests/test_orthogonalization.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import remote30_prototype as rp  # noqa: E402
from remote30_prototype import HeadRegion  # noqa: E402

FIXTURE = _ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf"
WEST_UNIT_POLY = [
    (244500.0, -243500.0), (253500.0, -243500.0),
    (253500.0, -221500.0), (244500.0, -221500.0),
]
TOL = 1e-9


def _is_ortho(edges) -> bool:
    return all(abs(b[0] - a[0]) <= TOL or abs(b[1] - a[1]) <= TOL for a, b, _L in edges)


def _span(p, q):
    """축평행 세그먼트 → (수평인가, 고정좌표, 하한, 상한)."""
    if abs(p[1] - q[1]) <= TOL:
        return (True, p[1], min(p[0], q[0]), max(p[0], q[0]))
    return (False, p[0], min(p[1], q[1]), max(p[1], q[1]))


def _planarity_defects(edges) -> tuple[int, int, int]:
    """(노드 관통, 직교 교차, 축선 겹침) — 셋 다 0 이어야 화면에 닫힌 영역이 없다."""
    spans = [_span(a, b) for a, b, _L in edges]
    nodes = {n for a, b, _L in edges for n in (a, b)}
    thru = cross = overlap = 0
    for i, (h1, f1, lo1, hi1) in enumerate(spans):
        a1, b1, _L1 = edges[i]
        for n in nodes:
            if n == a1 or n == b1:
                continue
            off, along = (n[1], n[0]) if h1 else (n[0], n[1])
            if abs(off - f1) <= TOL and lo1 + TOL < along < hi1 - TOL:
                thru += 1
        for j in range(i + 1, len(spans)):
            a2, b2, _L2 = edges[j]
            if {a1, b1} & {a2, b2}:
                continue
            h2, f2, lo2, hi2 = spans[j]
            if h1 != h2:
                if lo1 <= f2 <= hi1 and lo2 <= f1 <= hi2:
                    cross += 1
            elif abs(f1 - f2) <= TOL and min(hi1, hi2) - max(lo1, lo2) > TOL:
                overlap += 1
    return thru, cross, overlap


def _independent_cycles(edges) -> int:
    """E - V + C. 0 이면 숲(단일 성분이면 트리)."""
    adj: dict = {}
    for a, b, _L in edges:
        adj.setdefault(a, set()).add(b)
        adj.setdefault(b, set()).add(a)
    seen, comps = set(), 0
    for n in adj:
        if n in seen:
            continue
        comps += 1
        stack = [n]
        seen.add(n)
        while stack:
            u = stack.pop()
            for v in adj[u]:
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
    return len(edges) - len(adj) + comps


# ── orthogonalize_edges 단위 ────────────────────────────────────────────────

def test_diagonal_splits_into_two_axis_parallel_legs():
    a, b = (0.0, 0.0), (300.0, 400.0)
    out, _elb, nmap = rp.orthogonalize_edges([(a, b, 500.0)], anchors={a, b})
    assert len(out) == 2
    assert _is_ortho(out)
    assert nmap[a] == a and nmap[b] == b
    assert sum(L for _p, _q, L in out) == pytest.approx(500.0)
    legs = sorted(L for _p, _q, L in out)   # 길이는 성분비 배분 — 300:400
    assert legs == pytest.approx([500.0 * 3 / 7, 500.0 * 4 / 7])


def test_near_axis_edge_is_snapped_not_split():
    """3mm 어긋난 수평 간선은 축 스냅으로 흡수 — 0mm 스텁을 만들지 않는다."""
    a, b = (0.0, 0.0), (5000.0, 3.0)
    out, _elb, nmap = rp.orthogonalize_edges([(a, b, 5000.0)], anchors=set())
    assert len(out) == 1
    assert _is_ortho(out)
    assert nmap[a][1] == nmap[b][1] == pytest.approx(1.5)


def test_anchor_pins_the_axis_line():
    a, b = (0.0, 0.0), (5000.0, 3.0)
    out, _elb, nmap = rp.orthogonalize_edges([(a, b, 5000.0)], anchors={a})
    assert len(out) == 1
    assert nmap[a] == a                    # 앵커는 절대 안 움직인다
    assert nmap[b] == (5000.0, 0.0)


def test_conflicting_anchors_fall_back_to_split():
    """양 끝이 모두 앵커면 스냅 불가 — L자 분해로 처리하고 좌표는 보존."""
    a, b = (0.0, 0.0), (5000.0, 3.0)
    out, _elb, nmap = rp.orthogonalize_edges([(a, b, 5000.0)], anchors={a, b})
    assert nmap[a] == a and nmap[b] == b
    assert _is_ortho(out)
    assert sum(L for _p, _q, L in out) == pytest.approx(5000.0)


def test_no_node_collapse():
    """스냅이 서로 다른 노드를 한 점으로 합치면 안 된다(위상 변화 금지)."""
    a, b, c = (0.0, 0.0), (0.0, 3.0), (5000.0, 1.0)
    _out, _elb, nmap = rp.orthogonalize_edges([(a, c, 5000.0), (b, c, 5000.0)],
                                              anchors=set())
    assert len(set(nmap.values())) == len(nmap)


def test_split_preserves_elbow_fittings():
    a, b = (0.0, 0.0), (300.0, 400.0)
    elbows = {(min(a, b), max(a, b)): [((150.0, 200.0), 90.0)]}
    out, new_elb, _nmap = rp.orthogonalize_edges([(a, b, 500.0)], anchors={a, b},
                                                 elbow_fittings=elbows)
    assert sum(len(v) for v in new_elb.values()) == 1
    assert set(new_elb) <= {(min(p, q), max(p, q)) for p, q, _L in out}


def test_split_cannot_create_cycles():
    a, b, c = (0.0, 0.0), (300.0, 400.0), (900.0, 100.0)
    out, _elb, _nmap = rp.orthogonalize_edges([(a, b, 500.0), (b, c, 700.0)],
                                              anchors={a, b, c})
    assert _independent_cycles(out) == 0


def test_audit_shape():
    a, b = (0.0, 0.0), (300.0, 400.0)
    aud: dict = {}
    rp.orthogonalize_edges([(a, b, 500.0)], anchors={a, b}, audit=aud)
    assert {"snapped_nodes", "max_shift_mm", "snap_reverted_nodes",
            "snap_skipped_classes", "split_edges", "unresolved_diagonals",
            "edges_before", "edges_after"} == set(aud)


# ── 파이프라인 통합 (대명동 fixture) ────────────────────────────────────────

@pytest.fixture(scope="module")
def anchored():
    if not FIXTURE.exists():
        pytest.skip(f"fixture 없음: {FIXTURE}")
    bundle = rp.parse_dxf_bundle(FIXTURE)
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    region = HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(ents, layer_cat, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    kw = dict(alarm_xy=(cx, cy), head_region=region)
    audit_off: dict = {}
    off = rp.select_worst30_heads_anchored(ents, layer_cat, ortho=False,
                                           audit_out=audit_off, **kw)
    audit_on: dict = {}
    on = rp.select_worst30_heads_anchored(ents, layer_cat, audit_out=audit_on, **kw)
    return off, on, audit_off, audit_on


def test_pipeline_all_edges_orthogonal(anchored):
    _off, on, _ao, _an = anchored
    assert _is_ortho(on.edges)


def test_pipeline_is_a_tree(anchored):
    _off, on, _ao, _an = anchored
    assert _independent_cycles(on.edges) == 0


def test_pipeline_is_plane_embedded(anchored):
    """간선이 노드를 관통하거나 축선 위에서 포개지지 않는다.

    직교 교차(X)는 남을 수 있다 — 층이 다른 배관이 스쳐 지나가는 것인지 티인지
    평면 좌표로는 못 가리므로 추정으로 자르지 않고 감사에만 센다.
    """
    _off, on, _ao, an = anchored
    thru, cross, overlap = _planarity_defects(on.edges)
    assert (thru, overlap) == (0, 0)
    assert an["ortho"]["planar"]["unmarked_crossings"] >= cross


def test_pipeline_planarization_never_adds_length(anchored):
    """평면화는 포개진 중복·우회 배관을 걷어낼 뿐 배관을 만들어내지 않는다.

    추정 연결(용접·브리지)을 폐지한 뒤로 이 fixture 에는 걷어낼 중복이 남지
    않아 감산이 0 일 수 있다 — 늘지 않는 것이 여기서 검증할 불변식이다.
    """
    off, on, _ao, an = anchored
    lo = sum(L for _a, _b, L in off.edges)
    ln = sum(L for _a, _b, L in on.edges)
    assert ln <= lo
    p = an["ortho"]["planar"]
    # 감사값은 0.001mm 로 반올림 — 직교화까지는 길이가 보존된다.
    assert p["len_before_mm"] == pytest.approx(lo, abs=1e-3)
    assert p["len_after_mm"] == pytest.approx(ln, abs=1e-3)


def test_pipeline_anchors_and_distances_unchanged(anchored):
    off, on, _ao, _an = anchored
    assert off.source_pos == on.source_pos
    assert [h.pos for h in off.heads] == [h.pos for h in on.heads]
    assert off.distances == on.distances


def test_pipeline_no_zero_length_segments(anchored):
    _off, on, _ao, _an = anchored
    assert min(L for _a, _b, L in on.edges) > 0.0


def test_pipeline_ortho_audit_recorded(anchored):
    _off, _on, _ao, an = anchored
    o = an["ortho"]
    assert o["unresolved_diagonals"] == 0
    assert o["max_shift_mm"] <= rp.ORTHO_SNAP_EPS_MM
    assert "node_map" not in o          # 좌표 재매핑용 내부 값은 리포트에서 제거됨


def test_no_estimated_edge_points_at_empty_space(anchored):
    """head-drop 결합선이 최종망에 없는 좌표를 가리키면 안 된다.

    프론트가 이 좌표를 최종망 위에 점선으로 겹쳐 그린다. 스냅이 노드를 옮겼는데
    좌표를 그대로 두면 빈 자리에 점선이 뜨고, 평면화가 걷어낸 중복 배관에 붙어
    있던 연결은 아예 그릴 자리가 없다(그런 기록은 빠져야 한다).
    병합으로 이미 사라져 원래부터 최종망 밖이던 좌표는 대상이 아니다.
    """
    off, on, ao, an = anchored
    nodes_off = {n for a, b, _L in off.edges for n in (a, b)}
    nodes_on = {n for a, b, _L in on.edges for n in (a, b)}
    checked = 0
    for key in ("head_drops",):
        on_pts = {(r[pk][0], r[pk][1]) for r in (an.get(key) or [])
                  for pk in ("p1", "p2")}
        for rec in ao.get(key) or []:
            for pk in ("p1", "p2"):
                p = (rec[pk][0], rec[pk][1])
                if p not in nodes_off:
                    continue
                checked += 1
                assert p not in on_pts or p in nodes_on, \
                    f"{key}.{pk} {p} 가 최종망에 없는데 그대로 남아 있음"
    assert checked > 0, "검사 대상 추정-edge 끝점이 하나도 없음 — 테스트 무의미"
