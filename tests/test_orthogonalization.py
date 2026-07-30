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


def test_pipeline_total_length_preserved(anchored):
    off, on, _ao, _an = anchored
    lo = sum(L for _a, _b, L in off.edges)
    ln = sum(L for _a, _b, L in on.edges)
    assert ln == pytest.approx(lo, rel=1e-9)


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


def test_estimated_edge_coords_follow_snap(anchored):
    """브릿지·용접·head-drop 이 최종망 위에 있었다면 스냅 후에도 있어야 한다.

    프론트가 이 좌표를 최종망 위에 점선으로 겹쳐 그리므로, 스냅으로 노드만 옮기고
    좌표를 그대로 두면 표시가 어긋난다. 병합으로 이미 사라진 노드는 대상이 아니다.
    """
    off, on, ao, an = anchored
    nodes_off = {n for a, b, _L in off.edges for n in (a, b)}
    nodes_on = {n for a, b, _L in on.edges for n in (a, b)}
    checked = 0
    for key in ("bridges", "welds", "head_drops"):
        for rec_off, rec_on in zip(ao.get(key) or [], an.get(key) or []):
            for pk in ("p1", "p2"):
                if (rec_off[pk][0], rec_off[pk][1]) not in nodes_off:
                    continue
                checked += 1
                assert (rec_on[pk][0], rec_on[pk][1]) in nodes_on, \
                    f"{key}.{pk} 좌표가 직교화 노드 이동을 따라가지 않음"
    assert checked > 0, "검사 대상 추정-edge 끝점이 하나도 없음 — 테스트 무의미"
