# -*- coding: utf-8 -*-
"""유량 누적(load) 기반 배관망 확정 테스트 — ModuleA 유량누적 작업지시서 §4.

실행::

    python -m pytest tests/test_load_extraction.py -q
"""
from __future__ import annotations

import math
import random
import sys
from collections import defaultdict
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

import remote30_prototype as rp  # noqa: E402


def _mkgraph(edges):
    """[(a, b)] → (graph, edge_len). 길이는 실 유클리드 거리."""
    graph = defaultdict(set)
    edge_len = {}
    for a, b in edges:
        graph[a].add(b)
        graph[b].add(a)
        edge_len[(min(a, b), max(a, b))] = math.hypot(b[0] - a[0], b[1] - a[1])
    return graph, edge_len


def _key(a, b):
    return (min(a, b), max(a, b))


# ── L1 ──────────────────────────────────────────────────────────────────────

S = (0.0, 0.0)
A = (0.0, 1000.0)
B = (0.0, 2000.0)
H1 = (1000.0, 1000.0)
H2 = (1000.0, 2000.0)
H3 = (-1000.0, 2000.0)
D = (-1000.0, 1000.0)      # 헤드 없는 막다른 가지


def test_l1_subtree_terminal_counts():
    """소스→분기→헤드 3개 + 막다른 가지 1개: 주관 3 / 중간 2 / 말단 1 / 막다른 0."""
    graph, edge_len = _mkgraph([(S, A), (A, B), (A, H1), (B, H2), (B, H3), (A, D)])
    load = rp.compute_edge_load(graph, edge_len, S, [H1, H2, H3])
    assert load[_key(S, A)] == 3
    assert load[_key(A, B)] == 2
    assert load[_key(A, H1)] == 1
    assert load[_key(B, H2)] == 1
    assert load[_key(B, H3)] == 1
    assert load[_key(A, D)] == 0


def test_l1_source_incident_load_sums_to_head_count():
    """헤드 수 N 인 임의 트리에서 소스 인접 간선 부하 합 = N."""
    rng = random.Random(20260729)
    for _ in range(20):
        nodes = [S]
        edges = []
        for i in range(1, 40):
            parent = nodes[rng.randrange(len(nodes))]
            child = (parent[0] + rng.randrange(-3, 4) * 500.0 + i * 1.0,
                     parent[1] + rng.randrange(1, 4) * 500.0)
            nodes.append(child)
            edges.append((parent, child))
        graph, edge_len = _mkgraph(edges)
        leaves = [n for n in nodes if n != S and len(graph[n]) == 1]
        heads = rng.sample(leaves, k=min(len(leaves), rng.randrange(1, 8)))
        load = rp.compute_edge_load(graph, edge_len, S, heads)
        assert sum(load[_key(S, v)] for v in graph[S]) == len(heads)


def test_l1_unreachable_heads_collected():
    """다른 컴포넌트의 헤드는 조용히 버리지 않고 unreachable 로 수집된다."""
    far = (500000.0, 500000.0)
    far2 = (500000.0, 501000.0)
    graph, edge_len = _mkgraph([(S, A), (A, H1), (far, far2)])
    unreach: list = []
    load = rp.compute_edge_load(graph, edge_len, S, [H1, far2],
                                unreachable_out=unreach)
    assert load[_key(S, A)] == 1
    assert load[_key(far, far2)] == 0
    assert unreach == [far2]


def test_l2_grid_loop_with_all_positive_load_is_preserved():
    """네 변 부하가 (2,1,2,1)인 사각 루프 → 절단 0건, 사이클 유지."""
    P = [(0.0, 0.0), (1000.0, 0.0), (1000.0, 1000.0), (0.0, 1000.0)]
    ring = [(P[0], P[1]), (P[1], P[2]), (P[2], P[3]), (P[3], P[0])]
    graph, edge_len = _mkgraph(ring)
    load = dict(zip((_key(a, b) for a, b in ring), (2, 1, 2, 1)))
    audit: dict = {}
    pruned = rp.prune_by_load(graph, edge_len, load, audit)
    assert pruned == set()
    assert len(edge_len) == 4
    assert audit["residual_cycles"] == {"count": 1, "policy": "preserve"}


def test_l2_cycle_with_zero_load_edge_cuts_only_that_edge():
    """네 변 부하가 (3,2,1,0)인 사각 루프 → 부하 0인 변 1개만 절단."""
    P = [(0.0, 0.0), (1000.0, 0.0), (1000.0, 1000.0), (0.0, 1000.0)]
    ring = [(P[0], P[1]), (P[1], P[2]), (P[2], P[3]), (P[3], P[0])]
    graph, edge_len = _mkgraph(ring)
    load = dict(zip((_key(a, b) for a, b in ring), (3, 2, 1, 0)))
    audit: dict = {}
    pruned = rp.prune_by_load(graph, edge_len, load, audit)
    assert pruned == {_key(P[3], P[0])}
    assert len(edge_len) == 3
    assert audit["pruned"]["dead_edge_count"] == 1
    assert audit["residual_cycles"]["count"] == 0


def test_l2_dead_branch_removed_but_watered_path_kept():
    """도달 가능하지만 헤드가 없는 막다른 가지는 제거, 급수 경로는 보존."""
    graph, edge_len = _mkgraph([(S, A), (A, B), (A, H1), (B, H2), (B, H3), (A, D)])
    load = rp.compute_edge_load(graph, edge_len, S, [H1, H2, H3])
    audit: dict = {}
    pruned = rp.prune_by_load(graph, edge_len, load, audit)
    assert pruned == {_key(A, D)}
    assert all(load[k] >= 1 for k in edge_len)


def test_l2_force_tree_policy_breaks_residual_cycle():
    P = [(0.0, 0.0), (1000.0, 0.0), (1000.0, 1000.0), (0.0, 1000.0)]
    ring = [(P[0], P[1]), (P[1], P[2]), (P[2], P[3]), (P[3], P[0])]
    graph, edge_len = _mkgraph(ring)
    load = dict.fromkeys((_key(a, b) for a, b in ring), 1)
    audit: dict = {}
    rp.prune_by_load(graph, edge_len, load, audit,
                     on_residual_cycle="force_tree", source=P[0])
    assert audit["residual_cycles"] == {"count": 0, "policy": "force_tree"}
    assert audit["pruned"]["cycle_cut_count"] == 1
    assert len(edge_len) == 3


def test_l1_nontree_edges_are_zero():
    """사이클을 닫는 비트리 간선의 부하는 0."""
    #  S ─ A ─ B ─ H2   +  A ─ B 를 우회하는 긴 변 (A ─ C ─ B)
    C = (3000.0, 1500.0)
    graph, edge_len = _mkgraph([(S, A), (A, B), (B, H2), (A, C), (C, B)])
    load = rp.compute_edge_load(graph, edge_len, S, [H2])
    assert load[_key(S, A)] == 1
    assert load[_key(A, B)] == 1
    assert load[_key(B, H2)] == 1
    # 우회로는 최단경로 트리 밖 — 부하 0
    assert load[_key(A, C)] == 0
    assert load[_key(C, B)] == 0
