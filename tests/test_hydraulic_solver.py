# -*- coding: utf-8 -*-
"""core/hydraulic_solver.py — 유속 수렴 내경 사이징 테스트.

실행::

    python -m pytest tests/test_hydraulic_solver.py -q
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from hydraulic_solver import (  # noqa: E402
    BRANCH_PIPE_V_LIMIT,
    MAIN_PIPE_V_LIMIT,
    NOZZLE_K_FACTOR,
    build_topology,
    classify_pipe_roles,
    converge_bores_by_velocity,
    demand_flows,
    hazen_williams_drop_bar,
    inner_diameter_mm,
    solve_network,
    velocity_mps,
)


def make_net(n_branch: int = 4, n_head: int = 6, dia: int = 25):
    """교차배관 1줄 + 가지배관 n_branch 줄(각 n_head 개 헤드)인 트리."""
    nodes = [{"label": "S", "io_node": "Input", "elevation": 0.0}]
    pipes: list[dict] = []
    nozzles: list[dict] = []
    prev = "S"
    for b in range(n_branch):
        cn = f"c{b}"
        nodes.append({"label": cn, "elevation": 0.0})
        pipes.append({"label": f"pc{b}", "in": prev, "out": cn,
                      "dia": dia, "length": 6.0, "c": "120", "type": "KSD 3507"})
        prev = cn
        tip = cn
        for h in range(n_head):
            hn = f"h{b}_{h}"
            nodes.append({"label": hn, "elevation": 0.0})
            pipes.append({"label": f"pb{b}_{h}", "in": tip, "out": hn,
                          "dia": dia, "length": 3.0, "c": "120", "type": "KSD 3507"})
            nozzles.append({"label": f"nz{b}_{h}", "in": hn, "lib": "SP-HEAD"})
            tip = hn
    return nodes, pipes, nozzles


# ── 물리식 ──────────────────────────────────────────────────────────────────


@pytest.mark.parametrize("q,L,c,d", [(80, 3.0, 120, 27.5), (1920, 6.0, 100, 105.3),
                                     (400, 12.5, 140, 53.2)])
def test_hazen_williams_matches_validator(q, L, c, d):
    """검증기(pipenet_validator)와 같은 식이어야 한다 — 두 경로의 판정이 갈리면 안 됨."""
    from pipenet_validator import PipenetGuideValidator

    ref = PipenetGuideValidator._calc_hw_friction_loss_kgcm2(
        None, q_lpm=q, total_length_m=L, c_factor=c, actual_bore_mm=d)
    assert hazen_williams_drop_bar(q, L, c, d) == pytest.approx(ref, rel=1e-12)


def test_hazen_williams_degenerate_inputs_are_zero():
    assert hazen_williams_drop_bar(0, 3.0, 120, 27.5) == 0.0
    assert hazen_williams_drop_bar(80, 0.0, 120, 27.5) == 0.0
    assert hazen_williams_drop_bar(80, 3.0, 120, 0.0) == 0.0


def test_inner_diameter_uses_ks_table():
    """호칭경이 아니라 KS 실내경을 써야 한다 (25A → 27.5 mm)."""
    assert inner_diameter_mm(25) == pytest.approx(27.5)
    assert inner_diameter_mm(100) == pytest.approx(105.3)
    assert inner_diameter_mm(25) > 25


def test_velocity_formula():
    """v = Q/A. 27.5 mm 관에 80 L/min → 약 2.25 m/s."""
    area = math.pi / 4 * (0.0275 ** 2)
    assert velocity_mps(80, 27.5) == pytest.approx((80 / 60000.0) / area)
    assert velocity_mps(80, 0) == float("inf")


# ── 위상·역할 ───────────────────────────────────────────────────────────────


def test_role_classification_is_topological():
    """교차배관은 관경과 무관하게 10 m/s, 가지배관은 6 m/s."""
    nodes, pipes, nozzles = make_net()
    topo = build_topology(nodes, pipes, nozzles)
    roles = dict(zip((p["label"] for p in pipes), classify_pipe_roles(topo, pipes)))
    # 아래로 가지가 2 개 이상 달리는 구간 = 교차배관
    assert roles["pc0"] == "other"
    assert roles["pc1"] == "other"
    assert roles["pc2"] == "other"
    # 헤드만 매달린 단일 사슬 = 가지배관 (말단 교차구간 pc3 도 안전측으로 가지)
    assert roles["pb0_0"] == "branch"
    assert roles["pc3"] == "branch"


def test_role_falls_back_to_bore_without_nozzles():
    nodes, pipes, _ = make_net()
    for p in pipes:
        p["dia"] = 25 if p["label"].startswith("pb") else 100
    topo = build_topology(nodes, pipes, [])
    roles = dict(zip((p["label"] for p in pipes), classify_pipe_roles(topo, pipes)))
    assert roles["pb0_0"] == "branch"
    assert roles["pc0"] == "other"


def test_topology_orients_pipes_away_from_source():
    """in/out 기재 순서와 무관하게 source 로부터의 깊이로 상·하류가 잡혀야 한다."""
    nodes, pipes, nozzles = make_net(n_branch=1, n_head=2)
    pipes[0]["in"], pipes[0]["out"] = pipes[0]["out"], pipes[0]["in"]
    topo = build_topology(nodes, pipes, nozzles)
    assert topo.source == "S"
    assert topo.up[0] == "S" and topo.down[0] == "c0"


# ── 수렴 ────────────────────────────────────────────────────────────────────


def test_converge_eliminates_all_violations():
    nodes, pipes, nozzles = make_net()
    rep = converge_bores_by_velocity(nodes, pipes, nozzles)
    assert rep["violations_after"] == 0
    assert rep["converged"] is True
    assert rep["violations_before"] > 0
    for p in pipes:
        assert p["velocity_mps"] <= p["v_limit"] + 1e-9
        assert p["v_over"] is False


def test_converge_respects_role_limits():
    nodes, pipes, nozzles = make_net()
    converge_bores_by_velocity(nodes, pipes, nozzles)
    for p in pipes:
        expected = BRANCH_PIPE_V_LIMIT if p["pipe_role"] == "branch" else MAIN_PIPE_V_LIMIT
        assert p["v_limit"] == expected


def test_converge_keeps_bores_monotone():
    """상류 관경 >= 하류 관경 — 아니면 배관망이 성립하지 않는다."""
    nodes, pipes, nozzles = make_net()
    converge_bores_by_velocity(nodes, pipes, nozzles)
    topo = build_topology(nodes, pipes, nozzles)
    for i, p in enumerate(pipes):
        for k in topo.children.get(topo.down[i], ()):
            assert p["dia"] >= pipes[k]["dia"]


def test_converge_is_idempotent():
    """수렴한 망을 다시 돌리면 더 바뀔 것이 없어야 한다."""
    nodes, pipes, nozzles = make_net()
    converge_bores_by_velocity(nodes, pipes, nozzles)
    before = [p["dia"] for p in pipes]
    rep2 = converge_bores_by_velocity(nodes, pipes, nozzles)
    assert [p["dia"] for p in pipes] == before
    assert rep2["changed"] == 0


def test_converge_never_shrinks_existing_bores():
    nodes, pipes, nozzles = make_net()
    for p in pipes:
        p["dia"] = 100
    converge_bores_by_velocity(nodes, pipes, nozzles, keep_existing=True)
    assert all(p["dia"] >= 100 for p in pipes)


def test_converge_shrinks_when_keep_existing_off():
    """keep_existing=False 면 과대 관경을 유속 근거로 줄일 수 있어야 한다."""
    nodes, pipes, nozzles = make_net()
    for p in pipes:
        p["dia"] = 200
    converge_bores_by_velocity(nodes, pipes, nozzles, keep_existing=False)
    assert min(p["dia"] for p in pipes) < 200


def test_grossly_undersized_network_does_not_slam_to_max_bore():
    """전 구간 25A 같은 극단적 초기망에서도 과승급 없이 수렴해야 한다.

    얇은 망은 말단압 0.1 MPa 를 만족시킬 수 없어 해석이 발산한다(소스압 폭주).
    그 발산 유량을 그대로 믿고 승급하면 전 구간이 사다리 최대치로 튄다.
    """
    nodes, pipes, nozzles = make_net(n_branch=8, n_head=10, dia=25)
    rep = converge_bores_by_velocity(nodes, pipes, nozzles)
    assert rep["violations_after"] == 0
    # 유량에 따라 층층이 갈린 분포여야 한다 — 전 구간 최대경으로 튀면 1~2 종뿐.
    assert len({p["dia"] for p in pipes}) >= 6
    assert sum(1 for p in pipes if p["dia"] >= 150) <= len(pipes) // 10


def test_overdischarge_is_modeled():
    """말단 헤드만 80 L/min 이고 펌프에 가까운 헤드는 더 많이 토출해야 한다."""
    nodes, pipes, nozzles = make_net()
    rep = converge_bores_by_velocity(nodes, pipes, nozzles)
    assert rep["overdischarge_ratio_max"] > 1.0
    topo = build_topology(nodes, pipes, nozzles)
    sol = solve_network(pipes, topo)
    flows = sorted(sol.head_flow_lpm.values())
    assert flows[0] == pytest.approx(NOZZLE_K_FACTOR, rel=0.02)
    assert flows[-1] > flows[0]


def test_solver_pins_remotest_head_at_minimum():
    """설계 기준 역산 — 가장 먼 헤드의 압력이 NFTC 최소치(1.0 bar)."""
    nodes, pipes, nozzles = make_net()
    converge_bores_by_velocity(nodes, pipes, nozzles)
    topo = build_topology(nodes, pipes, nozzles)
    sol = solve_network(pipes, topo)
    head_p = [sol.node_pressure_bar[h] for h in topo.head_count]
    assert min(head_p) == pytest.approx(1.0, abs=1e-6)
    assert sol.source_pressure_bar > 1.0


def test_elevation_raises_required_source_pressure():
    """표고가 높아지면 소스 압력이 정수두만큼 더 필요하다."""
    nodes, pipes, nozzles = make_net(n_branch=1, n_head=2)
    topo = build_topology(nodes, pipes, nozzles)
    flat = solve_network(pipes, topo).source_pressure_bar
    for n in nodes:
        if n["label"] != "S":
            n["elevation"] = 10.0
    raised = solve_network(pipes, build_topology(nodes, pipes, nozzles)).source_pressure_bar
    assert raised - flat == pytest.approx(10.0 * 0.0980665, rel=0.05)


def test_demand_flows_accumulate_down_the_tree():
    nodes, pipes, nozzles = make_net(n_branch=2, n_head=3)
    topo = build_topology(nodes, pipes, nozzles)
    flows = dict(zip((p["label"] for p in pipes), demand_flows(topo)))
    assert flows["pc0"] == pytest.approx(80 * 6)   # 전체 6 헤드
    assert flows["pc1"] == pytest.approx(80 * 3)   # 두번째 가지 3 헤드
    assert flows["pb0_2"] == pytest.approx(80)     # 말단 헤드 1 개


def test_equipment_equivalent_length_increases_source_pressure():
    nodes, pipes, nozzles = make_net(n_branch=1, n_head=2)
    topo = build_topology(nodes, pipes, nozzles)
    base = solve_network(pipes, topo).source_pressure_bar
    with_eq = solve_network(pipes, topo, equiv_length_m={"pc0": 50.0}).source_pressure_bar
    assert with_eq > base


def test_empty_network_returns_neutral_report():
    rep = converge_bores_by_velocity([], [], [])
    assert rep["pipes"] == 0
    assert rep["converged"] is True
    assert rep["changed"] == 0


def test_network_without_nozzles_does_not_crash():
    nodes, pipes, _ = make_net(n_branch=1, n_head=2)
    rep = converge_bores_by_velocity(nodes, pipes, [])
    assert rep["pipes"] == len(pipes)
    assert rep["violations_after"] == 0
