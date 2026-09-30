"""Analytical balance checks independent of the Flask/legacy tree solver."""
from dataclasses import replace
import math

import pytest

from src.pipenet_converter.hydraulics.models import BAR_PER_M, Network, Node, Nozzle, Pipe, Settings, Size
from src.pipenet_converter.hydraulics.solver import resistance, solve
from src.pipenet_converter.hydraulics.sizing import propose


def example(loop=False):
    sizes = (Size(25, 27.6), Size(32, 35.7), Size(40, 41.6), Size(50, 52.9), Size(65, 67.9))
    pipes = [Pipe("a", "s", "j", 40, 120, sizes), Pipe("b", "j", "h", 10, 120, sizes)]
    if loop:
        pipes.append(Pipe("c", "s", "j", 40, 120, sizes))
    return Network((Node("s", 0), Node("j", 3), Node("h", 3)), tuple(pipes),
                   (Nozzle("1", "h", 80, 80, 1, 12),), "s")


def test_analytical_single_line():
    net = example()
    loss = sum(resistance(p.length_m, p.sizes[0].inner_mm, 120) * (80/60)**1.852 for p in net.pipes)
    result = solve(net, (0, 0), 1 + (3+loss)*BAR_PER_M)
    assert result.converged
    assert result.nozzle_flow_lpm["1"] == pytest.approx(80, abs=.001)
    assert result.pressure_bar["h"] == pytest.approx(1, abs=1e-5)


def test_loop_preserves_parallel_flow_and_direction():
    net = example(True)
    result = solve(net, (0, 0, 0), 3)
    assert result.converged
    assert result.flow_lpm["a"] == pytest.approx(result.flow_lpm["c"], abs=1e-4)
    assert result.flow_lpm["a"]*2 == pytest.approx(result.flow_lpm["b"], abs=1e-4)
    reverse = replace(net, pipes=tuple(replace(p, a=p.b, b=p.a) for p in net.pipes))
    rev = solve(reverse, (0, 0, 0), 3)
    assert rev.converged
    for label in result.flow_lpm:
        assert rev.flow_lpm[label] == pytest.approx(-result.flow_lpm[label], abs=1e-4)


def test_grid_zero_cross_flow_and_mass_balance():
    size = (Size(32, 35.7),)
    net = Network(tuple(Node(n, 0) for n in "sabh"), tuple(
        Pipe(str(i), a, b, 10, 120, size) for i, (a,b) in enumerate(
        [("s","a"),("s","b"),("a","h"),("b","h"),("a","b")])
    ), (Nozzle("head", "h", 80,80,1,12),), "s")
    result = solve(net,(0,)*5,3)
    assert result.converged
    assert abs(result.flow_lpm["4"]) < 1e-3
    assert result.source_flow_lpm == pytest.approx(sum(result.nozzle_flow_lpm.values()), abs=1e-4)
    assert result.mass_error_lps <= 1e-6 and result.energy_error_m <= 1e-5


def test_pump_only_matches_analytic_duty():
    result = propose(example(), Settings(mode="pump_only"))
    assert result["feasible"]
    assert result["pump"]["flow_lpm"] == pytest.approx(80, abs=.002)
    assert result["pump"]["differential_head_m"] > 3 + 1/BAR_PER_M


def test_fixed_supply_enlarges_and_keeps_locked_pipe():
    net = example(True)
    net = replace(net, pipes=tuple(replace(p, locked=p.label=="b") for p in net.pipes))
    result = propose(net, Settings(mode="fixed_supply", supply_pressure_bar=1.6))
    assert result["feasible"]
    assert next(r for r in result["pipes"] if r["label"]=="b")["after_dn"] == 25
    assert net.pipes[0].initial == 0


def test_infeasible_locks_not_success():
    net = replace(example(), pipes=tuple(replace(p, locked=True) for p in example().pipes))
    result = propose(net, Settings(mode="fixed_supply", supply_pressure_bar=.5))
    assert result["status"] == "bounds_exhausted"
    assert not result["feasible"]


def test_disconnected_missing_bore_and_nan_rejected():
    with pytest.raises(ValueError, match="분리"):
        replace(example(), nodes=example().nodes+(Node("orphan",0),)).validate()
    with pytest.raises(ValueError):
        Settings(supply_pressure_bar=math.nan).validate()
    with pytest.raises(ValueError):
        replace(example(), pipes=(replace(example().pipes[0],sizes=(Size(25,0),)),)).validate()


def test_iteration_limit_has_calculated_sizes():
    result = propose(example(), Settings(mode="fixed_supply", supply_pressure_bar=.5, max_iterations=1))
    assert result["status"] == "iteration_limit"
    assert all(p["after_dn"]==25 for p in result["pipes"])


def test_solver_budget_is_not_success():
    assert not solve(example(), (0,0), 5, max_iterations=1).converged


def test_low_k_requires_more_than_one_bar_for_80_lpm():
    net=example()
    net=replace(net,nozzles=(replace(net.nozzles[0],k_lpm_sqrt_bar=40),))
    result=propose(net,Settings(mode='pump_only'))
    assert result['feasible']
    assert result['nozzles'][0]['pressure_bar']==pytest.approx(4,abs=.001)
    assert result['nozzles'][0]['flow_lpm']==pytest.approx(80,abs=.001)


def test_overpressure_not_marked_feasible():
    result=propose(example(),Settings(mode='fixed_supply',supply_pressure_bar=100))
    assert not result['feasible']
    assert any(v['kind']=='head_max' for v in result['violations'])


def test_pump_suction_and_discharge_velocity_heads_are_explicit():
    a=propose(example(),Settings(mode='pump_only'))
    b=propose(example(),Settings(mode='pump_only',suction_total_head_m=2,discharge_velocity_head_m=.3))
    assert a['pump']['differential_head_m']-b['pump']['differential_head_m']==pytest.approx(1.7)


def test_conservative_unknown_role_and_velocity_enlargement():
    net=example()
    settings=Settings(mode='joint',branch_velocity_mps=1,other_velocity_mps=10)
    result=propose(net,settings)
    assert result['feasible']
    assert all(r['velocity_mps']<=1.00001 for r in result['pipes'])
    assert any(r['after_dn']>r['before_dn'] for r in result['pipes'])


def test_cancellation_propagates():
    class Stop(BaseException): pass
    def stop(): raise Stop()
    with pytest.raises(Stop):
        propose(example(),Settings(),checkpoint=stop)


def test_noninteger_iteration_budget_rejected():
    with pytest.raises(ValueError):
        Settings(max_iterations=2.5).validate()
