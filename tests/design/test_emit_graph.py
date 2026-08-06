# -*- coding: utf-8 -*-
"""지시서 §9.6 C560 — 위상 확정과 모듈 A 스키마 방출.

여기서 지키는 것은 넷이다. **단위가 한 dict 안에서 갈린다**는 것(mm 와 m 를 섞으면
관경은 맞는데 배관장이 천 배가 된다), **매달린 낙차는 표시에만 있다**는 것(수리
표고에 넣으면 두 번 계상된다), **분기점이 섬으로 남지 않는다**는 것(남으면 그 아래
가지배관이 통째로 미방호인데 수리계산은 통과한다), 그리고 **부속류를 위상에서 다시
만든다**는 것(승계하면 굵기·형상을 고쳐도 부속류가 옛것으로 남는다).
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))

from core.design.deterministic import emit_graph as EG  # noqa: E402
from core.design.deterministic import pipe_routing as PR  # noqa: E402
from core.design.deterministic import pipe_sizing as PS  # noqa: E402
from core.design.deterministic.zoning import FLEX_SPEC  # noqa: E402

_STEP = 3200.0          # mm — 헤드 간격
_TEE_A = 8000.0         # mm — 분기점이 교차배관 한복판에 서는 자리
_FLOW = 80.0            # L/min — 헤드 하나의 방수량


def _axes():
    """theta=0 이라 project/world 가 항등이다 — 좌표를 눈으로 따라갈 수 있다."""
    return PR.ZoneAxes(theta=0.0, cross_span_mm=10000.0, branch_span_mm=6400.0)


def _riser(*, levels=("1F", "2F")):
    made = [PR.RiserLevel(floor=f, elevation_m=3.5 * i, display_z_m=3.5 * i)
            for i, f in enumerate(levels)]
    return PR.Riser(id="RS-C1", core_id="C1", point=(0.0, 0.0), levels=made)


def _plan(*, left=("H-1", "H-2"), right=("H-3",)):
    axes = _axes()
    cross = PR.CrossMain(id="CM-1", b_mm=0.0, a_span=(5000.0, 15000.0),
                         path=[[5000.0, 0.0], [15000.0, 0.0]])
    branch = PR.Branch(id="BR-1", cross_id="CM-1", a_mm=_TEE_A,
                       tee=(_TEE_A, 0.0), left=list(left), right=list(right))
    main = PR.MainRun(id="MN-Z1", zone_id="Z-1F-01", valve=(1000.0, 0.0),
                      path=[[1000.0, 0.0], [5000.0, 0.0]])
    return PR.ZonePlan(zone_id="Z-1F-01", axes=axes, crosses=[cross],
                       branches=[branch], main=main)


def _heads():
    return [{"id": "H-1", "x": _TEE_A, "y": _STEP},
            {"id": "H-2", "x": _TEE_A, "y": 2 * _STEP},
            {"id": "H-3", "x": _TEE_A, "y": -_STEP}]


def _conns(*, direct=("H-3",)):
    return [{"head_id": h["id"], "room_id": "R-1",
             "orientation": "upright" if h["id"] in direct else "pendent",
             "kind": "direct" if h["id"] in direct else "flex",
             "flex": None if h["id"] in direct else dict(FLEX_SPEC)}
            for h in _heads()]


def _topo(**kw):
    return EG.build_topology(_riser(), [{"plan": _plan(), "floor": "1F"}],
                             _heads(), _conns(), **kw)


def _constraints():
    return {"flow_lpm_min": _FLOW,
            "velocity_limit_mps": {"branch": 6.0, "other": 10.0},
            "cross_main_min_dn": 40}


def _sizes(topo):
    """C570 을 실제로 태워 관경을 얻는다 — 손으로 채우면 두 번째 진실이 된다."""
    heads_at = {node_id: 1 for node_id in topo.head_nodes.values()}
    loads, unreached = PS.edge_load(topo.adjacency(), topo.source, heads_at)
    assert unreached == []
    roles = {}
    for edge in loads:
        roles[edge] = topo.edges[EG._key(*edge)].role
    sizes, flags = PS.size_pipes(loads, roles, _constraints(),
                                 lambda dn: dn * 1.0)
    return sizes, flags


# ── 위상 확정 ────────────────────────────────────────────────────────────

def test_분기점은_교차배관_한복판에_끼워진다():
    """C530 이 직선 중간점을 지우므로 끼워 넣지 않으면 섬으로 남는다."""
    topo = _topo()
    tee = next(n for n in topo.nodes.values()
               if abs(n.x - _TEE_A) < 1.0 and abs(n.y) < 1.0)
    assert len(topo.adjacency()[tee.id]) >= 3
    # 원래의 통짜 교차배관 간선은 남아 있으면 안 된다 — 남으면 고리가 된다.
    assert not EG.orient(topo) or all(f["code"] != "OFF_TREE_EDGE"
                                      for f in topo.flags)


def test_모든_헤드가_급수원에서_닿는다():
    topo = _topo()
    EG.orient(topo)
    assert [f for f in topo.flags] == []
    assert len(topo.head_nodes) == 3


def test_하향식_헤드는_급수구와_따로_선다():
    """평면 좌표가 같다고 합치면 신축배관 등가길이 22.4 m 가 통째로 사라진다."""
    topo = _topo()
    flex_edges = [e for e in topo.edges.values() if e.flex is not None]
    assert len(flex_edges) == 2
    assert all(e.length_m == pytest.approx(FLEX_SPEC["physical_length_m"])
               for e in flex_edges)


def test_매달린_낙차는_표시에만_있다():
    """스키마 §2.4 — `elevation` 에 넣으면 낙차가 두 번 계상된다."""
    topo = _topo()
    head = topo.nodes[topo.head_nodes["H-1"]]
    branch_node = topo.nodes[next(
        b for b in topo.adjacency()[head.id])]
    assert head.elevation_m == pytest.approx(branch_node.elevation_m)
    assert head.display_z_m == pytest.approx(
        branch_node.display_z_m - FLEX_SPEC["physical_length_m"])


def test_상향식_헤드는_가지배관_그_자리다():
    """반자가 없으면 헤드가 가지배관에 직결한다 — 급수구를 따로 두면 없는 관이 생긴다."""
    topo = _topo()
    head = topo.nodes[topo.head_nodes["H-3"]]
    assert head.kind == "head"
    assert all(e.flex is None for key, e in topo.edges.items() if head.id in key)


def test_입상관은_길이가_0_이_아니다():
    """같은 (x, y) 에 서므로 평면만 재면 0 이 된다."""
    topo = _topo()
    key = EG._key("RS-C1-1F", "RS-C1-2F")
    assert topo.edges[key].length_m == pytest.approx(3.5)
    assert topo.edges[key].rise_m == pytest.approx(3.5)


def test_입상관에_없는_층은_채우지_않고_보고한다():
    """금지사항 G9 — 지어낸 층고는 그대로 낙차 압력이 된다."""
    topo = EG.build_topology(_riser(levels=("1F",)),
                             [{"plan": _plan(), "floor": "3F"}],
                             _heads(), _conns())
    codes = [f["code"] for f in topo.flags]
    assert codes == ["RISER_LEVEL_MISSING"]
    assert topo.flags[0]["heads"] == ["H-1", "H-2", "H-3"]


def test_접속방식이_없는_헤드는_배관에_붙지_않는다():
    topo = EG.build_topology(_riser(), [{"plan": _plan(), "floor": "1F"}],
                             _heads(), _conns()[:2])
    assert [f["code"] for f in topo.flags] == ["HEAD_CONNECTION_MISSING"]
    assert "H-3" not in topo.head_nodes


# ── 꺾임 흡수 ────────────────────────────────────────────────────────────

def _bent():
    """ㄱ 자로 꺾인 교차배관. 꼭짓점이 노드로 남으면 모듈 A 그래프와 형상이 다르다."""
    plan = _plan()
    plan.crosses[0].path = [[5000.0, 0.0], [15000.0, 0.0], [15000.0, 5000.0]]
    return EG.build_topology(_riser(), [{"plan": plan, "floor": "1F"}],
                             _heads(), _conns())


def test_통과점은_간선_하나로_합쳐지고_각도는_남는다():
    topo = _bent()
    before = len(topo.nodes)
    assert EG.absorb_bends(topo) >= 1
    assert len(topo.nodes) < before
    assert any(any(abs(a - 90.0) < 0.5 for a in e.elbows)
               for e in topo.edges.values())


def test_헤드와_밸브는_합쳐지지_않는다():
    topo = _bent()
    heads = set(topo.head_nodes.values())
    EG.absorb_bends(topo)
    assert heads <= set(topo.nodes)
    assert "AV-Z-1F-01" in topo.nodes


# ── 방출 ─────────────────────────────────────────────────────────────────

def _emit():
    topo = _topo()
    EG.absorb_bends(topo)
    sizes, size_flags = _sizes(topo)
    assert size_flags == []
    return topo, EG.emit_tables(topo, sizes, _constraints())


def test_단위는_한_dict_안에서_갈린다():
    """스키마 §6-2 — nodes.x/y 는 mm, nodes.elevation 과 pipes.length 는 m."""
    _, tables = _emit()
    riser_top = next(n for n in tables.nodes if n["design_id"] == "RS-C1-2F")
    assert riser_top["elevation"] == pytest.approx(3.5)
    assert riser_top["x"] == 0 and isinstance(riser_top["x"], int)

    cross = next(p for p in tables.pipes if p["role"] == "cross")
    # 교차배관 한 토막은 3 m 안팎이다. mm 로 새면 여기서 천 배로 튄다.
    assert 0.01 <= cross["length"] <= 20.0
    assert cross["dia"] in (25, 32, 40, 50, 65, 80, 100, 125, 150)


def test_배관장은_낙차보다_짧을_수_없다():
    _, tables = _emit()
    assert all(p["length"] >= abs(p["elev"]) - 1e-9 for p in tables.pipes)
    assert all(p["length"] >= 0.01 for p in tables.pipes)


def test_헤드마다_노즐이_하나씩_붙는다():
    topo, tables = _emit()
    assert len(tables.nozzles) == len(topo.head_nodes)
    node_labels = {n["label"] for n in tables.nodes}
    assert all(z["in"] in node_labels for z in tables.nozzles)
    # m³/s 는 L/min 에서 유도한다 — 손으로 자른 상수를 쓰면 되돌릴 때 어긋난다.
    assert all(z["flow_m3s"] == pytest.approx(z["flow_lmin"] / 60000.0)
               for z in tables.nozzles)


def test_신축배관은_등가길이로_실린다():
    """도면 물리길이(0.7 m)를 등가길이 칸에 쓰면 손실이 30 분의 1 로 준다."""
    _, tables = _emit()
    assert len(tables.equipment) == 2
    for eq in tables.equipment:
        assert eq["desc"] == "FX"
        assert eq["eq_len"] == pytest.approx(FLEX_SPEC["equivalent_length_m"])
        assert eq["drawing_len_mm"] == pytest.approx(
            FLEX_SPEC["physical_length_m"] * 1000.0)
        assert eq["source"] == "designed" and eq["override_flag"] is False


def test_FX_규격이름은_등록된_프로파일이어야_한다():
    """모르는 이름을 적으면 `override_flag` 없이는 FX 검토 게이트가 막는다."""
    from core.remote30_constants import FX_SPEC_PROFILES
    assert EG.FLEX_SPEC_REF in FX_SPEC_PROFILES
    profile = FX_SPEC_PROFILES[EG.FLEX_SPEC_REF]
    # 이름만 맞고 값이 다르면 화면과 SDF 가 서로 다른 등가길이를 말하게 된다.
    assert profile["eq_len_m"] == pytest.approx(FLEX_SPEC["equivalent_length_m"])
    assert int(profile["nominal_dn"]) == int(FLEX_SPEC["dn"])
    assert profile["inner_dia_mm"] == pytest.approx(FLEX_SPEC["inner_dia_mm"])


def test_라벨은_10부터_흐르는_순이다():
    """모듈 A 와 같은 규약. 급수원이 10 이어야 표를 위에서 아래로 읽으면 순서가 된다."""
    topo, tables = _emit()
    assert tables.labels[topo.source] == "10"
    assert next(n for n in tables.nodes
                if n["design_id"] == topo.source)["io_node"] == "Input"
    assert sum(1 for n in tables.nodes if n["io_node"] == "Input") == 1


# ── 부속류 재구성 ────────────────────────────────────────────────────────

def test_tee_는_들어오는_노드의_차수로_판정한다():
    """[문서정합] 모듈 A 의 실제 규칙(`:6160-6165`)이다 — 각도를 보지 않는다."""
    pipes = [{"label": "P1", "in": "10", "out": "11"},
             {"label": "P2", "in": "11", "out": "12"},
             {"label": "P3", "in": "11", "out": "13"}]
    tees = [f for f in EG.reconstruct_fittings(pipes, {}) if f["type"] == "tee"]
    # 11 은 차수 3(P1 in/out 합산 + P2 + P3)이라 그 노드에서 나가는 관이 분기다.
    assert {f["pipe"] for f in tees} == {"P2", "P3"}


@pytest.mark.parametrize("angle, expect", [
    (90.0, "elbow"), (45.0, "elbow-45"), (44.0, "elbow-45"),
    (70.0, "elbow"), (60.0, None), (1.0, None),
])
def test_꺾임_각도_구간은_모듈A와_같다(angle, expect):
    pipes = [{"label": "P1", "in": "10", "out": "11"}]
    got = [f["type"] for f in EG.reconstruct_fittings(pipes, {"P1": [angle]})]
    assert got == ([expect] if expect else [])


def test_부속류_이름은_HAS_변환기가_아는_것이어야_한다():
    """모르는 이름은 HAS 에서 조용히 사라진다 — 그만큼의 등가길이가 없어진다."""
    # `has_converter` 는 평면 import(`from hb_rules import ...`)라 그것이 놓인
    # 디렉터리 자체가 경로에 올라야 읽힌다. 배포판마다 `core/` 아래에도 루트에도
    # 놓이므로 둘 다 올리고 평면 이름으로 부른다. 설계 패키지가 이 짐을 지지
    # 않으려고 여기서만 손댄다.
    sys.path.insert(0, str(_ROOT / "core"))
    from has_converter import _FITTING_TO_CNT
    pipes = [{"label": "P1", "in": "10", "out": "11"},
             {"label": "P2", "in": "11", "out": "12"},
             {"label": "P3", "in": "11", "out": "13"}]
    fittings = EG.reconstruct_fittings(pipes, {"P1": [90.0, 45.0]})
    assert fittings
    assert all(f["type"] in _FITTING_TO_CNT for f in fittings)


def test_구간_밖_꺾임은_조용히_버리지_않는다():
    """46.5~70도 는 부속류가 없다. 그 등가길이가 빠졌다는 사실은 남겨야 한다."""
    topo = _topo()
    EG.absorb_bends(topo)
    next(iter(topo.edges.values())).elbows = [60.0]
    sizes, _ = _sizes(topo)
    tables = EG.emit_tables(topo, sizes, _constraints())
    assert [f["code"] for f in tables.flags] == ["BEND_NOT_A_FITTING"]


def test_관경이_없는_간선은_보고하고_뺀다():
    """금지사항 G9 — 굵기를 지어내면 그 굵기로 수리계산이 통과한다."""
    topo = _topo()
    EG.absorb_bends(topo)
    sizes, _ = _sizes(topo)
    # 신축배관 간선은 관경을 표가 아니라 규격에서 받으므로 지워도 티가 안 난다.
    sizes.pop(next(e for e in sizes if topo.edges[EG._key(*e)].flex is None))
    tables = EG.emit_tables(topo, sizes, _constraints())
    assert "PIPE_SIZE_MISSING" in [f["code"] for f in tables.flags]


def test_고리는_트리에_들지_않은_간선으로_드러난다():
    topo = _topo()
    a, b = "RS-C1-2F", "AV-Z-1F-01"
    topo.edges[EG._key(a, b)] = EG.Edge(role="main", length_m=1.0, rise_m=0.0)
    EG.orient(topo)
    assert "OFF_TREE_EDGE" in [f["code"] for f in topo.flags]
