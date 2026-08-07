# -*- coding: utf-8 -*-
"""라이저 형상 상수 외부화 — FNCADnet 작업지시서 4-1/4-3.

옥상 수평 길이 20.95 / 14.93, PRV 접근 0.5, T분기→알람밸브 1.0 은 대명동 201동
수작업 PIPENET 모델에서 나온 실측값이다. 코드에 박혀 있으면 다른 현장에서도
그 숫자가 조용히 나가므로 ProjectContext 로 올리고, 아무도 확정하지 않았으면
경고 블록이 그 사실을 말하게 한다.

실행::

    python -m pytest tests/test_riser_geometry.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pytest  # noqa: E402

from remote30_constants import (  # noqa: E402
    REDUCER_SNAP_TO_TEE_MM, RISER_PRV_APPROACH_M, RISER_ROOF_RUN_AFTER_DROP_M,
    RISER_ROOF_RUN_TO_RISER_M, TEE_TO_ALARM_VALVE_RISE_M, TEE_TO_ALARM_VALVE_RUN_M,
)
from remote30_full_network import (  # noqa: E402
    BuildingPressureProfile, FloorRow, ProjectContext, SourceTag, ZoneSpec, ZoneType,
    build_riser, count_reducers_snapped_to_tee, project_context_from_form,
)

GEOMETRY_FIELDS = (
    "roof_run_to_riser_m", "roof_run_after_drop_m", "roof_to_top_floor_drop_m",
    "tee_to_alarm_valve_m", "tee_branch_above_slab_m", "top_floor_extra_height_m",
    "water_hammer_arrester",
)


def _profile() -> BuildingPressureProfile:
    return BuildingPressureProfile(building_name="시험동", floors=[
        FloorRow("옥상층", 4.0, 0.0),
        FloorRow("20층", 2.9, 8.0),
        FloorRow("18층", 2.9, 13.8, after_prv_m=2.0),
        FloorRow("16층", 2.9, 19.6, after_prv_m=7.8),
    ])


def _ctx(**extra) -> ProjectContext:
    return ProjectContext(
        zone_spec=ZoneSpec(zone_type=ZoneType.LSP_1STAGE, target_floor="16층",
                           prv1_target_pa=700000.0),
        floor_profile=_profile(), **extra)


def _by_label(pipes: list[dict]) -> dict[str, dict]:
    return {str(p["label"]): p for p in pipes}


def test_아무도_확정하지_않으면_전부_미확정으로_올라온다():
    lines = _ctx().warning_lines()
    for name in GEOMETRY_FIELDS:
        assert any(line.startswith(f"[미확정] {name} ") for line in lines), name


def test_기본값은_대명동_실측값_그대로():
    """지금까지 나가던 숫자가 그대로 나가야 한다 — 이 단계는 리팩터링이다."""
    pipes = _by_label(build_riser(_ctx()).pipes)
    assert pipes["r1"]["length"] == pytest.approx(RISER_ROOF_RUN_TO_RISER_M)
    assert pipes["r3"]["length"] == pytest.approx(RISER_ROOF_RUN_AFTER_DROP_M)
    assert pipes["r5"]["length"] == pytest.approx(RISER_PRV_APPROACH_M)
    assert pipes["r7"]["elev"] == pytest.approx(TEE_TO_ALARM_VALVE_RISE_M)
    assert pipes["r7"]["length"] == pytest.approx(
        TEE_TO_ALARM_VALVE_RISE_M + TEE_TO_ALARM_VALVE_RUN_M)


def test_현장값을_넣으면_그_값으로_나간다():
    pipes = _by_label(build_riser(_ctx(
        roof_run_to_riser_m=8.4, roof_run_after_drop_m=3.2,
        tee_to_alarm_valve_m=0.7)).pipes)
    assert pipes["r1"]["length"] == pytest.approx(8.4)
    assert pipes["r3"]["length"] == pytest.approx(3.2)
    assert pipes["r7"]["elev"] == pytest.approx(0.7)
    assert pipes["r7"]["length"] == pytest.approx(0.7 + TEE_TO_ALARM_VALVE_RUN_M)


def test_알람밸브는_T분기점보다_위():
    """수작업 모델 전량이 rise=+1.0 — 부호가 뒤집히면 낙차 부호가 통째로 틀어진다."""
    riser = build_riser(_ctx())
    elev = {str(n["label"]): n["elevation"] for n in riser.nodes}
    assert elev["10"] > elev["5"]
    assert elev["10"] - elev["5"] == pytest.approx(TEE_TO_ALARM_VALVE_RISE_M)


def test_옥탑_낙차를_주면_압력표_층고_대신_그_값을_쓴다():
    """옥탑~최상층 낙차는 건물 고유값이라 도면 규칙으로 못 구한다."""
    default_top = {str(n["label"]): n["elevation"] for n in build_riser(_ctx()).nodes}["3"]
    assert default_top == pytest.approx(-4.0)      # 압력표 최상단 층고
    given = build_riser(_ctx(roof_to_top_floor_drop_m=3.75))
    assert {str(n["label"]): n["elevation"] for n in given.nodes}["3"] == pytest.approx(-3.75)


def test_폼에_적힌_형상만_사용자확정으로_찍힌다():
    form = {"zone_type": "lsp_1stage", "target_floor": "16층", "prv1_target_kgf": "7.0",
            "roof_run_to_riser_m": "8.4", "water_hammer_arrester": "1"}
    ctx = project_context_from_form(form)
    assert ctx.roof_run_to_riser_m == pytest.approx(8.4)
    assert ctx.tag("roof_run_to_riser_m") is SourceTag.USER_CONFIRMED
    assert ctx.water_hammer_arrester is True
    assert ctx.roof_run_after_drop_m == pytest.approx(RISER_ROOF_RUN_AFTER_DROP_M)
    assert ctx.tag("roof_run_after_drop_m") is SourceTag.DEFAULT


def test_직렬화_왕복에_형상이_실린다():
    ctx = _ctx(roof_run_to_riser_m=8.4, roof_to_top_floor_drop_m=3.75,
               water_hammer_arrester=True)
    back = ProjectContext.from_dict(ctx.to_dict())
    for name in GEOMETRY_FIELDS:
        assert getattr(back, name) == getattr(ctx, name), name


# ── 레듀서 T분기 귀속 (4-3) ────────────────────────────────────────────────

def _pipe(label: str, a: str, b: str, dia: int) -> dict:
    return {"label": label, "in": a, "out": b, "dia": dia}


def test_분기점에서_관경이_바뀌면_그_분기점에_귀속():
    nodes = [{"label": str(i), "x": i * 1000, "y": 0} for i in range(1, 5)]
    pipes = [_pipe("p1", "1", "2", 100), _pipe("p2", "2", "3", 80),
             _pipe("p3", "2", "4", 80)]
    assert count_reducers_snapped_to_tee(nodes, pipes) == 1


def test_분기점에서_먼_전이는_귀속되지_않는다():
    """전이 노드 3 은 분기점 2 에서 REDUCER_SNAP_TO_TEE_MM 밖."""
    far = REDUCER_SNAP_TO_TEE_MM * 3
    nodes = [{"label": "1", "x": 0, "y": 0}, {"label": "2", "x": 100, "y": 0},
             {"label": "3", "x": 100 + far, "y": 0}, {"label": "4", "x": 100, "y": 500},
             {"label": "5", "x": 100 + far, "y": 500}]
    pipes = [_pipe("p1", "1", "2", 100), _pipe("p2", "2", "4", 100),
             _pipe("p3", "2", "3", 100), _pipe("p4", "3", "5", 80)]
    assert count_reducers_snapped_to_tee(nodes, pipes) == 0

    near = [{**n, "x": 100 + REDUCER_SNAP_TO_TEE_MM / 2} if n["label"] in ("3", "5") else n
            for n in nodes]
    assert count_reducers_snapped_to_tee(near, pipes) == 1


def test_관경이_한_가지면_셀_것이_없다():
    nodes = [{"label": str(i), "x": i * 100, "y": 0} for i in range(1, 5)]
    pipes = [_pipe("p1", "1", "2", 100), _pipe("p2", "2", "3", 100),
             _pipe("p3", "2", "4", 100)]
    assert count_reducers_snapped_to_tee(nodes, pipes) == 0
