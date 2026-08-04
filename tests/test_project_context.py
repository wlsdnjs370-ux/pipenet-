# -*- coding: utf-8 -*-
"""ProjectContext — FNCADnet 작업지시서 모듈 A T4.

예전엔 폼 딕셔너리(``job["spec_form"]``)를 라이저 생성 함수까지 통째로 들고
다녔다. 그래서 "사용자가 4.5 m 라고 확정한 천장고" 와 "아무도 안 정해서 코드
기본값이 쓰인 천장고" 가 구분되지 않았고, 리포트는 둘 다 똑같이 숫자로만 찍혔다.
ProjectContext 는 값과 출처를 같이 들고 다녀 후자를 [미확정] 로 드러낸다.

실행::

    python -m pytest tests/test_project_context.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pytest  # noqa: E402

from remote30_constants import FX_DEFAULT_PROFILE  # noqa: E402
from remote30_full_network import (  # noqa: E402
    ProjectContext, SourceTag, ZoneType,
    project_context_from_form, project_context_from_job,
)

# 낙차압 5 m 는 헤드 최소 방수압(0.1 MPa ≒ 10.2 m)에 못 미치고, 12 m 는 넘는다.
PRESSURE_TABLE = '[{"floor_label": "49F", "height_m": 2.9, "head_drop_m": 5.0},' \
                 ' {"floor_label": "26F", "height_m": 2.9, "head_drop_m": 12.0}]'

BASE_FORM = {
    "zone_type": "lsp_1stage",
    "target_floor": "16층",
    "prv1_target_kgf": "7.0",
    "pressure_table_json": PRESSURE_TABLE,
}


def _ctx(**extra) -> ProjectContext:
    return project_context_from_job({"spec_form": {**BASE_FORM, **extra}})


def test_빈칸은_사용자확정이_아니다():
    ctx = _ctx()
    fields = {item["field"] for item in ctx.unconfirmed()}
    assert "project_title" in fields
    assert "machine_room_ceiling_m" in fields
    assert ctx.tag("project_title") is SourceTag.DEFAULT


def test_채운_값만_사용자확정으로_찍힌다():
    ctx = _ctx(project_title="대명동 201동", machine_room_ceiling_m="4.5")
    assert ctx.tag("project_title") is SourceTag.USER_CONFIRMED
    assert ctx.machine_room_ceiling_m == 4.5
    fields = {item["field"] for item in ctx.unconfirmed()}
    assert "project_title" not in fields
    assert "zone_name" in fields


def test_숫자가_아니면_확정으로_치지_않는다():
    ctx = _ctx(machine_room_ceiling_m="사층")
    assert ctx.machine_room_ceiling_m is None
    assert ctx.tag("machine_room_ceiling_m") is SourceTag.DEFAULT


def test_모르는_FX규격은_기본값으로_떨어지고_미확정으로_남는다():
    ctx = _ctx(fx_profile_key="FX_없는규격")
    assert ctx.fx_profile_key == FX_DEFAULT_PROFILE
    assert ctx.tag("fx_profile_key") is SourceTag.DEFAULT


def test_경고줄은_미확정_항목마다_한_줄():
    ctx = _ctx()
    lines = ctx.warning_lines()
    assert len(lines) == len(ctx.unconfirmed())
    assert all(line.startswith("[미확정] ") for line in lines)
    assert any("프로젝트명" in line for line in lines)


def test_자연낙차_후보는_계산만_하고_확정하지_않는다():
    ctx = _ctx()
    cands = ctx.natural_fall_candidates()
    assert [c["floor_label"] for c in cands] == ["49F", "26F"]
    assert [c["ok"] for c in cands] == [False, True]
    assert cands[0]["required_m"] == pytest.approx(10.2, abs=0.05)
    # 후보는 있어도 확정값은 비어 있다 — 자동으로 고르지 않는다.
    assert ctx.suggested_natural_fall_floor() == "26F"
    assert ctx.natural_fall_start_floor is None


def test_압력표가_없으면_후보도_없다():
    ctx = project_context_from_job({"spec_form": {
        "zone_type": "lsp_1stage", "target_floor": "16층", "prv1_target_kgf": "7.0"}})
    assert ctx.natural_fall_candidates() == []
    assert ctx.suggested_natural_fall_floor() is None


def test_context_override_가_run폼을_덮는다():
    job = {"spec_form": {**BASE_FORM, "natural_fall_start_floor": "49F"},
           "context_override": {"natural_fall_start_floor": "26F"}}
    ctx = project_context_from_job(job)
    assert ctx.natural_fall_start_floor == "26F"
    assert ctx.tag("natural_fall_start_floor") is SourceTag.USER_CONFIRMED


def test_리포트_제목_우선순위():
    assert _ctx(project_title="대명동 201동", zone_name="저층부").report_title() == "대명동 201동"
    assert _ctx(zone_name="저층부").report_title() == "저층부"
    assert _ctx().report_title() == "Remote 30 전체 — lsp_1stage 16층"


def test_직렬화_왕복():
    ctx = _ctx(project_title="대명동 201동", zone_name="저층부",
               machine_room_ceiling_m="4.5", roof_tank_water_level_m="3.2")
    back = ProjectContext.from_dict(ctx.to_dict())
    assert back.report_title() == ctx.report_title()
    assert back.machine_room_ceiling_m == 4.5
    assert back.natural_fall_candidates() == ctx.natural_fall_candidates()
    assert back.warning_lines() == ctx.warning_lines()
    assert back.zone_spec.prv1_target_pa == ctx.zone_spec.prv1_target_pa


def test_titled_는_제목만_확정하고_나머지는_전부_미확정():
    ctx = ProjectContext.titled("Combined")
    assert ctx.report_title() == "Combined"
    assert ctx.zone_spec.zone_type is ZoneType.LSP_1STAGE
    fields = {item["field"] for item in ctx.unconfirmed()}
    assert "project_title" not in fields
    assert "floor_profile" in fields


def test_구역은_비어_있으면_확정으로_치지_않는다():
    empty = project_context_from_form(BASE_FORM, material_zones=[])
    assert empty.tag("material_zones") is SourceTag.DEFAULT
    filled = project_context_from_form(
        BASE_FORM, material_zones=[{"rect": (0, 0, 1, 1), "kind": "unit_dwelling"}])
    assert filled.tag("material_zones") is SourceTag.USER_CONFIRMED
