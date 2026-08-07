# -*- coding: utf-8 -*-
"""Pump-fan ↔ SLF Pump-section 정합 — FNCADnet 수정작업 지시서 작업 2.

PIPENET 은 펌프 성능곡선이 정의돼 있지 않으면 양정을 스스로 선정한다. 그래서
<Pump-fan> 이 SLF 에 없는 Library-pump 를 가리켜도 계산은 돌아가고, 출력물만
봐서는 "실제로 살 수 있는 펌프인지" 와 무관한 계산서라는 것을 알 수 없다.
여기서 지키는 것은 두 가지다 — 정격유량·양정이 있으면 곡선을 만들어 넣고,
없으면 조용히 넘어가지 말고 미확정으로 드러낸다.

실행::

    python -m pytest tests/test_pump_library.py -q
"""
from __future__ import annotations

import sys
import xml.etree.ElementTree as ET
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pytest  # noqa: E402

from remote30_constants import DEFAULT_PUMP_LIBRARY_NAME  # noqa: E402
from remote30_full_network import (  # noqa: E402
    PUMP_OVERLOAD_HEAD_RATIO, PUMP_OVERLOAD_Q_RATIO, PUMP_SHUTOFF_HEAD_RATIO,
    BuildingPressureProfile, CombinedTables, FloorRow, ProjectContext, ZoneSpec,
    ZoneType, build_riser, emit_full_sdf,
)

RATED_Q_LPM = 2900.0
RATED_H_M = 162.0


def _profile() -> BuildingPressureProfile:
    return BuildingPressureProfile(building_name="시험동", floors=[
        FloorRow("옥상층", 4.0, 0.0),
        FloorRow("20층", 2.9, 8.0),
        FloorRow("18층", 2.9, 13.8, after_prv_m=2.0),
        FloorRow("16층", 2.9, 19.6, after_prv_m=7.8),
    ])


def _ctx(**pump) -> ProjectContext:
    return ProjectContext(
        zone_spec=ZoneSpec(zone_type=ZoneType.HSP_PUMP, target_floor="16층",
                           prv1_target_pa=700000.0, **pump),
        floor_profile=_profile(), project_title="펌프 라이브러리 시험")


def _emit(ctx: ProjectContext, out_dir: Path) -> tuple[Path, Path]:
    riser = build_riser(ctx)
    combined = CombinedTables(nodes=riser.nodes, pipes=riser.pipes,
                              pumps=riser.pumps, valves=riser.valves)
    sdf = out_dir / "pump_library.sdf"
    emit_full_sdf(combined, sdf, ctx=ctx)
    return sdf, sdf.with_suffix(".slf")


def _library_pumps(sdf: Path) -> set[str]:
    return {(el.text or "").strip()
            for el in ET.parse(sdf).getroot().iter("Library-pump")}


def _pump_definitions(slf: Path) -> dict[str, ET.Element]:
    sec = ET.parse(slf).getroot().find("Pump-section")
    assert sec is not None, "표준 SLF 에 Pump-section 이 있어야 한다"
    return {(pd.find("Item-name").text or "").strip(): pd
            for pd in sec.findall("Pump-definition")
            if pd.find("Item-name") is not None}


# ── 표준 SLF 의 실제 내용 ────────────────────────────────────────────────

def test_기본_펌프이름은_표준_SLF_에_없다(tmp_path):
    """기본값이 SLF 와 정합하다고 믿으면 안 된다는 사실 자체를 못 박아 둔다.

    SLF 사본마다 펌프 이름이 다르다. 언젠가 이 이름이 SLF 에 실리면 이 테스트가
    깨지는데, 그때는 아래 "주입" 테스트들의 전제를 다시 확인하라는 뜻이다.
    """
    ctx = _ctx()
    _, slf = _emit(ctx, tmp_path)
    assert DEFAULT_PUMP_LIBRARY_NAME not in _pump_definitions(slf)


# ── 정격값이 있으면 곡선을 만들어 넣는다 ─────────────────────────────────

def test_정격유량_양정이_있으면_곡선을_주입한다(tmp_path):
    ctx = _ctx(pump_rated_q_lpm=RATED_Q_LPM, pump_rated_h_m=RATED_H_M)
    sdf, slf = _emit(ctx, tmp_path)

    refs = _library_pumps(sdf)
    assert refs == {DEFAULT_PUMP_LIBRARY_NAME}
    assert refs <= set(_pump_definitions(slf))
    assert ctx.emit_findings == []


def test_주입된_곡선은_NFPC_3점(tmp_path):
    """체절(유량 0, 정격양정 140%) / 정격 / 150% 유량에서 정격양정 65%."""
    ctx = _ctx(pump_rated_q_lpm=RATED_Q_LPM, pump_rated_h_m=RATED_H_M)
    _, slf = _emit(ctx, tmp_path)
    pdef = _pump_definitions(slf)[DEFAULT_PUMP_LIBRARY_NAME]

    q_si = RATED_Q_LPM / 60000.0            # L/min → m³/s
    peak_q_si = q_si * PUMP_OVERLOAD_Q_RATIO
    assert float(pdef.get("min-flow")) == 0.0
    assert float(pdef.get("max-flow")) == pytest.approx(peak_q_si, rel=1e-6)

    points = pdef.find("Set-of-pump-points").findall("Pump-point")
    flows = [float(p.get("flow")) for p in points]
    assert flows == pytest.approx([0.0, q_si, peak_q_si], rel=1e-6)

    # 압력은 양정[m] × ρg — 비(比)로 보면 단위와 무관하게 3점 규약이 드러난다.
    heads = [float(p.get("pressure")) for p in points]
    assert heads[0] / heads[1] == pytest.approx(PUMP_SHUTOFF_HEAD_RATIO, rel=1e-3)
    assert heads[2] / heads[1] == pytest.approx(PUMP_OVERLOAD_HEAD_RATIO, rel=1e-3)


# ── 정격값이 없으면 조용히 넘어가지 않는다 ───────────────────────────────

def test_정격값이_없으면_미확정으로_올라온다(tmp_path):
    ctx = _ctx()
    sdf, slf = _emit(ctx, tmp_path)

    # 곡선을 지어내지 않았다 — SDF 는 참조하는데 SLF 에는 정의가 없다.
    assert _library_pumps(sdf) == {DEFAULT_PUMP_LIBRARY_NAME}
    assert DEFAULT_PUMP_LIBRARY_NAME not in _pump_definitions(slf)

    fields = {item["field"] for item in ctx.unconfirmed()}
    assert "pump_library_name" in fields
    line = next(ln for ln in ctx.warning_lines() if "pump_library_name" in ln)
    assert DEFAULT_PUMP_LIBRARY_NAME in line and "주입 실패" in line


def test_같은_항목은_한_줄만_올라온다(tmp_path):
    """Pump-fan 이 2대(1차·2차 부스터)여도 참조 이름이 같으면 경고는 하나."""
    ctx = _ctx()
    riser = build_riser(ctx)
    assert len(riser.pumps) == 2
    _emit(ctx, tmp_path)
    assert len(ctx.emit_findings) == 1
    assert len(ctx.warning_lines()) == len(ctx.unconfirmed())


# ── 입력 경로 ────────────────────────────────────────────────────────────

def test_폼과_JSON_왕복에서_정격값이_보존된다():
    from remote30_full_network import (project_context_from_form,
                                       zone_spec_from_form)
    spec = zone_spec_from_form({"zone_type": "hsp_pump",
                                "pump_rated_q_lpm": "2900",
                                "pump_rated_h_m": "162"})
    assert (spec.pump_rated_q_lpm, spec.pump_rated_h_m) == (2900.0, 162.0)

    ctx = ProjectContext(zone_spec=spec)
    back = ProjectContext.from_dict(ctx.to_dict())
    assert back.zone_spec.pump_rated_q_lpm == 2900.0
    assert back.zone_spec.pump_rated_h_m == 162.0
    assert project_context_from_form  # 폼 진입점이 살아 있는지만 확인


def test_빈칸은_0_이_아니라_None():
    """0 으로 때우면 "정격양정 0m 펌프" 라는 거짓 주장이 곡선까지 흘러간다."""
    from remote30_full_network import zone_spec_from_form
    spec = zone_spec_from_form({"zone_type": "hsp_pump",
                                "pump_rated_q_lpm": "", "pump_rated_h_m": "  "})
    assert spec.pump_rated_q_lpm is None
    assert spec.pump_rated_h_m is None
