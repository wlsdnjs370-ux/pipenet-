# -*- coding: utf-8 -*-
"""라이저 표고 유도 — FNCADnet 작업지시서 모듈 A T5.

예전엔 대명동 201동 답안에서 뽑은 숫자(-98.45 / -73.35 / -3.75 …)가 코드에
박혀 있어서, 압력표를 안 올린 다른 현장도 그 표고로 조용히 계산됐다. 이제는
표에서 유도하고, 표가 없으면 무엇이 없는지 지목하며 멈춘다.

실행::

    python -m pytest tests/test_riser_elevation.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pytest  # noqa: E402

from remote30_full_network import (  # noqa: E402
    BuildingPressureProfile, FloorRow, MissingProjectInputError,
    ProjectContext, ZoneSpec, ZoneType, build_riser,
)

ROOF_H = 4.0     # 옥상층 층고 — 라이저가 내려오기 시작하는 표고의 근거
PRV1_PA = 700000.0
PRV2_PA = 350000.0


def _profile(rows: list[FloorRow]) -> BuildingPressureProfile:
    return BuildingPressureProfile(building_name="시험동", floors=rows)


def _one_stage_profile() -> BuildingPressureProfile:
    """옥상 → 20층 자연낙차 → 18층부터 1차 감압."""
    return _profile([
        FloorRow("옥상층", ROOF_H, 0.0),
        FloorRow("20층", 2.9, 8.0),
        FloorRow("19층", 2.9, 10.9),
        FloorRow("18층", 2.9, 13.8, after_prv_m=2.0),
        FloorRow("17층", 2.9, 16.7, after_prv_m=4.9),
        FloorRow("16층", 2.9, 19.6, after_prv_m=7.8),
    ])


def _two_stage_profile() -> BuildingPressureProfile:
    """1차 감압(18층~) → 2차 감압(지하1층~) — 감압이후 수두가 다시 낮아진다."""
    return _profile([
        FloorRow("옥상층", ROOF_H, 0.0),
        FloorRow("20층", 2.9, 8.0),
        FloorRow("18층", 2.9, 13.8, after_prv_m=2.0),
        FloorRow("16층", 2.9, 19.6, after_prv_m=7.8),
        FloorRow("지하1층", 3.5, 60.0, after_prv_m=1.5),
        FloorRow("지하2층", 3.5, 63.5, after_prv_m=5.0),
    ])


def _ctx(zone_type: ZoneType, target: str,
         profile: BuildingPressureProfile | None,
         *, prv2: float | None = None) -> ProjectContext:
    ctx = ProjectContext.titled("시험")
    ctx.zone_spec = ZoneSpec(zone_type=zone_type, target_floor=target,
                             prv1_target_pa=PRV1_PA, prv2_target_pa=prv2)
    ctx.floor_profile = profile
    return ctx


def _elev(riser, label: str) -> float:
    return next(n["elevation"] for n in riser.nodes if n["label"] == label)


def test_1차감압_표고가_전부_압력표에서_나온다():
    riser = build_riser(_ctx(ZoneType.LSP_1STAGE, "16층", _one_stage_profile()))
    assert _elev(riser, "3") == pytest.approx(-ROOF_H)     # 라이저 진입 = 옥상 층고
    assert _elev(riser, "8") == pytest.approx(-13.8)       # 감압 시작 = 18층
    assert _elev(riser, "10") == pytest.approx(-19.6)      # AV = 대상 층 16층


def test_수직관_길이가_표고차와_어긋나지_않는다():
    riser = build_riser(_ctx(ZoneType.LSP_1STAGE, "16층", _one_stage_profile()))
    by_label = {p["label"]: p for p in riser.pipes}
    assert by_label["r2"]["length"] == pytest.approx(ROOF_H)
    assert by_label["r4"]["length"] == pytest.approx(13.8 - ROOF_H)


def test_2차감압_표고는_감압이후수두가_다시_낮아지는_층():
    riser = build_riser(_ctx(ZoneType.LLSP_2STAGE, "지하2층", _two_stage_profile(),
                             prv2=PRV2_PA))
    assert _elev(riser, "8") == pytest.approx(-13.8)    # 1차 = 18층
    assert _elev(riser, "87") == pytest.approx(-60.0)   # 2차 = 지하1층
    assert _elev(riser, "10") == pytest.approx(-63.5)
    assert len(riser.valves) == 2


def test_자연낙차는_감압밸브_자리가_아예_없다():
    riser = build_riser(_ctx(ZoneType.LSP_GRAVITY, "20층", _one_stage_profile()))
    labels = {n["label"] for n in riser.nodes}
    assert labels.isdisjoint({"8", "89", "87", "88"})
    assert riser.valves == []
    assert _elev(riser, "10") == pytest.approx(-8.0)
    # 감압이 없으니 라이저는 옥상에서 AV 직전까지 한 번에 내려온다.
    r4 = next(p for p in riser.pipes if p["label"] == "r4")
    assert r4["length"] == pytest.approx(abs(-9.0 - -ROOF_H))


def test_펌프식은_수원_대신_펌프_토출이_라이저_머리():
    riser = build_riser(_ctx(ZoneType.HSP_PUMP, "16층", _one_stage_profile()))
    inputs = [n["label"] for n in riser.nodes if str(n.get("io_node")) == "Input"]
    assert inputs == ["100"]
    assert riser.pumps and riser.pumps[0]["in"] == "100"
    assert _elev(riser, "10") == pytest.approx(-19.6)


def test_압력표가_없으면_멈추고_무엇이_없는지_말한다():
    with pytest.raises(MissingProjectInputError) as exc:
        build_riser(_ctx(ZoneType.LSP_1STAGE, "16층", None))
    assert "압력표" in str(exc.value)


def test_대상층_행이_없으면_그_층을_지목한다():
    profile = _profile([FloorRow("옥상층", ROOF_H, 0.0), FloorRow("20층", 2.9, 8.0)])
    with pytest.raises(MissingProjectInputError) as exc:
        build_riser(_ctx(ZoneType.LSP_GRAVITY, "16층", profile))
    assert "16층" in str(exc.value)


def test_감압이후_열이_비면_감압밸브_표고를_지어내지_않는다():
    profile = _profile([FloorRow("옥상층", ROOF_H, 0.0), FloorRow("16층", 2.9, 19.6)])
    with pytest.raises(MissingProjectInputError) as exc:
        build_riser(_ctx(ZoneType.LSP_1STAGE, "16층", profile))
    assert "감압이후" in str(exc.value)


def test_2차_감압_구간을_못_찾으면_멈춘다():
    profile = _one_stage_profile()  # 감압 구간이 하나뿐 — 2차 시작점이 없다
    with pytest.raises(MissingProjectInputError) as exc:
        build_riser(_ctx(ZoneType.LLSP_2STAGE, "16층", profile, prv2=PRV2_PA))
    assert "2차 감압" in str(exc.value)


def test_옥상_층고가_비면_라이저_시작_표고를_지어내지_않는다():
    profile = _profile([FloorRow("옥상층", 0.0, 0.0), FloorRow("16층", 2.9, 19.6)])
    with pytest.raises(MissingProjectInputError) as exc:
        build_riser(_ctx(ZoneType.LSP_GRAVITY, "16층", profile))
    assert "층고" in str(exc.value)
