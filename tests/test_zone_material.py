# -*- coding: utf-8 -*-
"""관종 구역(material_zones) 분리 — FNCADnet 작업지시서 모듈 A T2.

헤드 선정 영역(zones)과 관종 구역(material_zones)은 다른 개념이다.
예전엔 zones 를 그대로 CPVC 구역으로 넘겨서, 헤드를 고르려고 친 사각형이
지하주차장이든 복도든 전부 CPVC 로 찍혔다.

실행::

    python -m pytest tests/test_zone_material.py -q
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import remote30_prototype as rp  # noqa: E402
from remote30_constants import (  # noqa: E402
    CPVC_C_FACTOR, CPVC_PIPE_TYPE, STEEL_C_FACTOR, STEEL_PIPE_TYPE,
    ZONE_KIND_UNIT_DWELLING,
)

# AV(0,0) → 3-노드 직선 배관. 헤드 2개, 관로 2개.
AV = (0.0, 0.0)
MID = (2000.0, 0.0)
FAR = (4000.0, 0.0)

# MID~FAR 구간의 중점(3000,0)만 덮는 사각형 — AV~MID 중점(1000,0)은 밖.
FAR_HALF_RECT = (2500.0, -500.0, 4500.0, 500.0)


def _selection() -> rp.SelectionResult:
    heads = [rp.HeadCandidate(pos=MID, raw=MID, block_name="", layer="SP"),
             rp.HeadCandidate(pos=FAR, raw=FAR, block_name="", layer="SP")]
    return rp.SelectionResult(
        source_pos=AV,
        source_kind="manual",
        heads=heads,
        distances=[2000.0, 4000.0],
        edges=[(AV, MID, 2000.0), (MID, FAR, 2000.0)],
        nodes_in_subgraph=[AV, MID, FAR],
    )


def _pipes(material_zones=None):
    return rp.build_input_tables(_selection(), material_zones=material_zones).pipes


def test_no_zone_kind_means_steel_not_cpvc():
    """유형 미지정 도면에서 CPVC 관로는 0개다."""
    pipes = _pipes(None)
    assert pipes
    assert [p["type"] for p in pipes] == [STEEL_PIPE_TYPE] * len(pipes)
    assert [p["c"] for p in pipes] == [STEEL_C_FACTOR] * len(pipes)


def test_every_pipe_records_material_source():
    for p in _pipes(None):
        assert p["material_source"] == "default"
    for p in _pipes([{"rect": FAR_HALF_RECT, "kind": ZONE_KIND_UNIT_DWELLING}]):
        assert p["material_source"] in {"zone", "default"}


def test_unit_dwelling_zone_makes_cpvc_only_inside():
    by_source = {p["material_source"]: p for p in
                 _pipes([{"rect": FAR_HALF_RECT, "kind": ZONE_KIND_UNIT_DWELLING}])}
    assert by_source["zone"]["type"] == CPVC_PIPE_TYPE
    assert by_source["zone"]["c"] == CPVC_C_FACTOR
    assert by_source["default"]["type"] == STEEL_PIPE_TYPE


def test_parking_and_corridor_zones_stay_steel():
    for kind in ("parking", "corridor"):
        pipes = _pipes([{"rect": FAR_HALF_RECT, "kind": kind}])
        assert [p["type"] for p in pipes] == [STEEL_PIPE_TYPE] * len(pipes)
        assert {p["material_source"] for p in pipes} == {"zone", "default"}


def test_unknown_kind_falls_back_to_steel():
    """모르는 유형은 CPVC 로 찍지 않는다 — DXF 에 재질 정보가 없으므로 '모름'은 기본값이다."""
    pipes = _pipes([{"rect": FAR_HALF_RECT, "kind": "미확정"}])
    assert [p["type"] for p in pipes] == [STEEL_PIPE_TYPE] * len(pipes)
    assert [p["material_source"] for p in pipes] == ["default"] * len(pipes)


def test_cpvc_zones_parameter_is_gone():
    """헤드 영역을 그대로 CPVC 로 넘기던 옛 인자는 되살리지 않는다."""
    import inspect
    params = inspect.signature(rp.build_input_tables).parameters
    assert "cpvc_zones" not in params
    assert "material_zones" in params


def test_browser_and_server_agree_on_payload_key():
    """화면이 보내는 이름과 서버가 읽는 이름이 어긋나면 조용히 강관으로 떨어진다.

    구역 유형을 골라도 CPVC 가 안 나오는데 오류도 안 뜨는 형태라 눈치채기 어렵다.
    """
    import pytest

    html_path = _ROOT / "templates" / "remote30_prototype.html"
    if not html_path.exists():  # domain-slim 은 모듈 A 화면이 없다
        pytest.skip("모듈 A 템플릿 없음")
    assert html_path.read_text(encoding="utf-8").count("material_zones: materialZonesPayload()") == 2

    for name in ("r30_prototype.py", "r30_combined.py"):
        path = _ROOT / "routes" / name
        if not path.exists():  # domain-slim 은 라우트가 서버 파일에 인라인이다
            continue
        src = path.read_text(encoding="utf-8")
        assert 'get("material_zones")' in src
        assert "normalize_material_zones" in src


def test_material_zone_parser_drops_malformed():
    parsed = rp.normalize_material_zones([
        {"rect": [0, 0, 1, 1], "kind": ZONE_KIND_UNIT_DWELLING},
        {"rect": [0, 0, 1]},          # 좌표 부족
        {"kind": "parking"},          # rect 없음
        "unit_dwelling",              # dict 아님
        {"rect": [0, 0, 1, "x"], "kind": "parking"},  # 숫자 아님
        {"rect": [2, 2, 3, 3]},       # 유형 없음 — 유지(계산에서 기본값 처리)
    ])
    assert [z["kind"] for z in parsed] == [ZONE_KIND_UNIT_DWELLING, ""]
    assert parsed[0]["rect"] == (0.0, 0.0, 1.0, 1.0)
