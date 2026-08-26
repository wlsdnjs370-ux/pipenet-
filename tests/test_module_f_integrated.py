# -*- coding: utf-8 -*-
"""[H-7] 통합 골든 — 계통도 실도면으로 결합한 결과를 고정한다.

여기서 지키는 것은 «수치가 안 변한다» 가 아니라 **어긋나면 알아차린다** 이다.
계통도 추출·결합·변환 어느 한 곳이 조용히 달라지면 이 골든이 깨진다.

골든 갱신(의도된 변경일 때만)::

    MODULE_F_GOLDEN_UPDATE=1 python -m pytest tests/test_module_f_integrated.py -q
    # (PowerShell) $env:MODULE_F_GOLDEN_UPDATE=1; python -m pytest ... -q

실도면이 없는 환경에서는 통째로 건너뛴다 — 없는 것과 틀린 것은 다르다.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

GOLDEN = _ROOT / "tests" / "module_f_integrated_golden.json"
UPDATE = os.environ.get("MODULE_F_GOLDEN_UPDATE") == "1"
SYSTEM_DXF = _ROOT / "data" / "sample_problem" / "대명동201동 계통도.dxf"
# 두 점은 배관 끝점의 좌표 극값으로 고른다 — 사람 클릭을 재현 가능하게 고정한다.
SNAP_MM = 5000.0


def _head_tables():
    """G 의 설계 표 모양 — 급수원 1(→10) · 헤드 둘."""
    class _G:
        nodes = [
            {"label": "1", "x": 0, "y": 0, "elevation": 0.0,
             "io_node": "Input", "pressure_pa": 101325.0},
            {"label": "2", "x": 3000, "y": 0, "elevation": 0.0, "io_node": "No"},
            {"label": "3", "x": 6000, "y": 0, "elevation": 0.0, "io_node": "No"},
            {"label": "4", "x": 6000, "y": 3000, "elevation": 0.0, "io_node": "No"},
        ]
        pipes = [
            {"label": "P1", "in": "1", "out": "2", "type": "KSD 3507",
             "dia": 65, "length": 3.0, "elev": 0.0, "c": 120,
             "status": "Normal", "group": "Unset"},
            {"label": "P2", "in": "2", "out": "3", "type": "KSD 3507",
             "dia": 40, "length": 3.0, "elev": 0.0, "c": 120,
             "status": "Normal", "group": "Unset"},
            {"label": "P3", "in": "3", "out": "4", "type": "KSD 3507",
             "dia": 25, "length": 3.0, "elev": 0.0, "c": 120,
             "status": "Normal", "group": "Unset"},
        ]
        nozzles = [
            {"label": "1", "in": "3", "out": "@/1", "status": "1",
             "lib": "SP-HEAD", "flow_lmin": 80, "flow_m3s": 80 / 60000.0},
            {"label": "2", "in": "4", "out": "@/2", "status": "1",
             "lib": "SP-HEAD", "flow_lmin": 80, "flow_m3s": 80 / 60000.0},
        ]
        fittings = [{"pipe": "P2", "in": "2", "out": "3",
                     "type": "Elbow 90", "count": 1}]
        equipment: list = []
        meta = [("제목", "통합 골든")]
    return _G()


@pytest.fixture(scope="module")
def measured():
    if not SYSTEM_DXF.is_file():
        pytest.skip(f"계통도 실도면 없음: {SYSTEM_DXF.name}")

    from routes.module_f.emit import cross_check, emit_merged
    from routes.module_f.merge import combined_summary, merge_network
    from routes.module_f.subdrawing import (entities_to_world, extract_system,
                                            parse_subdrawing)

    ents, _parsed = parse_subdrawing(SYSTEM_DXF)
    w = entities_to_world(ents)
    pts = [p for s in w.segs for p in (s[2], s[3])]
    lo = min(pts, key=lambda p: (p[0], p[1]))
    hi = max(pts, key=lambda p: (p[0], p[1]))
    riser = extract_system(ents, lo, hi, snap_tolerance_mm=SNAP_MM)

    got = merge_network(_head_tables(), riser=riser, mode="lsp_gravity")
    summary = combined_summary(got)

    tmp = Path(tempfile.mkdtemp(prefix="mf_golden_"))
    try:
        files = emit_merged(got["combined"], tmp, title="통합 골든")
        cc = cross_check(files)
        out = {
            "riser": {"nodes": len(riser["nodes"]),
                      "pipes": len(riser["pipes"]),
                      "av_node_label": str(riser.get("av_node_label"))},
            "combined": {"nodes": summary["nodes"], "pipes": summary["pipes"],
                         "nozzles": summary["nozzles"],
                         "attached": summary["attached"]},
            "formats": {k: {"total_m": v.get("total_m"),
                            "nozzles": v.get("nozzles")}
                        for k, v in cc["per_format"].items()
                        if "error" not in v},
            "agree": cc["agree"],
        }
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return out


def test_통합_골든(measured):
    if UPDATE:
        GOLDEN.write_text(
            json.dumps(measured, ensure_ascii=False, indent=2, sort_keys=True),
            encoding="utf-8")
        pytest.skip("골든 갱신됨")
    assert GOLDEN.exists(), (
        "골든 없음 — MODULE_F_GOLDEN_UPDATE=1 로 먼저 고정하세요")
    want = json.loads(GOLDEN.read_text(encoding="utf-8"))
    assert measured == want, "통합 산출이 골든과 다릅니다"


def test_기준점이_10이다(measured):
    """특허 S550 · S740 — 이 규약이 깨지면 결합 자체가 성립하지 않는다."""
    assert measured["riser"]["av_node_label"] == "10"


def test_절점은_합에서_공통_하나를_뺀_수다(measured):
    """S740 — 라이저의 10 과 헤드망의 10 은 한 절점이다."""
    assert measured["combined"]["nodes"] == measured["riser"]["nodes"] + 4 - 1


def test_세_형식이_같은_배관망이다(measured):
    """S760 — 절점 수가 아니라 총 연장·노즐이 불변량이다(S443)."""
    assert measured["agree"] is True
    assert set(measured["formats"]) == {"sdf", "kfp", "has"}
    lengths = {v["total_m"] for v in measured["formats"].values()}
    assert len(lengths) == 1, f"연장이 갈린다: {lengths}"
    nozzles = {v["nozzles"] for v in measured["formats"].values()}
    assert nozzles == {measured["combined"]["nozzles"]}
