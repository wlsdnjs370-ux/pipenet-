# -*- coding: utf-8 -*-
"""FX 신축배관 실체화 — 헤드 노드가 진짜 말단(스퍼)이 되는지."""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
for p in (BASE, BASE / "core"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

from remote30_prototype import (  # noqa: E402
    PipeTables, _materialize_fx_pipes, _rewrite_slf_fx_schedules,
)
from remote30_constants import (  # noqa: E402
    FX_DEFAULT_PROFILE, FX_SPEC_PROFILES, fx_schedule_name,
)


def _tables():
    """가지관 중간에 헤드 H 가 놓인 최소망: A→H→B, H 에 노즐 + FX 등가길이."""
    nodes = [{"label": lbl, "x": x, "y": 0.0, "elevation": 0.0, "io_node": io}
             for lbl, x, io in (("A", 0.0, "Input"), ("H", 1.0, "No"), ("B", 2.0, "No"))]
    pipes = [{"label": "1", "in": "A", "out": "H", "type": "KSD 3507", "dia": 40,
              "length": 1.0, "elev": 0.0, "c": 120},
             {"label": "2", "in": "H", "out": "B", "type": "KSD 3507", "dia": 32,
              "length": 1.0, "elev": 0.0, "c": 120}]
    nozzles = [{"label": "1", "in": "H", "out": "@/1", "flow_lmin": 80}]
    equipment = [{"desc": "FX", "pipe": "1", "in": "A", "out": "H", "eq_len": 3.0}]
    fittings = [{"pipe": "2", "in": "H", "out": "B", "type": "elbow", "count": "1"}]
    return PipeTables(nodes=nodes, pipes=pipes, nozzles=nozzles,
                      fittings=fittings, equipment=equipment, meta=[])


def test_head_becomes_leaf_after_fx_materialize():
    t, sched, _geoms = _materialize_fx_pipes(_tables())
    fx_label = next(iter(sched))
    touching = [p for p in t.pipes if "H" in (p["in"], p["out"])]
    assert [p["label"] for p in touching] == [fx_label]


def test_sibling_pipe_moves_off_head_node():
    t, _sched, _geoms = _materialize_fx_pipes(_tables())
    fx = next(p for p in t.pipes if p["out"] == "H")
    sibling = next(p for p in t.pipes if p["label"] == "2")
    # 하류 배관은 헤드가 아니라 신축배관 상류 노드 F 에서 뻗어야 한다.
    assert sibling["in"] == fx["in"] != "H"


def test_fitting_endpoints_follow_moved_pipe():
    t, _sched, _geoms = _materialize_fx_pipes(_tables())
    sibling = next(p for p in t.pipes if p["label"] == "2")
    fitting = next(f for f in t.fittings if f["pipe"] == "2")
    assert (fitting["in"], fitting["out"]) == (sibling["in"], sibling["out"])


def test_nozzle_stays_on_head_node():
    t, _sched, _geoms = _materialize_fx_pipes(_tables())
    assert [nz["in"] for nz in t.nozzles] == ["H"]


def test_hanbaek_profile_registered_without_changing_default():
    """한백표준은 라이브러리에 추가만 한다 — 아무것도 안 고르면 지금까지와 같은 값이다."""
    assert FX_DEFAULT_PROFILE == "평균"
    assert FX_SPEC_PROFILES["한백표준"] == {
        "eq_len_m": 22.4, "nominal_dn": 25, "inner_dia_mm": 28.0,
        "c_factor": 120, "phys_len_m": 0.7,
    }
    assert fx_schedule_name(25, 28.0) == "FX_25A_28"


def _two_head_tables(spec_a, spec_b):
    """A 에서 헤드 2개가 갈라지는 최소망 — 헤드마다 다른 FX 규격을 붙인다."""
    nodes = [{"label": lbl, "x": x, "y": 0.0, "elevation": 0.0, "io_node": io}
             for lbl, x, io in (("A", 0.0, "Input"), ("H1", 1.0, "No"), ("H2", 2.0, "No"))]
    pipes = [{"label": "1", "in": "A", "out": "H1", "type": "KSD 3507", "dia": 40,
              "length": 1.0, "elev": 0.0, "c": 120},
             {"label": "2", "in": "A", "out": "H2", "type": "KSD 3507", "dia": 40,
              "length": 2.0, "elev": 0.0, "c": 120}]
    nozzles = [{"label": "1", "in": "H1", "out": "@/1", "flow_lmin": 80},
               {"label": "2", "in": "H2", "out": "@/2", "flow_lmin": 80}]
    equipment = [{"desc": "FX", "pipe": "1", "in": "A", "out": "H1",
                  "eq_len": FX_SPEC_PROFILES[spec_a]["eq_len_m"], "spec_ref": spec_a},
                 {"desc": "FX", "pipe": "2", "in": "A", "out": "H2",
                  "eq_len": FX_SPEC_PROFILES[spec_b]["eq_len_m"], "spec_ref": spec_b}]
    return PipeTables(nodes=nodes, pipes=pipes, nozzles=nozzles,
                      fittings=[], equipment=equipment, meta=[])


def test_same_profile_twice_makes_one_schedule():
    _t, sched, geoms = _materialize_fx_pipes(_two_head_tables("한백표준", "한백표준"))
    assert set(sched.values()) == {"FX_25A_28"}
    assert geoms == {"FX_25A_28": (25, 28.0, 120.0)}


def test_mixed_profiles_make_two_schedules():
    """한 도면에 두 규격이 섞여도 스케줄이 서로 덮어쓰지 않는다."""
    t, sched, geoms = _materialize_fx_pipes(_two_head_tables("평균", "한백표준"))
    assert set(geoms) == {"FX_20A_216", "FX_25A_28"}
    assert geoms["FX_20A_216"] == (20, 21.6, 120.0)
    bore_by_sched = {sched[p["label"]]: p["dia"] for p in t.pipes
                     if p["label"] in sched}
    assert bore_by_sched == {"FX_20A_216": 20, "FX_25A_28": 25}


def test_slf_gets_a_schedule_per_used_spec(tmp_path):
    """SLF 에 규격별 스케줄이 다 실려야 PIPENET 이 내경을 찾는다(없으면 Unset)."""
    import xml.etree.ElementTree as ET

    slf = tmp_path / "std.slf"
    slf.write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        "<Library><Schedule-section>"
        "<Schedule><Item-name>FX</Item-name></Schedule>"
        "</Schedule-section></Library>",
        encoding="utf-8",
    )
    _t, _sched, geoms = _materialize_fx_pipes(_two_head_tables("평균", "한백표준"))
    _rewrite_slf_fx_schedules(slf, geoms)

    sec = ET.parse(slf).getroot().find("Schedule-section")
    by_name = {s.findtext("Item-name"): s for s in sec.findall("Schedule")}
    assert set(by_name) == {"FX_20A_216", "FX_25A_28"}
    size = by_name["FX_25A_28"].find("Metric-definition/Size-definition")
    assert (size.get("internal"), size.get("nominal")) == ("28", "25")
