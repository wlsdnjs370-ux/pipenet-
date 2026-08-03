# -*- coding: utf-8 -*-
"""FX 신축배관 실체화 — 헤드 노드가 진짜 말단(스퍼)이 되는지."""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
for p in (BASE, BASE / "core"):
    if str(p) not in sys.path:
        sys.path.insert(0, str(p))

from remote30_prototype import PipeTables, _materialize_fx_pipes  # noqa: E402


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
