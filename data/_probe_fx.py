# -*- coding: utf-8 -*-
"""[§29] FX 를 켜면 산출물이 어떻게 달라지는가 — 끄면 한 바이트도 안 바뀌는가."""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_R = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_R, os.path.join(_R, "cad_project_editor_g"), os.path.join(_R, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
from services.cad_import.design.tables import build_design_tables

def net(kinds):
    nodes = {"N1": {"coords": [1.0, 1.0, 0.0], "type_id": "pump"},
             "N2": {"coords": [5.0, 1.0, 0.0], "type_id": "base"}}
    pipes = {"P1": {"start": "N1", "end": "N2", "length_m": 4.0}}
    nhk = {}
    for i, k in enumerate(kinds, start=1):
        nid = f"H{i}"
        nodes[nid] = {"coords": [5.0 + i, 1.0, 0.6], "type_id": "head"}
        pipes[f"PH{i}"] = {"start": "N2", "end": nid, "length_m": 0.7}
        nhk[nid] = k
    return {"pipe_data": pipes, "nodes_meta_runtime": nodes,
            "node_head_kinds": nhk}

KINDS = ["하향식", "하향식", "상향식", "상하향식"]
n = net(KINDS)
w = {"heads": [], "loads": {}}
com = dict(bores={k: (25, "시험") for k in n["pipe_data"]},
           tree_loads={"P1": 4})
off = build_design_tables(n, w, {}, [], **com)
on = build_design_tables(n, w, {}, [], fx_profile="평균", **com)
print(f"헤드 종류 {KINDS}")
print(f"  끄면  기기 {len(off.equipment)}개 · meta 신축배관 = {dict(off.meta).get('신축배관(FX)')}")
print(f"  켜면  기기 {len(on.equipment)}개 · meta 신축배관 = {dict(on.meta).get('신축배관(FX)')}")
for e in on.equipment:
    print(f"     {e['desc']} · 배관 {e['pipe']} · {e['eq_len']}m · {e.get('eq_len_src')}")
print(f"\n  ★끈 상태의 표가 종전과 같은가 — 노드 {len(off.nodes)} 배관 {len(off.pipes)}"
      f" 노즐 {len(off.nozzles)} 부속 {len(off.fittings)} 기기 {len(off.equipment)}")
try:
    build_design_tables(n, w, {}, [], fx_profile="없는규격", **com)
except ValueError as e:
    print(f"\n  모르는 규격은 거절: {str(e)[:70]}")
