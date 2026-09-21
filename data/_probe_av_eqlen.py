# -*- coding: utf-8 -*-
"""알람밸브 A/V 기기행의 등가길이가 SDF 로 무엇이 나가는가 — 실측."""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
if _G not in sys.path:
    sys.path.insert(0, _G)

from services.cad_import.design.tables import build_design_tables

# 배관 하나 · 접속점 + 알람밸브 노드. 관경은 100A 로 준다.
net = {
    "pipe_data": {"P1": {"start": "N1", "end": "N2", "length_m": 2.0}},
    "nodes_meta_runtime": {
        "N1": {"coords": [0.0, 0.0, 0.0], "type_id": "pump"},
        "N2": {"coords": [2.0, 0.0, 0.0], "type_id": "base"},
    },
}
worst = {"heads": [], "loads": {(0, 1): 1}}
tbl = build_design_tables(net, worst, {"P1": (0, 1)}, [],
                          bores={"P1": (100, "시험")},
                          valve_nodes=["N2"])
for e in tbl.equipment:
    print(f"기기 {e['desc']} · 배관 {e['pipe']} · 관경 100A"
          f" · 등가길이 {e['eq_len']} m · 근거 {e.get('eq_len_src')}")
print("등가길이 미해결 =", dict(tbl.meta).get("등가길이 미해결"))

# 라이브러리에 값이 없는 호칭경(15A)이면 어떻게 되는가
tbl2 = build_design_tables(net, worst, {"P1": (0, 1)}, [],
                           bores={"P1": (15, "시험")},
                           valve_nodes=["N2"])
for e in tbl2.equipment:
    print(f"15A → 등가길이 {e['eq_len']} m · 근거 {e.get('eq_len_src')}")
print("15A 미해결 =", dict(tbl2.meta).get("등가길이 미해결"),
      "· 목록", tbl2.unresolved["length_items"])
