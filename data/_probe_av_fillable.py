# -*- coding: utf-8 -*-
"""알람밸브 등가길이가 «화면에서 채울 수 있는» 자리로 나오는가 — 한 바퀴.

라이브러리에 없는 호칭경(15A)에 알람밸브가 놓인 판을 만들어:
  ① 미해결 쌍 목록에 뜨는가        (화면이 채울 자리를 여기서 만든다)
  ② 그 쌍을 채우면 실제로 풀리는가  (값이 표에 들어가고 미해결이 0 이 되는가)
"""
from __future__ import annotations

import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
if _G not in sys.path:
    sys.path.insert(0, _G)

from services.cad_import.design.tables import build_design_tables  # noqa: E402

net = {
    "pipe_data": {"P1": {"start": "N1", "end": "N2", "length_m": 4.0}},
    "nodes_meta_runtime": {
        "N1": {"coords": [1.0, 1.0, 0.0], "type_id": "pump"},
        "N2": {"coords": [5.0, 1.0, 0.0], "type_id": "base"},
    },
}
worst = {"heads": [], "loads": {}}
common = dict(bores={"P1": (15, "시험")}, valve_nodes=["N1"],
              tree_loads={"P1": 30})

t = build_design_tables(net, worst, {"P1": (0, 1)}, [], **common)
un = t.unresolved
print("① 채울 자리")
print(f"    미해결 개수  {dict(t.meta)['등가길이 미해결']}")
print(f"    쌍 목록      {un['pairs']}")
print(f"    자리 목록    {un['length_items']}")
av_pair = [p for p in un["pairs"] if p["kind"] == "alarm_valve"]
print(f"    → 알람밸브 쌍이 {'있다' if av_pair else '★없다(채울 길이 없다)'}")

print("\n② 그 쌍을 채우면")
t2 = build_design_tables(
    net, worst, {"P1": (0, 1)}, [],
    fitting_overrides={"eq_len": [{"kind": "alarm_valve", "dia": 15,
                                   "m": 6.5, "note": "KFI 자료"}]},
    **common)
row = [e for e in t2.equipment if e["desc"] == "A/V"][0]
print(f"    A/V 등가길이 {row['eq_len']} m · 근거 {row.get('eq_len_src')}")
print(f"    미해결 개수  {dict(t2.meta)['등가길이 미해결']}")
print(f"    쌍 목록      {t2.unresolved['pairs']}")
print(f"    meta 「직접 입력 — 등가길이」 "
      f"{dict(t2.meta).get('직접 입력 — 등가길이')}")
