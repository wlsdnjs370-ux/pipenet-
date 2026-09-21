# -*- coding: utf-8 -*-
"""통합(S700) 상태가 «슬롯에 딸리는가» — slots.py 는 안 딸린다고 적어 뒀다.

`_slot_capture` 는 SESSION_KEYS 밖 **전부** 를 슬롯 상태로 걷어간다(그 설계는
옳다 — 도면별 키를 열거하면 늘 때마다 빠뜨린다). 그런데 급수방식·펌프 제원·
결합 결과는 «세 슬롯의 결과를 모으는» 세션 전역 값이다. 둘 중 하나가 틀렸다.
"""
from __future__ import annotations

import importlib
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for p in (_ROOT, os.path.join(_ROOT, "core")):
    if p not in sys.path:
        sys.path.insert(0, p)
os.environ.setdefault("LOGIN_PASSWORD", "probe")
os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")

srv = importlib.import_module("대조 서버")
srv.app.config["TESTING"] = True
c = srv.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

from routes.module_f.jobs import _new_session      # noqa: E402

sess = _new_session(slot="plan")
sid = sess["id"]
print("세션", sid, "· 활성", sess["active"])

r = c.post("/api/module-f/merge/mode",
           json={"sid": sid, "mode": "hsp_pump", "source_drop_m": 3.5,
                 "pump": {"head_m": 80}})
print("급수방식 지정 :", r.status_code, r.get_json())

# 결합 결과가 있었던 것처럼 흉내 — 실제 결합은 세 도면이 필요하다.
sess["merged"] = {"fake": True}
sess["merge_summary"] = {"nodes": 123}
print("전환 전 세션 :", {k: sess.get(k) for k in
                     ("supply_mode", "source_drop_m", "pump_spec",
                      "merged", "merge_summary")})

r = c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "system"})
print("슬롯 전환    :", r.status_code, (r.get_json() or {}).get("switched"))
print("전환 후 세션 :", {k: sess.get(k) for k in
                     ("supply_mode", "source_drop_m", "pump_spec",
                      "merged", "merge_summary")})

st = c.get(f"/api/module-f/merge/state?sid={sid}").get_json() or {}
print("merge/state  :", {k: st.get(k) for k in
                         ("supply_mode", "mode", "ready")} or st)
print()
lost = [k for k in ("supply_mode", "source_drop_m", "pump_spec",
                    "merged", "merge_summary") if sess.get(k) is None]
print("★슬롯 전환으로 잃은 통합 상태:", lost or "없음")
