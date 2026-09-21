# -*- coding: utf-8 -*-
"""브라우저 흐름 그대로 — 평면도 먼저, 슬롯 바꾸고, 계통도. 어디서 500 인가."""
from __future__ import annotations

import importlib
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(_ROOT)
for p in (_ROOT, os.path.join(_ROOT, "core")):
    if p not in sys.path:
        sys.path.insert(0, p)
os.environ.setdefault("LOGIN_PASSWORD", "probe")
os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
os.environ["EXPOSE_TRACEBACK"] = "1"          # 실패 원문을 그대로 받는다

srv = importlib.import_module("대조 서버")
srv.app.config["TESTING"] = True
c = srv.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

PLAN = "routes/제출용[최종]/1. 입력도면 대명동 단위세대 평면도.dxf"
SUB = "data/uploads/1. 입력도면 대명동 단위세대 계통도.dxf"


def wait(sid, limit=900):
    for _ in range(limit * 20):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.05)
    return {"state": "timeout"}


def show(tag, rv):
    print(f"{tag}: {rv.status_code}")
    if rv.status_code >= 400:
        d = rv.get_json() or {}
        print("   메시지:", str(d.get("message"))[:200])
        tb = d.get("traceback")
        if tb:
            print("   ---- traceback 꼬리 ----")
            print("\n".join(tb.strip().splitlines()[-12:]))
    return rv


rv = show("① 평면도 업로드", c.post("/api/module-f/slot/open", data={
    "kind": "plan", "dxf_file": (open(PLAN, "rb"), os.path.basename(PLAN))},
    content_type="multipart/form-data"))
sid = (rv.get_json() or {}).get("sid")
print("   잡:", wait(sid)["state"])

show("② 계통도 슬롯으로 전환",
     c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "system"}))

rv = show("③ 계통도 업로드", c.post("/api/module-f/slot/open", data={
    "kind": "system", "sid": sid,
    "dxf_file": (open(SUB, "rb"), os.path.basename(SUB))},
    content_type="multipart/form-data"))
j = wait(sid)
print("   잡:", j["state"], (j.get("error") or "")[:160])

show("④ sub/state", c.get(f"/api/module-f/sub/state?sid={sid}"))
show("⑤ sub/graph", c.post("/api/module-f/sub/graph", json={"sid": sid}))
