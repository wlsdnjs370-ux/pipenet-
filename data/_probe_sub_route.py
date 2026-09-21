# -*- coding: utf-8 -*-
"""`/sub/graph` 가 실제로 도는가 + 레이어 고르기가 그래프를 바꾸는가."""
from __future__ import annotations

import importlib
import io
import json
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

srv = importlib.import_module("대조 서버")
srv.app.config["TESTING"] = True
c = srv.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

DXF = os.environ.get("MF_SUB_DXF",
                     "data/uploads/1. 입력도면 대명동 단위세대 계통도.dxf")
if not os.path.isfile(DXF):
    print("계통도 DXF 없음:", DXF)
    raise SystemExit(1)


def wait(sid, limit=600):
    for _ in range(limit * 20):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.05)
    return {"state": "timeout"}


rv = c.post("/api/module-f/slot/open", data={
    "kind": "system", "dxf_file": (open(DXF, "rb"), os.path.basename(DXF))},
    content_type="multipart/form-data")
sid = (rv.get_json() or {})["sid"]
print("열기:", rv.status_code, wait(sid)["state"])

r = c.post("/api/module-f/sub/graph", json={"sid": sid})
d = r.get_json() or {}
print("\n[전체 레이어]", r.status_code)
print(f"   노드 {len(d.get('nodes') or [])} · 간선 {len(d.get('edges') or [])}"
      f" · 추측연결 {d.get('forced')} · 성분 {d.get('components')}")
print(f"   JSON {len(json.dumps(d)) / 1024:.1f} KB")
lay = d.get("layers") or []
print(f"   레이어 {len(lay)}종 — 상위 5:")
for L in lay[:5]:
    print(f"      {L['cat']:<8} {L['n']:>6} 개  {L['layer']}")

pipe_layers = [L["layer"] for L in lay if L["cat"] == "PIPE"]
print(f"\n[PIPE 로 분류된 레이어만 {len(pipe_layers)}종]")
r2 = c.post("/api/module-f/sub/graph",
            json={"sid": sid, "layers": pipe_layers})
d2 = r2.get_json() or {}
print(f"   {r2.status_code} · 노드 {len(d2.get('nodes') or [])} "
      f"· 간선 {len(d2.get('edges') or [])} · 고른 것 {d2.get('chosen')}")

# 세션에 남아 추출까지 가는가
r3 = c.get(f"/api/module-f/sub/state?sid={sid}")
print("\n/sub/state:", r3.status_code, r3.get_json())
