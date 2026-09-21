# -*- coding: utf-8 -*-
"""`_verify_module_f.py` 4단이 «실패» 인가 «느림» 인가를 가른다.

그 스크립트의 `wait(..., limit=600)` 은 시간이 다 되면 `{"state":"timeout"}`
을 돌려주고, 부르는 쪽은 그것을 그냥 「변환 잡 실패」로 적는다 — 오류와
느림이 한 문장으로 뭉개진다. 여기서는 넉넉히 기다려 어느 쪽인지 본다.

이 길은 `reopen`(저장된 찍기) → `convert/run` 이라 업로드 코드에 닿지 않는다.
"""
from __future__ import annotations

import importlib.util
import json
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)

spec = importlib.util.spec_from_file_location("daejo", os.path.join(ROOT, "대조 서버.py"))
srv = importlib.util.module_from_spec(spec)
sys.modules["daejo"] = srv
spec.loader.exec_module(srv)
app = srv.app
app.config["TESTING"] = True

KEY = sys.argv[1] if len(sys.argv) > 1 else "B1F 현장조사 소화설비 평면도"
LIMIT = int(sys.argv[2]) if len(sys.argv) > 2 else 2400


def wait(c, sid, what, limit=LIMIT):
    t0 = time.time()
    last = ""
    while time.time() - t0 < limit:
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error"):
            print(f"   {what} → {j['state']} · {j['elapsed']}s", flush=True)
            if j["state"] == "error":
                print("   ", j.get("error"), flush=True)
            return j
        cur = (j.get("lines") or [""])[-1][:76]
        if cur != last:
            print(f"   … {int(time.time() - t0):4d}s | {cur}", flush=True)
            last = cur
        time.sleep(1.0)
    print(f"   {what} → 아직도 도는 중 ({limit}s 넘김)", flush=True)
    return {"state": "timeout"}


with app.test_client() as c:
    with c.session_transaction() as s:
        s["authed"] = True
    j = c.post("/api/module-f/reopen", json={"key": KEY}).get_json()
    assert j.get("ok"), j
    sid = j["sid"]
    if wait(c, sid, "배관망 열기")["state"] != "done":
        raise SystemExit("열기부터 안 된다")
    r = c.post("/api/module-f/convert/run",
               json={"sid": sid, "dto": {}, "outputs": {"full_kfp": True}})
    print("변환 요청:", r.get_json(), flush=True)
    jb = wait(c, sid, "수리계산 입력 변환")
    print("잡 상태:", jb.get("state"), flush=True)
    if jb.get("state") == "done":
        res = c.get(f"/api/module-f/convert/result?sid={sid}").get_json()["result"]
        print("변환 결과 ok:", res.get("ok"), flush=True)
        print(json.dumps(res.get("summary") or res.get("blockers"),
                         ensure_ascii=False)[:400], flush=True)
