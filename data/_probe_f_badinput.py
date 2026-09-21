# -*- coding: utf-8 -*-
"""모듈 F — «살아 있는 세션 + 잘못된 입력» 에서 500 이 나는가.

기존 안전망(`test_모든_라우트가_예외를_안_던진다`)은 세션 없이 두드리므로
대부분 410 에서 멈춘다 — **핸들러 본문을 안 지난다.** 여기서는 손질까지 간
진짜 세션에 이상한 값을 넣어 본문을 지나게 한다.

500 은 «사람이 읽을 수 없는 실패» 다. 이 저장소의 규약은 실패도 문장으로
말하는 것이므로, 500 이 나오면 그 자리가 결함이다.
"""
from __future__ import annotations

import importlib
import json
import os
import sys
import time

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for p in (_ROOT, os.path.join(_ROOT, "core")):
    if p not in sys.path:
        sys.path.insert(0, p)
os.environ.setdefault("LOGIN_PASSWORD", "probe")
os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")

srv = importlib.import_module("대조 서버")
srv.app.config["TESTING"] = True
app = srv.app
c = app.test_client()
with c.session_transaction() as s:
    s["authed"] = True


def wait(sid, limit=600):
    for _ in range(int(limit / 0.05)):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.05)
    return {"state": "timeout"}


# ── 손질까지 간 세션 하나
items = (c.get("/api/module-f/saved").get_json() or {}).get("items") or []
live = [it for it in items if it.get("source_exists")]
if not live:
    print("원본이 남은 저장본이 없다 — 검사 불가")
    raise SystemExit(1)
key = live[0]["key"]
sid = (c.post("/api/module-f/reopen", json={"key": key}).get_json() or {})["sid"]
print("세션", sid, "·", key, "·", wait(sid)["state"])

# ── 이상한 값들. «형식은 맞지만 뜻이 안 되는» 것을 고른다.
WEIRD = [
    {},
    {"x": "abc", "y": None, "max_d": []},
    {"x": 1e308, "y": -1e308, "max_d": 0},
    {"k": -5}, {"k": 10 ** 9}, {"k": "삼십"},
    {"zones": "not-a-list"}, {"zones": [[1, 2]]}, {"zones": [{}]},
    {"zones": [[float("inf"), 0, 1, 1]]},
    {"zones": [[float("nan"), 0, 1, 1]]},
    {"zones": [[1e308, -1e308, 1e308, -1e308]]},
    {"waypoints": [[float("inf"), 0]]}, {"waypoints": [[1e308, 1e308]]},
    {"pump_x": 1e308, "pump_y": 0, "av_x": 0, "av_y": 0},
    {"source_x": float("nan"), "source_y": 0, "conn_x": 0, "conn_y": 0},
    {"sheet": 99999}, {"source": "Z999"},
    {"eps_mm": -1}, {"eps_mm": "넓게"},
    {"mode": 123}, {"kind": []}, {"slot": {"a": 1}},
    {"heads": []}, {"heads": {"indices": ["x"]}}, {"heads": {"conf_min": "높음"}},
    {"dto": "문자열"}, {"outputs": "전부"},
    {"rows": "문자열"}, {"rows": [1, 2, 3]}, {"rows": [{"a": "x"}]},
    {"waypoints": "여기"}, {"waypoints": [[1]]},
    {"ceiling_m": "높이"}, {"snap_tolerance_mm": -3},
    {"layers": 5}, {"pump": "펌프"}, {"source_drop_m": "깊이"},
    {"method": []}, {"what": []},
]

rules = []
for r in app.url_map.iter_rules():
    if "/api/module-f/" not in r.rule or "<" in r.rule:
        continue
    if "POST" in r.methods:
        rules.append(("POST", r.rule))
    elif "GET" in r.methods:
        rules.append(("GET", r.rule))
rules.sort()
print(f"라우트 {len(rules)}개 × 입력 {len(WEIRD)}종\n")

# 무거운 잡을 띄우는 자리는 뺀다 — 검사가 몇 시간이 된다(그 자리들은
# 이미 잡 가드가 있고, 본문 검증은 잡 «앞» 에서 끝난다).
SKIP = {"/api/module-f/convert/run", "/api/module-f/design/build",
        "/api/module-f/auto/run", "/api/module-f/auto/network",
        "/api/module-f/auto/handoff", "/api/module-f/edit/autojoin/apply",
        "/api/module-f/edit/anchor-click", "/api/module-f/merge/build",
        "/api/module-f/merge/emit", "/api/module-f/pick/adopt",
        "/api/module-f/pick/suggest", "/api/module-f/open",
        "/api/module-f/slot/open", "/api/module-f/reopen",
        "/api/module-f/sub/extract", "/api/module-f/job/stream"}

crashes = []
for meth, rule in rules:
    if rule in SKIP:
        continue
    for w in WEIRD:
        body = dict(w)
        body["sid"] = sid
        try:
            if meth == "POST":
                rv = c.post(rule, json=body)
            else:
                rv = c.get(rule, query_string={k: json.dumps(v) if isinstance(
                    v, (list, dict)) else v for k, v in body.items()})
        except Exception as exc:  # noqa: BLE001 — 던지는 것 자체가 결함이다
            crashes.append((meth, rule, w, f"{type(exc).__name__}: {exc}"))
            continue
        if rv.status_code >= 500:
            msg = (rv.get_json() or {}).get("message") or rv.get_data(as_text=True)
            crashes.append((meth, rule, w, f"HTTP {rv.status_code} · {msg[:120]}"))
        # 잡이 떴으면 다음 검사를 막지 않게 기다린다
        if (c.get(f"/api/module-f/job?sid={sid}").get_json() or {}
                ).get("state") == "run":
            wait(sid, 120)

print(f"★500·예외 {len(crashes)}건")
seen = set()
for meth, rule, w, why in crashes:
    k = (rule, why[:60])
    if k in seen:
        continue
    seen.add(k)
    print(f"  {meth} {rule}\n     입력 {w}\n     → {why}")
