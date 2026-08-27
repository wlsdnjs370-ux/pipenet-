# -*- coding: utf-8 -*-
"""[H-0] 도면 슬롯을 **실제 앱의 라우트로** 확인한다 (특허 S650).

단위 테스트는 slots.py 의 상태기계만 본다. 여기서는 `대조 서버.py` 를 그대로
올려 HTTP 계층까지 태운다 — 라우트 등록·로그인 게이트·JSON 규약이 실제로
맞물리는지는 이 층에서만 드러난다.

DXF 를 올리는 무거운 경로(`/slot/open`)는 여기서 돌리지 않는다. 엔진 부팅과
파싱이 수십 초라 검증 한 번에 묶기에 적절하지 않고, 그 경로는 평면도 열기와
**같은 잡**(`api_open._open_job`)이라 이미 덮여 있다. 대신 입력 검사만 본다.

    python scripts/_verify_module_f_slots.py
"""
from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FAILS: list[str] = []


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def _load_app():
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    os.environ.setdefault("DESIGN_WORKBENCH_ENABLED", "1")
    spec = importlib.util.spec_from_file_location(
        "server_app_slotverify", str(ROOT / "대조 서버.py"))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.app


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    app = _load_app()
    app.config["TESTING"] = True

    from routes.module_f.jobs import _new_session

    print("[H-0] 슬롯 라우트 실측 — 특허 S650")

    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True   # 로그인 게이트 통과 (routes/auth.py 규약)

        # ① 세션을 서버 안에서 직접 하나 만든다 — DXF 없이 슬롯만 본다.
        sess = _new_session()
        sess["key"] = "평면.dxf"
        sess["pick"] = object()
        sid = sess["id"]

        r = c.get(f"/api/module-f/slot/state?sid={sid}")
        check("GET /slot/state 200", r.status_code == 200, f"HTTP {r.status_code}")
        d = r.get_json() or {}
        check("응답이 세 슬롯", len(d.get("slots") or []) == 3,
              f"{len(d.get('slots') or [])}개")
        check("활성은 평면도", d.get("active") == "plan", str(d.get("active")))
        by = {x["kind"]: x for x in d.get("slots", [])}
        check("평면도가 찍기 단계로 보고", by.get("plan", {}).get("stage") == "pick",
              str(by.get("plan", {}).get("stage")))
        check("계통도는 비어 있다", by.get("system", {}).get("opened") is False)

        # ② 계통도로 바꾸면 평면도 상태가 따라오지 않는다.
        r = c.post("/api/module-f/slot/switch",
                   json={"sid": sid, "kind": "system"})
        check("POST /slot/switch 200", r.status_code == 200, f"HTTP {r.status_code}")
        d = r.get_json() or {}
        check("활성이 계통도", d.get("active") == "system", str(d.get("active")))
        by = {x["kind"]: x for x in d.get("slots", [])}
        check("계통도는 여전히 비었다", by.get("system", {}).get("opened") is False)
        check("평면도는 보존된다", by.get("plan", {}).get("key") == "평면.dxf",
              str(by.get("plan", {}).get("key")))

        # ③ 돌아오면 평면도가 그대로 살아 있다.
        r = c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "plan"})
        d = r.get_json() or {}
        by = {x["kind"]: x for x in d.get("slots", [])}
        check("평면도 복귀 — 단계 보존", by.get("plan", {}).get("stage") == "pick",
              str(by.get("plan", {}).get("stage")))

        # ④ 입력 검사.
        r = c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "없는것"})
        check("없는 도면 종류는 400", r.status_code == 400, f"HTTP {r.status_code}")
        r = c.post("/api/module-f/slot/switch", json={"sid": "없는세션", "kind": "plan"})
        check("만료 세션은 410", r.status_code == 410, f"HTTP {r.status_code}")
        r = c.post("/api/module-f/slot/open", data={"kind": "plan"})
        check("/slot/open 은 DXF 없으면 400", r.status_code == 400,
              f"HTTP {r.status_code}")

        # 올리기와 읽기가 갈렸다 — 안 올린 채로 읽으라면 막힌다.
        r = c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "auto"})
        check("안 올린 채 /slot/read 는 400", r.status_code == 400,
              f"HTTP {r.status_code}")
        # 평면도는 방식을 고르지 않으면 읽지 않는다(파서가 갈린다).
        sess["dxf"] = __import__("os").path.abspath(__file__)   # 존재하는 파일
        r = c.post("/api/module-f/slot/read", json={"sid": sid})
        check("평면도는 방식 없이 못 읽는다", r.status_code == 400,
              (r.get_json() or {}).get("message", "")[:50])
        r = c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "없는방식"})
        check("모르는 방식은 400", r.status_code == 400, f"HTTP {r.status_code}")
        sess["dxf"] = None

        # ⑤ 잡이 도는 중에는 슬롯을 못 바꾼다 — 워커가 남의 슬롯에 쓰는 것을 막는다.
        sess["job"] = {"state": "run", "phase": "도면 읽기", "started": 0.0,
                       "ended": None, "error": None, "result": None}
        r = c.post("/api/module-f/slot/switch", json={"sid": sid, "kind": "system"})
        check("작업 중 전환은 409", r.status_code == 409, f"HTTP {r.status_code}")
        sess["job"] = None

        # ⑥ 옛 엔드포인트가 그대로 있다 (하위호환).
        rules = {r.rule for r in app.url_map.iter_rules()}
        for path in ("/api/module-f/open", "/api/module-f/reopen",
                     "/api/module-f/world", "/api/module-f/slot/state",
                     "/api/module-f/slot/switch", "/api/module-f/slot/open",
                     "/api/module-f/slot/read"):
            check(f"라우트 존재 {path}", path in rules)

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 슬롯 라우트 전 항목 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
