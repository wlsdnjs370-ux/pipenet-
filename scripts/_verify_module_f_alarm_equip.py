# -*- coding: utf-8 -*-
"""[§29] 찍은 알람밸브가 «살아 있는 서버» 의 산출물에 실리는가.

시험은 합성 판으로 본다. 여기서는 띄워 둔 서버에 실도면을 태워, 사람이 손질에서
알람밸브를 찍었을 때 기기표에 A/V 가 서고 등가길이가 채워지는지 본다.

★오래 안 실렸던 이유가 «자리는 있는데 아무도 안 넘긴다» 였으므로, 검사도
  «넘겨졌나» 가 아니라 «산출물에 있나» 로 한다.
"""
from __future__ import annotations

import json
import os
import sys
import time
import urllib.parse
import urllib.request
import http.cookiejar

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
BASE = os.environ.get("MF_BASE", "http://127.0.0.1:5051")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)
fails: list[str] = []


def check(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '실패'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        fails.append(f"{name} — {detail}")


def _password():
    p = os.path.join(_ROOT, ".env")
    if os.path.isfile(p):
        for ln in open(p, encoding="utf-8"):
            if ln.startswith("LOGIN_PASSWORD="):
                return ln.split("=", 1)[1].strip()
    return os.environ.get("LOGIN_PASSWORD", "")


def main() -> int:
    cj = http.cookiejar.CookieJar()
    op = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))

    def get(p):
        return json.loads(op.open(BASE + p, timeout=900).read())

    def post(p, body):
        req = urllib.request.Request(
            BASE + p, data=json.dumps(body).encode(),
            headers={"Content-Type": "application/json"})
        try:
            return json.loads(op.open(req, timeout=900).read())
        except urllib.error.HTTPError as exc:
            try:
                return json.loads(exc.read())
            except Exception:
                return {"ok": False, "message": f"HTTP {exc.code}"}

    op.open(BASE + "/login",
            urllib.parse.urlencode({"password": _password()}).encode()).read()

    # ★값이 정확히 같을 때만 고른다 — 접두사로 고르면 찍기만 된 저장본이 걸린다.
    import glob
    from routes.module_f.common import IMPORT_WORK_ROOT
    ready = set()
    for f in glob.glob(str(IMPORT_WORK_ROOT / "DWG" / "*_유저손질.json")):
        try:
            j = json.load(open(f, encoding="utf-8"))
        except Exception:
            continue
        if j.get("sources") or j.get("valve_picks"):
            ready.add(os.path.basename(f)[: -len("_유저손질.json")])
    items = (get("/api/module-f/saved") or {}).get("items") or []
    pool = [i for i in items if i.get("source_exists") and i["key"] in ready]
    if not pool:
        print("손질까지 된 저장본이 없다 — 검사 불가")
        return 0
    key = pool[0]["key"]
    sid = (post("/api/module-f/reopen", {"key": key}) or {}).get("sid")
    print(f"[1] 세션 — {key}")
    check("열렸다", bool(sid), str(sid))
    if not sid:
        return 1

    def wait(limit=2400):
        for _ in range(int(limit / 0.5)):
            j = get(f"/api/module-f/job?sid={sid}")
            if j.get("state") in ("done", "error", "idle"):
                return j
            time.sleep(0.5)
        return {"state": "timeout"}

    wait()
    st = (get(f"/api/module-f/edit/state?sid={sid}") or {}).get("state") or {}
    valves = st.get("valves") or []
    srcs = st.get("sources") or []
    print(f"[2] 손질 — 알람밸브 {len(valves)}곳 · 접속점 {len(srcs)}곳")

    if not valves and srcs:
        # 옛 저장본은 급수원만 있다(리팩터링 7 이전). 사람이 화면에서 하듯
        # 그 자리에 알람밸브를 찍는다 — 원클릭이 바로 그 동작이다.
        # ★`sources` 는 노드 번호가 아니라 **좌표 쌍** 이다(views._edit_state).
        xy = srcs[0] if (srcs and isinstance(srcs[0], (list, tuple))
                         and len(srcs[0]) >= 2) else None
        if xy is None:
            print(f"  접속점 좌표를 못 읽었다({srcs[:1]}) — 건너뜀")
            return 0
        post("/api/module-f/edit/anchor-click",
             {"sid": sid, "x": xy[0], "y": xy[1]})
        j = wait()
        check("알람밸브 원클릭이 돈다", j.get("state") == "done",
              str(j.get("error") or j.get("state")))
        st = (get(f"/api/module-f/edit/state?sid={sid}") or {}).get("state") or {}
        valves = st.get("valves") or []
    check("알람밸브가 찍혀 있다", bool(valves), f"{len(valves)}곳")

    print("[3] 표 확정 → 기기표")
    post("/api/module-f/design/build", {"sid": sid, "k": 30})
    j = wait()
    check("표가 섰다", j.get("state") == "done", str(j.get("error")))
    d = get(f"/api/module-f/design/preview?sid={sid}")
    tables = d.get("tables") or {}
    eq = [e for e in (tables.get("equipment") or [])
          if str(e.get("desc")) == "A/V"]
    check("기기표에 A/V 가 있다", bool(eq), f"{len(eq)}행")
    if eq:
        row = eq[0]
        dia = {str(p.get("label")): p.get("dia")
               for p in (tables.get("pipes") or [])}.get(str(row.get("pipe")))
        check("등가길이가 0 이 아니다",
              float(row.get("eq_len") or 0.0) > 0.0,
              f"{row.get('eq_len')} m · 배관 {row.get('pipe')} {dia}A"
              f" · 근거 {row.get('eq_len_src')}")
        loads = {str(p.get("label")): p.get("load")
                 for p in ((d.get("view") or {}).get("pipes") or [])}
        check("물이 지나는 관에 붙었다",
              (loads.get(str(row.get("pipe"))) or 0) > 0,
              f"담당 헤드 {loads.get(str(row.get('pipe')))}")
    meta = dict((k, v) for k, v in (tables.get("meta") or []))
    check("등가길이 미해결이 안 늘었다", meta.get("등가길이 미해결") == "0",
          str(meta.get("등가길이 미해결")))

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("알람밸브 기기 — 살아 있는 서버에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
