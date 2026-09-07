# -*- coding: utf-8 -*-
"""[§29] 신축배관(FX) 설정이 «살아 있는 서버» 끝까지 가는가.

시험은 합성 판으로 본다(`tests/test_g_fx_equipment.py`). 여기서 보는 것은 셋:

  ⑴ 화면에서 고른 값이 표를 만드는 데까지 실제로 간다 — 오래 안 실렸던 이유가
     「자리는 있는데 아무도 안 넘긴다」였으므로, 검사도 «넘겼나» 가 아니라
     **산출물이 그렇게 말하나** 로 한다.
  ⑵ 켜지 않으면 한 줄도 안 바뀐다(D-F11-1).
  ⑶ 붙인 것을 **화면에서 볼 자리**가 있다 — 기기표.

★이 도면(대명동)은 헤드가 전부 상향식이라 켜도 **0개가 맞다**. 그러니 0 을
  「고장」으로 읽지 않도록, 산출물이 그 사유를 말하는지까지 본다.
"""
from __future__ import annotations

import glob
import json
import os
import sys
import time
import urllib.error
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
    print(f"  [{'OK  ' if ok else '실패'}] {name}"
          + (f" · {detail}" if detail else ""))
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

    def build(fx):
        body = {"sid": sid, "k": 30}
        if fx is not None:
            body["fx_profile"] = fx
        post("/api/module-f/design/build", body)
        j = wait()
        if j.get("state") != "done":
            return None, j
        d = get(f"/api/module-f/design/preview?sid={sid}")
        return d, j

    print("[2] 안 켰을 때 — 기본")
    d0, j0 = build("")
    check("표가 섰다", d0 is not None, str(j0.get("error")))
    if d0 is None:
        return 1
    t0 = d0.get("tables") or {}
    m0 = dict((k, v) for k, v in (t0.get("meta") or []))
    eq0 = t0.get("equipment") or []
    fx0 = [e for e in eq0 if str(e.get("desc")) == "FX"]
    check("meta 가 «안 함» 이라고 말한다", m0.get("신축배관(FX)") == "안 함",
          str(m0.get("신축배관(FX)")))
    check("FX 가 한 개도 없다", not fx0, f"{len(fx0)}개")

    print("[3] 켰을 때 — 평균(20A · 15.6m)")
    d1, j1 = build("평균")
    check("표가 섰다", d1 is not None, str(j1.get("error")))
    if d1 is None:
        return 1
    t1 = d1.get("tables") or {}
    m1 = dict((k, v) for k, v in (t1.get("meta") or []))
    note = str(m1.get("신축배관(FX)") or "")
    eq1 = t1.get("equipment") or []
    fx1 = [e for e in eq1 if str(e.get("desc")) == "FX"]
    check("고른 값이 산출물까지 간다", note.startswith("평균"), note)
    # 0 이면 «왜 0 인지» 를 말해야 한다 — 이 도면은 헤드가 전부 상향식이다.
    if fx1:
        check("등가길이가 규격대로다",
              all(abs(float(e.get("eq_len") or 0) - 15.6) < 1e-9 for e in fx1),
              f"{len(fx1)}개")
        check("무엇이 붙였는지 행이 말한다",
              all(str(e.get("eq_len_src")) == "신축배관 평균" for e in fx1))
    else:
        check("0 개인 사유를 산출물이 말한다", "없음" in note, note)

    print("[4] 배관은 안 건드린다 — 켠 것은 «기기» 뿐")
    p0 = [(r.get("label"), r.get("dia"), r.get("length")) for r in
          (t0.get("pipes") or [])]
    p1 = [(r.get("label"), r.get("dia"), r.get("length")) for r in
          (t1.get("pipes") or [])]
    check("배관표가 그대로다", p0 == p1,
          f"{len(p0)}행 → {len(p1)}행")

    print("[5] 화면에 볼 자리가 있다 — 기기표")
    page = op.open(BASE + "/module-f").read().decode("utf-8", "replace")
    check("표 고르개에 «기기» 가 있다", '<option value="equipment">' in page)
    # ★«몇 행인가» 로 판정하지 않는다. 기기표의 내용은 사람이 무엇을 찍었느냐에
    #   달렸다 — 알람밸브를 안 찍고 헤드가 전부 상향식이면 **0행이 맞다**.
    #   여기서 볼 것은 그 표가 화면까지 실려 오느냐다(키가 없으면 늘 빈다).
    check("기기표가 화면으로 실려 온다", "equipment" in t1,
          f"{len(eq1)}행 — {[e.get('desc') for e in eq1] or '이 저장본은 찍은 기기가 없다'}")

    print("\n" + "=" * 56)
    if fails:
        print(f"실패 {len(fails)}건")
        for f in fails:
            print("  -", f)
        return 1
    print("신축배관(FX) — 살아 있는 서버에서 확인")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
