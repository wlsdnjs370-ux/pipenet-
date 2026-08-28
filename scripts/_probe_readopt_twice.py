# -*- coding: utf-8 -*-
"""[F-11a 조사] 두 번째 채택이 첫 채택을 지우는가.

F-11a 로 LH306 이 자동 채택되어 손질까지 가게 되자, 검증기가 그 뒤에 「이 기준
으로 다시 채택」을 눌렀을 때 **재료 0묶음 · 헤드 0픽** 이 나왔다.

의심: E 의 찍기는 «같은 서명을 다시 누르면 취소» 다(문양 서명 토글). 재채택이
같은 묶음을 다시 클릭하면 켜 놓은 것을 끄게 된다. 그러면 §0.1 의 «수렴성»
(같은 수정을 두 번 시키지 않는다)이 깨진다 — 사람이 기준을 바꿔 다시 채택했을
뿐인데 앞서 찍힌 것이 사라진다.

한 번 채택 → 상태, 두 번 채택 → 상태 를 나란히 재서 가린다.

    python scripts/_probe_readopt_twice.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "samples" / "dxf" / "LH306동_평면도.dxf"


def wait(c, sid, limit=6000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def world(c, sid):
    s = (c.get(f"/api/module-f/world?sid={sid}").get_json() or {}).get("state") or {}
    return (len(s.get("materials") or []), s.get("n_heads"), s.get("n_clicks"))


def adopt(c, sid, conf):
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True, "heads": {"conf_min": conf}})
    j = wait(c, sid)
    r = (c.get(f"/api/module-f/convert/result?sid={sid}").get_json()
         or {}).get("result") or {}
    return j, r


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEF
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})

        print(f"\n■ {dxf.name}")
        print(f"    처음        재료 {world(c, sid)[0]}묶음 · "
              f"헤드 {world(c, sid)[1]} · 클릭 {world(c, sid)[2]}")

        j, res = adopt(c, sid, 0.75)
        m1, h1, k1 = world(c, sid)
        print(f"    1차 채택    {j['state']} · 찍힘 {res.get('head_applied')} "
              f"· 이미 {res.get('head_already')} → 재료 {m1}묶음 · 헤드 {h1} "
              f"· 클릭 {k1}")

        j, res = adopt(c, sid, 0.75)
        m2, h2, k2 = world(c, sid)
        print(f"    2차 채택    {j['state']} · 찍힘 {res.get('head_applied')} "
              f"· 이미 {res.get('head_already')} → 재료 {m2}묶음 · 헤드 {h2} "
              f"· 클릭 {k2}")
        if res.get("error"):
            print(f"      ★2차 오류 — {res['error']}")

        ok = (m2 >= m1 and (h2 or 0) >= (h1 or 0))
        print(f"\n  {'[OK] 두 번 채택해도 안 지워진다' if ok else '★두 번째 채택이 앞의 것을 지운다 — 수렴성 위반'}")
        if not ok:
            print("     (E 의 찍기는 같은 서명을 다시 누르면 «취소» 다 —")
            print("      재채택이 그 토글을 다시 밟는지 확인해야 한다.)")
        return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
