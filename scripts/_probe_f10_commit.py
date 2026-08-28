# -*- coding: utf-8 -*-
"""[F-10a] 기본 흐름이 조립에서 왜 죽는가 — 추정하지 말고 잡 오류를 읽는다.

브라우저 검증에서 「배관망 구성」 잡이 0.3초 만에 error 로 끝났고, 헤드가
0픽이었다. 대명동은 정찰 높음 띠가 0 이라(실측 높음 0 · 중간 40 · 낮음 2)
기본 기준 0.9 로는 채택되는 헤드가 없다.

그래서 서버에서 그대로 재현해 **오류 문장을 그대로** 본다. 0.9 와 0.75 를
나란히 돌려, 헤드 0 이 원인인지 다른 것인지 가른다.

    python scripts/_probe_f10_commit.py
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else (ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf")


def wait(c, sid, limit=1800):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.2)
    return {"state": "timeout"}


def run(c, conf):
    with open(DXF, "rb") as fh:
        r = c.post("/api/module-f/slot/open",
                   data={"dxf_file": (fh, DXF.name), "kind": "plan"},
                   content_type="multipart/form-data")
    d = r.get_json() or {}
    if "sid" not in d:
        print(f"  ★열기 실패 {r.status_code} — {str(d)[:200]}")
        raise SystemExit(1)
    sid = d["sid"]
    wait(c, sid)
    rec = c.get(f"/api/module-f/recon?sid={sid}").get_json().get("recon") or {}
    print(f"  정찰 {rec.get('state')} · 띠 {rec.get('bands')} · "
          f"묶음 {rec.get('bundles')}")

    c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
    c.post("/api/module-f/pick/adopt",
           json={"sid": sid, "materials": True, "heads": {"conf_min": conf}})
    j = wait(c, sid)
    res = j.get("result") or {}
    if not res:
        res = (c.get(f"/api/module-f/convert/result?sid={sid}")
               .get_json().get("result") or {})
    print(f"  채택 {j['state']} · 재료 {len(res.get('mat_applied') or [])}묶음 "
          f"· 헤드 {res.get('head_applied')} · 유령 {res.get('head_skipped')}")
    if j.get("error"):
        print(f"    ★채택 오류 — {j['error']}")

    r = c.post("/api/module-f/pick/commit", json={"sid": sid})
    if r.status_code != 200:
        print(f"  ★조립 거절 {r.status_code} — {r.get_json()}")
        return sid
    j = wait(c, sid)
    print(f"  조립 {j['state']}")
    if j.get("error"):
        print(f"    ★조립 오류 — {j['error']}")
        for ln in (j.get("lines") or [])[-6:]:
            print(f"      {ln}")
    return sid


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import importlib
    srv = importlib.import_module("대조 서버")
    app = srv.app
    app.config["TESTING"] = True

    if not DXF.is_file():
        print("도면 없음:", DXF)
        return 1

    with app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True          # 게이트 키 이름은 «authed» 다
        for conf in (0.9, 0.75):
            print(f"\n■ 채택 기준 {conf}")
            run(c, conf)
    print("\n  0.9 만 죽고 0.75 는 산다면 원인은 «헤드 0» 이다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
