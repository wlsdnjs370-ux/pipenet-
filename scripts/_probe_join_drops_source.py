# -*- coding: utf-8 -*-
"""[F-10g 조사] 이음(살리기)이 급수원을 지우는가.

완주 실측에서 잡혔다 — 원클릭으로 급수·밸브를 놓고(급수 1 · 밸브 1) 흐린
배관을 이은 뒤에 **급수 0 · 밸브 1** 이 됐다. 그러면 「살리기 → 다시 계산」이
「급수 시작 위치를 먼저 찍어야」로 막힌다. 상무가 요구한 바로 그 왕복이다.

추측하지 않고 한 클릭씩 재서 «어느 클릭에서» 사라지는지 가린다.

    python scripts/_probe_join_drops_source.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=3000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def st(c, sid):
    s = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
    return (len(s.get("sources") or []), len(s.get("valves") or []),
            s.get("counts", {}).get("pts"), s.get("counts", {}).get("edges"), s)


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
        with c.session_transaction() as s_:
            s_["authed"] = True
        with open(dxf, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, dxf.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True, "heads": {"conf_min": 0.9}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)
        _, _, _, _, s = st(c, sid)

        # 헤드에 가장 가까운 배관 끝점에 원클릭
        pts = []
        for g in s["body_groups"]:
            a = g["segs"]
            for i in range(0, len(a) - 3, 4):
                pts.append((float(a[i]), float(a[i + 1])))
        hx, hy = float(s["heads"][0][0]), float(s["heads"][0][1])
        px, py = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": px, "y": py})
        j = wait(c, sid)
        print(f"원클릭 {j['state']}")
        n = st(c, sid)
        print(f"  급수 {n[0]} · 밸브 {n[1]} · 절점 {n[2]} · 간선 {n[3]}")

        # 끊긴 곳 찾기 → 후보 하나를 이음 모드로 두 번 클릭
        c.post("/api/module-f/edit/autojoin/scan", json={"sid": sid})
        s2 = st(c, sid)[4]
        lines = ((s2.get("autojoin") or {}).get("lines")) or []
        print(f"  이음 후보 {len(lines)}곳")
        if not lines:
            print("  후보가 없어 판단 보류")
            return 3
        c.post("/api/module-f/edit/mode", json={"sid": sid, "mode": "이음"})
        ln = lines[0]
        for k, (x, y) in enumerate(((ln[0], ln[1]), (ln[2], ln[3])), 1):
            r = c.post("/api/module-f/edit/click",
                       json={"sid": sid, "x": x, "y": y, "max_d": 300.0})
            rep = (r.get_json() or {}).get("report")
            n = st(c, sid)
            print(f"  이음 클릭 {k} — {rep} → 급수 {n[0]} · 밸브 {n[1]} · "
                  f"절점 {n[2]} · 간선 {n[3]}")

        r = c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
        print(f"\n  다시 계산 {r.status_code} — "
              f"{(r.get_json() or {}).get('message', 'ok')}")
    print("\n  급수가 «이음 클릭» 에서 0 이 되면 그 클릭이 범인이다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
