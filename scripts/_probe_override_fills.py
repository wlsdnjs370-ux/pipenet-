# -*- coding: utf-8 -*-
"""[§18] 직접 입력이 «정말로» 미해결을 채우는가 — 그리고 나머지는 안 건드리는가.

두 가지를 한 번에 잰다:

  ① 채우기가 먹는가   — 미해결 n건 → 직접 입력 뒤 0건, 「직접 입력」 n건
  ② 나머지가 그대로인가 — 부속 개수·등가길이 합이 «채운 자리 말고는» 불변

②가 더 중요하다. 덮어쓰기가 규칙이 옳게 판정한 자리까지 건드리면, 사람이
한 자리를 채웠다고 산출 전체가 조용히 달라진다.

    python scripts/_probe_override_fills.py [도면.dxf]
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=6000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def snapshot(c, sid):
    pv = c.get(f"/api/module-f/design/preview?sid={sid}").get_json()
    tbl = pv["tables"]
    meta = {k: v for k, v in tbl["meta"]}
    un = tbl.get("unresolved") or {}
    # 부속 표를 지문으로 — (배관, 종류, 개수) 다중집합
    fit = sorted((str(r.get("pipe")), str(r.get("type")), int(r.get("count") or 0))
                 for r in tbl.get("fittings") or [])
    return {
        "n_kind": int(meta.get("부속 판정 불가") or 0),
        "n_len": int(meta.get("등가길이 미해결") or 0),
        "ov_kind": int(meta.get("직접 입력 — 부속 판정") or 0),
        "ov_len": int(meta.get("직접 입력 — 등가길이") or 0),
        "items": un.get("kind_items") or [],
        "applied": un.get("applied") or [],
        "fittings": fit,
        "pipes": [(str(p["label"]), p.get("dia")) for p in tbl["pipes"]],
    }


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
        s = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        pts = []
        for g in s["body_groups"]:
            a = g["segs"]
            for i in range(0, len(a) - 3, 4):
                pts.append((float(a[i]), float(a[i + 1])))
        hx, hy = float(s["heads"][0][0]), float(s["heads"][0][1])
        px, py = min(pts, key=lambda q: (q[0] - hx) ** 2 + (q[1] - hy) ** 2)
        c.post("/api/module-f/edit/anchor-click",
               json={"sid": sid, "x": px, "y": py})
        wait(c, sid)
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        if wait(c, sid)["state"] != "done":
            print("표 확정 실패")
            return 2

        before = snapshot(c, sid)
        print(f"\n■ 채우기 전")
        print(f"    부속 판정 불가 {before['n_kind']} · 등가길이 미해결 "
              f"{before['n_len']} · 직접 입력 {before['ov_kind']}/{before['ov_len']}")
        if not before["items"]:
            print("    미해결이 없어 채울 것이 없다 — 다른 도면으로 재 볼 것")
            return 3
        for x in before["items"][:3]:
            print(f"      {x['pipe']} · {x['where']} · 노드 {x['node']}")

        # ── 자리 하나만 채운다. 나머지가 그대로인지 보려면 하나여야 한다.
        target = before["items"][0]
        r = c.post("/api/module-f/design/fitting-override", json={
            "sid": sid,
            "kind": [{"node": target["node"], "pipe": target["pipe"],
                      "kind": "none", "note": "현장 확인 — 직선이다"}]})
        print(f"\n■ 직접 입력 저장 {r.status_code} — "
              f"{(r.get_json() or {}).get('message', '')}")
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
        if wait(c, sid)["state"] != "done":
            print("재확정 실패")
            return 2
        after = snapshot(c, sid)
        print(f"\n■ 채우기 뒤")
        print(f"    부속 판정 불가 {after['n_kind']} · 등가길이 미해결 "
              f"{after['n_len']} · 직접 입력 {after['ov_kind']}/{after['ov_len']}")
        for a in after["applied"][:3]:
            print(f"      적용 — {a.get('what')} · {a.get('pipe')} · "
                  f"{a.get('kind')} · 사유 {a.get('note')!r}")

        print(f"\n■ ① 채우기가 먹었나")
        ok1 = (after["n_kind"] == before["n_kind"] - int(target.get("n") or 1)
               and after["ov_kind"] >= 1)
        print(f"    미해결 {before['n_kind']} → {after['n_kind']} · "
              f"직접 입력 {after['ov_kind']}건   "
              f"{'[OK]' if ok1 else '[FAIL]'}")

        print(f"\n■ ② 나머지가 그대로인가 (더 중요하다)")
        if before["pipes"] != after["pipes"]:
            print("    ★배관 호칭경이 바뀌었다 — 채우기가 관경을 건드렸다")
            ok2 = False
        else:
            # 채운 배관 말고는 부속 표가 같아야 한다.
            tp = str(target["pipe"])
            b = [x for x in before["fittings"] if x[0] != tp]
            a2 = [x for x in after["fittings"] if x[0] != tp]
            ok2 = (b == a2)
            print(f"    호칭경 불변 · 채운 배관({tp}) 밖 부속 표 "
                  f"{'동일' if ok2 else '★달라졌다'}   "
                  f"{'[OK]' if ok2 else '[FAIL]'}")
            if not ok2:
                only_b = [x for x in b if x not in a2][:3]
                only_a = [x for x in a2 if x not in b][:3]
                print(f"      전에만 {only_b}")
                print(f"      후에만 {only_a}")
        return 0 if (ok1 and ok2) else 1


if __name__ == "__main__":
    sys.exit(main())
