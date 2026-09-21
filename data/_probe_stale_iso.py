# -*- coding: utf-8 -*-
"""[아이소가 옛 것] 최불리를 다시 고른 뒤 표를 안 확정하면 무엇이 보이나.

의심: 사용자가 손질에서 최불리를 **다시** 고르면 `sess["worst"]` 는 새 선정이
되지만 `sess["design"]`(표·아이소)은 **직전 build 것 그대로**다. 그러면
  · 「평면에서 보기」  = 새 선정 (corridor 가 바뀐다)
  · 아이소            = 옛 표
가 되어 「최불리에서 등록한 배관망이 아이소에 반영이 안 된다」로 보인다.

짐작으로 고치지 않는다 — 재본다.
"""
from __future__ import annotations

import io as _io
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

DXF = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


def wait(c, sid, limit=20000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with isolated_workdir(prefix="stale_"), srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(DXF, "rb") as f:
            raw = f.read()
        sid = c.post("/api/module-f/open", data={
            "dxf_file": (_io.BytesIO(raw), os.path.basename(str(DXF)))},
            content_type="multipart/form-data").get_json()["sid"]
        wait(c, sid)
        c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
        c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
        c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "complete"})
        c.post("/api/module-f/pick/mode",
               json={"sid": sid, "action": "slot", "slot": "상향"})
        c.post("/api/module-f/pick/suggest", json={"sid": sid})
        wait(c, sid)
        cands = ((c.get(f"/api/module-f/convert/result?sid={sid}")
                  .get_json()["result"] or {}).get("candidates") or [])
        for c_ in cands:
            d = c.post("/api/module-f/pick/click",
                       json={"sid": sid, "x": c_["x"], "y": c_["y"],
                             "max_d": 300}).get_json()
            if (d.get("report") or {}).get("동작") == "취소":
                c.post("/api/module-f/pick/click",
                       json={"sid": sid, "x": c_["x"], "y": c_["y"],
                             "max_d": 300})
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        b = sess["edit"].board
        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        seg = est["body_groups"][0]["segs"]
        c.post("/api/module-f/edit/mode",
               json={"sid": sid, "mode": "급수시작위치"})
        c.post("/api/module-f/edit/click",
               json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                     "y": (seg[1] + seg[3]) / 2, "max_d": 2000})

        disks = [(float(d[0]), float(d[1])) for d in b.disks]
        xs = [p[0] for p in disks]; ys = [p[1] for p in disks]
        mid = (min(xs) + max(xs)) / 2
        zoneL = [min(xs) - 1, min(ys) - 1, mid, max(ys) + 1]
        zoneR = [mid, min(ys) - 1, max(xs) + 1, max(ys) + 1]

        print("\n■ 최불리를 다시 고른 뒤 «표 확정» 을 안 누르면")
        c.post("/api/module-f/edit/worst",
               json={"sid": sid, "k": 8, "zones": [zoneL]})
        A = sorted(int(i) for i in sess["worst"]["heads"])
        print(f"  1) 최불리 A (왼쪽 영역 · K=8) : 헤드 {A}")
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 8})
        wait(c, sid)
        tblA = sess["design"]["tables"]
        nozA = sorted(int(z.get("in")) for z in tblA.nozzles)
        print(f"     표 확정 → 노즐 {len(nozA)} · 절점 {len(tblA.nodes)}")

        # ── 사람이 손질로 돌아가 «다른 영역» 으로 다시 고른다 (표는 안 누른다)
        c.post("/api/module-f/edit/worst",
               json={"sid": sid, "k": 8, "zones": [zoneR]})
        B = sorted(int(i) for i in sess["worst"]["heads"])
        print(f"  2) 최불리 B (오른쪽 영역 · K=8): 헤드 {B}")
        print(f"     A 와 겹치는 헤드: {len(set(A) & set(B))}개")

        pv = c.get(f"/api/module-f/design/preview?sid={sid}").get_json()
        tblN = sess["design"]["tables"]
        nozN = sorted(int(z.get("in")) for z in tblN.nozzles)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        shown = len(((st.get("worst") or {}).get("heads")) or ())
        print(f"\n  3) 표를 다시 안 누르고 수리계산 화면을 열면")
        print(f"     평면 보기(corridor) 헤드 : {shown}  ← 새 선정 B")
        print(f"     아이소(표) 노즐          : {len(nozN)}  ← {'옛 표 A 그대로' if nozN == nozA else '갱신됨'}")
        print(f"     preview 가 «옛 것» 이라고 말하나 : "
              f"{'예' if (pv.get('stale') or pv.get('message')) else '★아니오 — 아무 말이 없다'}")
        same = nozN == nozA
        told = bool(pv.get("stale"))
        if same and told:
            print("\n  ★표는 옛 것 그대로지만 화면이 «옛 것» 이라고 말한다 — 조치됨")
            print(f"     stale: {pv['stale']}")
            return 0
        if same:
            print("\n  ★★재현 — 평면은 새 선정, 아이소는 옛 표. 화면은 아무 말도 안 한다.")
            return 2
        print("\n  갱신되고 있다(이 경로는 아니다)")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
