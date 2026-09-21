# -*- coding: utf-8 -*-
"""[표 들여다보기] 헤드 표고가 실제로 어떻게 실리나 — 짐작 말고 값을 본다."""
from __future__ import annotations

import io as _io
import os
import sys
import time
from collections import Counter
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
    K = 10
    with isolated_workdir(prefix="elev_"), srv.app.test_client() as c:
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
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": K})
        c.post("/api/module-f/design/build", json={"sid": sid, "k": K})
        wait(c, sid)
        tbl = sess["design"]["tables"]

        print("\n■ 표 들여다보기")
        print(f"  절점 {len(tbl.nodes)} · 배관 {len(tbl.pipes)}"
              f" · 노즐 {len(tbl.nozzles)}")
        print("\n  [절점 표고 분포]")
        for v, n in Counter(round(float(x.get("elevation") or 0), 3)
                            for x in tbl.nodes).most_common():
            print(f"    {v:>8} m : {n}개")
        print("\n  [노즐 행 3개]")
        for z in tbl.nozzles[:3]:
            print("   ", {k: z.get(k) for k in list(z)[:8]})
        nz = {str(z.get("in")) for z in tbl.nozzles}
        print("\n  [노즐이 앉은 절점 3개]")
        for n in tbl.nodes:
            if str(n.get("label")) in nz:
                print("   ", n)
                nz.discard(str(n.get("label")))
                if len(nz) <= len(tbl.nozzles) - 3:
                    break
        print("\n  [meta 에서 표고 관련]")
        for k, v in dict(tbl.meta).items():
            if any(t in str(k) for t in ("표고", "고도", "헤드", "lift", "Z")):
                print(f"    {k} = {v}")
        print("\n  [손질판 헤드 종류]", Counter(b.disk_kinds).most_common())
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
