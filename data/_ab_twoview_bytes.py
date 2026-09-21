# -*- coding: utf-8 -*-
"""[두 화면 선정일치 · §4 기준 4] 교체 없는 세션의 **산출 바이트**를 찍는다.

같은 도면·같은 손으로 표를 확정하고 `.sdf`+`.slf` 를 낸 뒤 해시를 적는다.
엔진을 되돌린 판(HEAD)과 이 판에서 각각 돌려 두 해시를 맞대면, 「§2-4 의
사유 기록이 산출을 안 건드린다」가 말이 아니라 수로 남는다.

    python data/_ab_twoview_bytes.py            # 지금 판
    (엔진 3파일을 HEAD 로 되돌린 뒤 다시)       # 이전 판
"""
from __future__ import annotations

import hashlib
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"


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
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(PLAN, "rb") as fh:
            r = c.post("/api/module-f/slot/open",
                       data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
                       content_type="multipart/form-data")
        sid = (r.get_json() or {})["sid"]
        wait(c, sid)
        c.post("/api/module-f/slot/read", json={"sid": sid, "method": "manual"})
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": (rec.get("adopt") or {})
                               .get("conf_min")}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 10})
        c.post("/api/module-f/design/build", json={"sid": sid, "k": 10})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패")
            return 1
        j = c.post("/api/module-f/design/emit", json={"sid": sid}).get_json()
        print(f"\n■ 산출 바이트 · emit ok={j.get('ok')}")
        h = (sess["design"]["got"].get("handoff") or {})
        print(f"  handoff 키: {sorted(h)}")
        for key in ("design_sdf_path", "design_slf_path"):
            p = sess.get(key)
            if p and os.path.isfile(p):
                b = open(p, "rb").read()
                print(f"  {os.path.basename(p):<44} {len(b):>9,}B "
                      f"sha256 {hashlib.sha256(b).hexdigest()[:16]}")
        tbl = sess["design"]["tables"]
        print(f"  표 — 노드 {len(tbl.nodes)} · 배관 {len(tbl.pipes)}"
              f" · 노즐 {len(tbl.nozzles)} · 부속 {len(tbl.fittings)}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
