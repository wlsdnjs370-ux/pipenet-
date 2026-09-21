# -*- coding: utf-8 -*-
"""[헤드를 고를 수 있나] 종류를 바꾸려면 먼저 골라야 한다 — 그 길이 있나.

화면에는 「고른 헤드 종류」 단추 세 개가 있는데, 정작 «헤드 고르기» 모드는
없다. 고를 수 있는 길은 이음·삭제 모드의 클릭뿐이다:

    이음 모드 : _prefer_head 가 먼저 — 헤드가 배관보다 가까우면 고름,
                그런데 두 번째 클릭은 «이음» 이 되어 버린다
    삭제 모드 : **배관 삭제가 먼저** — 헤드 위를 눌러도 근처에 배관이 있으면
                헤드가 아니라 배관이 지워진다

  실제로 어떻게 되는지 헤드 열 개에 대해 두 모드로 눌러 본다.
"""
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
    with isolated_workdir(prefix="hsel_"), srv.app.test_client() as c:
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

        heads = [(i, float(d[0]), float(d[1])) for i, d in enumerate(b.disks)][:10]
        print("\n■ 헤드를 고를 수 있나 — 헤드 10개 × 두 모드 (중심을 정확히 클릭)")
        for mode in ("이음", "삭제", "헤드"):
            c.post("/api/module-f/edit/mode", json={"sid": sid, "mode": mode})
            acts, sel_ok, e0 = Counter(), 0, len(b.edges)
            for (i, x, y) in heads:
                r = c.post("/api/module-f/edit/click",
                           json={"sid": sid, "x": x, "y": y,
                                 "max_d": 200}).get_json()
                rep = r.get("report") or {}
                acts[str(rep.get("동작") or "없음")] += 1
                if (r.get("state") or {}).get("selected_head"):
                    sel_ok += 1
                # 이음 모드는 pending 이 남으면 다음 클릭이 이음이 된다 — 비운다
                if mode == "이음" and rep.get("동작") == "대기":
                    c.post("/api/module-f/edit/mode",
                           json={"sid": sid, "mode": "이음"})
            e1 = len(sess["edit"].board.edges)
            print(f"\n  [{mode} 모드]  헤드가 고름 상태로 남은 클릭 {sel_ok}/10")
            print(f"    동작 분포: {dict(acts)}")
            print(f"    간선 {e0} → {e1}"
                  + ("   ★★배관이 지워졌다" if e1 != e0 else "   (배관 불변)"))

        print("\n  ※ 「고른 헤드 종류」 단추는 selected_head 가 있어야 먹는다"
              " — 고르지 못하면 종류를 영영 못 바꾼다.")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
