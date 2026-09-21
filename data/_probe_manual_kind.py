# -*- coding: utf-8 -*-
"""[수동 지정이 아이소에 오나] 손질에서 고친 헤드 종류가 표·아이소에 반영되나.

상하향은 그림 취향이 아니라 **표고의 부호**다(상향 +0.3 m · 하향 −0.3 m).
등각에서 스텁이 어느 쪽으로 서는지가 그것으로 갈린다. 그래서 「수동 지정한
헤드가 아이소에 반영 안 된다」는 지적은 이 값으로 재면 참·거짓이 갈린다.

  1) 설계면적 안의 헤드 하나를 고른다
  2) 손질에서 그 헤드의 종류를 **반대로** 바꾼다 (/edit/kind)
  3) 표를 확정하고 그 헤드의 표고 부호가 따라 바뀌는지 본다
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


def build(c, sid, k):
    c.post("/api/module-f/design/build", json={"sid": sid, "k": k})
    return wait(c, sid).get("state")


def head_elevs(sess):
    """표에서 노즐이 달린 절점의 표고 — {라벨: 표고}."""
    tbl = sess["design"]["tables"]
    noz = {str(z.get("in")) for z in tbl.nozzles}
    return {str(n.get("label")): float(n.get("elevation") or 0.0)
            for n in tbl.nodes if str(n.get("label")) in noz}


def main() -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    K = 10

    with isolated_workdir(prefix="mankind_"), srv.app.test_client() as c:
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
        if build(c, sid, K) != "done":
            print("★표 확정 실패"); return 1

        picked = [int(i) for i in sess["worst"]["heads"]]
        kinds0 = list(b.disk_kinds)
        e0 = head_elevs(sess)
        print(f"\n■ 수동 지정한 헤드 종류가 표에 오나 · K={K}")
        print(f"  설계면적 헤드 {len(picked)} · 표 노즐 표고 "
              f"{sorted(set(round(v, 3) for v in e0.values()))}")
        print(f"  손질이 든 종류(설계면적): "
              f"{sorted(set(kinds0[i] for i in picked))}")

        # ── 설계면적 안의 헤드 하나를 고르고 종류를 반대로
        target = picked[0]
        cur = kinds0[target]
        new = "상향식" if cur != "상향식" else "하향식"
        x, y = float(b.disks[target][0]), float(b.disks[target][1])
        # ★«헤드» 모드로 고른다 — 삭제 모드는 배관을 먼저 먹는다(실측 10/10).
        c.post("/api/module-f/edit/mode", json={"sid": sid, "mode": "헤드"})
        r = c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": x, "y": y, "max_d": 200}).get_json()
        sel = ((r.get("state") or {}).get("selected_head"))
        rk = c.post("/api/module-f/edit/kind",
                    json={"sid": sid, "kind": new}).get_json()
        print(f"\n  헤드 {target} ({x:.0f},{y:.0f}) · {cur} → {new}"
              f" · 고름 {'성공' if sel else '실패'}"
              f" · 적용 {rk.get('applied')}")
        after = list(sess["edit"].board.disk_kinds)
        print(f"  손질판의 그 헤드 종류: {after[target]}"
              f"  {'★안 바뀜' if after[target] == cur else '바뀜'}")

        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": K})
        if build(c, sid, K) != "done":
            print("★재확정 실패"); return 1
        e1 = head_elevs(sess)
        s0 = sorted(set(round(v, 3) for v in e0.values()))
        s1 = sorted(set(round(v, 3) for v in e1.values()))
        print(f"\n  표 노즐 표고  전 {s0}  →  후 {s1}")
        same = (s0 == s1)
        print("\n  " + ("★★표고가 그대로다 — 손질에서 고친 종류가 표에 안 온다"
                        if same else
                        "★표고가 따라 바뀌었다 — 수동 지정이 표에 온다"))
        return 2 if same else 0


if __name__ == "__main__":
    raise SystemExit(main())
