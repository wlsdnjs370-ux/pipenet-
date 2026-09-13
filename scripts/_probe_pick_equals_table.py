# -*- coding: utf-8 -*-
"""[요구 한 줄] 최불리 선정의 헤드 개수·좌표 == 표 확정의 헤드 개수·좌표.

사용자 지시(2026-09-13): 「간단히, 최불리 선정으로 나온 헤드 개수랑 표 확정으로
나온 헤드 개수와 좌표가 정확히 똑같아지게만 조치해 달라.」

그래서 **그것만** 잰다. 다른 수는 안 본다.

    ㉠ 최불리 선정 직후 : sess["worst"]["heads"] → board.disks 좌표
    ㉡ 표 확정 직후     : tbl.nozzles 의 절점 좌표 → board mm 로 되돌린 것

  둘을 좌표로 1:1 짝지어 «개수가 같은가 · 전부 짝이 있는가» 만 본다.
  문턱은 헤드 반지름에 딸린 값이다(상향식은 접속 절점이 원 밑을 지나는 관
  위에 있어 중심에서 최대 반지름만큼 떨어진다 — B1F 실측 150mm).

    python scripts/_probe_pick_equals_table.py [--k 30] [--zones N]
"""
from __future__ import annotations

import argparse
import io as _io
import math
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


def near_for(r):
    return float(r) + 60.0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--zones", type=int, default=0, help="영역 개수 (0=없음)")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with isolated_workdir(prefix="equal_"), srv.app.test_client() as c:
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

        body = {"sid": sid, "k": args.k}
        if args.zones:
            ds = [(float(d[0]), float(d[1])) for d in b.disks]
            xs = [p[0] for p in ds]; ys = [p[1] for p in ds]
            lo, hi = min(xs), max(xs)
            step = (hi - lo) / args.zones
            body["zones"] = [[lo + i * step - 1, min(ys) - 1,
                              lo + (i + 1) * step + 1, max(ys) + 1]
                             for i in range(args.zones)]
        t0 = time.perf_counter()
        jw = c.post("/api/module-f/edit/worst", json=body).get_json()
        t_sel = time.perf_counter() - t0
        if not jw.get("ok"):
            print(f"★최불리 실패 — {str(jw)[:220]}")
            return 1
        picked = [int(i) for i in (sess["worst"].get("heads") or ())]
        A = [(float(b.disks[i][0]), float(b.disks[i][1]),
              float(b.disks[i][2])) for i in picked]

        t0 = time.perf_counter()
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패")
            return 1
        t_tbl = time.perf_counter() - t0
        got, tbl = sess["design"]["got"], sess["design"]["tables"]
        origin = got.get("origin_mm") or (0.0, 0.0)
        ox, oy = float(origin[0]) - 1000.0, float(origin[1]) - 1000.0
        at = {str(n.get("label")): (float(n.get("x") or 0) + ox,
                                    float(n.get("y") or 0) + oy)
              for n in tbl.nodes}
        B = [at[str(z.get("in"))] for z in tbl.nozzles
             if str(z.get("in")) in at]

        print(f"\n■ 최불리 선정 == 표 확정 · 대명동 · K={args.k}"
              f" · 영역 {args.zones or '없음'}")
        print(f"  ㉠ 최불리 선정 헤드 : {len(A)}개   ({t_sel:.1f}s)")
        print(f"  ㉡ 표 확정   노즐   : {len(B)}개   ({t_tbl:.1f}s)")

        used, gaps, miss = set(), [], []
        for (x, y, r) in A:
            best, bd = None, 1e18
            for t, p in enumerate(B):
                if t in used:
                    continue
                dd = math.dist(p, (x, y))
                if dd < bd:
                    best, bd = t, dd
            if best is None or bd > near_for(r):
                miss.append((x, y))
                continue
            used.add(best)
            gaps.append(bd)
        extra = [B[t] for t in range(len(B)) if t not in used]
        gaps.sort()
        print(f"  ★1:1 로 맞은 헤드   : {len(gaps)}/{len(A)}"
              + (f"  (최대 오차 {gaps[-1]:.1f} mm)" if gaps else ""))
        print(f"  ★선정에만 있는 헤드 : {len(miss)}")
        for (x, y) in miss[:6]:
            print(f"      ({x:.0f}, {y:.0f})")
        print(f"  ★표에만 있는 노즐   : {len(extra)}")
        for (x, y) in extra[:6]:
            print(f"      ({x:.0f}, {y:.0f})")

        ok = (len(A) == len(B) and not miss and not extra)
        print("\n  " + ("★개수·좌표가 정확히 같다" if ok else
                        "★★아직 다르다"))
        return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
