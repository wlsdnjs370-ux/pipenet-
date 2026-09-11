# -*- coding: utf-8 -*-
"""[영역 넘어감] 영역1 안에서 고른 헤드가 영역2 로 «강제로» 넘어가는가.

사용자 지적(2026-09-11): 「대명동 기준, 영역1내에 있는 2개 헤드가 영역 2 및
배관망을 강제로 넘어간다.」

의심: 못 붙는 헤드를 다음 순위로 채울 때(`worst_cand`) 후보 범위가 **영역
합집합**이라, 영역1 의 헤드가 못 붙으면 그 자리를 **영역2** 의 헤드가 채운다.
그러면 설계면적이 두 영역에 걸치고 corridor 가 건물을 가로지른다.

짐작으로 고치지 않는다 — 영역별로 세어 본다.

    python data/_probe_zone_cross.py [--k 10]
"""
from __future__ import annotations

import argparse
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


def in_rect(p, r):
    return r[0] <= p[0] <= r[2] and r[1] <= p[1] <= r[3]


def zone_of(p, rects):
    for i, r in enumerate(rects):
        if in_rect(p, r):
            return i + 1
    return 0                      # 0 = 어느 영역에도 없음


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=10)
    ap.add_argument("--tight", action="store_true",
                    help="영역1 을 K개에 딱 맞게 좁혀 그린다(사용자 상황)")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with isolated_workdir(prefix="zonecross_"), srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        # ── 완성품 시험과 **같은 찍기 경로** — 이 경로라야 «못 붙는 헤드» 가
        #    실제로 생긴다(실측 133개 중 2개).
        with open(DXF, "rb") as f:
            raw = f.read()
        sid = c.post("/api/module-f/open", data={
            "dxf_file": (_io.BytesIO(raw), os.path.basename(str(DXF)))},
            content_type="multipart/form-data").get_json()["sid"]
        wait(c, sid)
        c.post("/api/module-f/pick/mode", json={"sid": sid, "action": "pipe"})
        c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
        c.post("/api/module-f/pick/mode",
               json={"sid": sid, "action": "complete"})
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
        xs = [p[0] for p in disks]
        ys = [p[1] for p in disks]
        x0, x1 = min(xs), max(xs)
        y0, y1 = min(ys), max(ys)
        # 영역 두 개 — 왼쪽 절반(영역1) · 오른쪽 절반(영역2). 사용자가 화면에서
        # 사각형 두 개를 그리는 것과 같은 입력이다(합집합).
        mid = (x0 + x1) / 2
        if args.tight:
            # ★사용자 상황 — 영역1 을 «K개에 딱 맞게» 좁혀 그린다. 그 안의
            #   헤드 중 둘이 전개에 안 붙으면 채울 것이 영역1 에 안 남고,
            #   후보 범위가 합집합이라 **영역2** 에서 끌어온다.
            import math as _m
            seed = 123 if len(disks) > 123 else 0
            near = sorted(range(len(disks)),
                          key=lambda i: _m.dist(disks[i], disks[seed]))
            core = near[:args.k]
            cx0 = min(disks[i][0] for i in core) - 300
            cx1 = max(disks[i][0] for i in core) + 300
            cy0 = min(disks[i][1] for i in core) - 300
            cy1 = max(disks[i][1] for i in core) + 300
            z1 = [cx0, cy0, cx1, cy1]
            # 영역2 — 영역1 에서 가장 먼 헤드 둘레. 사람이 「이쪽도 후보」 라고
            # 찍은 두 번째 사각형이다.
            far = max(range(len(disks)),
                      key=lambda i: _m.dist(disks[i], disks[seed]))
            near2 = sorted(range(len(disks)),
                           key=lambda i: _m.dist(disks[i], disks[far]))[:args.k]
            z2 = [min(disks[i][0] for i in near2) - 300,
                  min(disks[i][1] for i in near2) - 300,
                  max(disks[i][0] for i in near2) + 300,
                  max(disks[i][1] for i in near2) + 300]
        else:
            z1 = [x0 - 1, y0 - 1, mid, y1 + 1]
            z2 = [mid, y0 - 1, x1 + 1, y1 + 1]
        rects = [z1, z2]
        n1 = sum(1 for p in disks if in_rect(p, z1))
        n2 = sum(1 for p in disks if in_rect(p, z2))
        print(f"\n■ 영역 넘어감 계측 · 도면 헤드 {len(disks)} · K={args.k}")
        print(f"  영역1 (x ≤ {mid:.0f}) 헤드 {n1} · 영역2 헤드 {n2}")

        jw = c.post("/api/module-f/edit/worst",
                    json={"sid": sid, "k": args.k,
                          "zones": [z1, z2]}).get_json()
        if not jw.get("ok"):
            print(f"★최불리 실패 — {str(jw)[:200]}")
            return 1
        picked = [int(i) for i in (sess["worst"].get("heads") or ())]
        pz = [zone_of(disks[i], rects) for i in picked]
        print(f"  [손질] 고른 {len(picked)}개 — "
              f"영역1 {pz.count(1)} · 영역2 {pz.count(2)}")

        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패")
            return 1
        fin = [int(i) for i in
               ((sess["design"]["got"].get("worst") or {}).get("heads") or ())]
        fz = [zone_of(disks[i], rects) for i in fin]
        print(f"  [표]  최종 {len(fin)}개 — "
              f"영역1 {fz.count(1)} · 영역2 {fz.count(2)}")

        h = (sess["design"]["got"].get("handoff") or {})
        print(f"\n  채움 {h.get('filled', 0)}개")
        for tag, key in (("빠진 것", "swapped_out"), ("채운 것", "swapped_in")):
            for r in (h.get(key) or ()):
                i = int(r["disk"])
                print(f"    {tag} · 헤드 {i} · 영역{zone_of(disks[i], rects)}"
                      f" · {r.get('xy')}"
                      + (f" — {r.get('why_text') or ''}" if r.get("why") else ""))

        for m in (h.get("messages") or ()):
            print(f"    [문구] {m}")
        cross = sum(1 for z in fz if z == 2) - sum(1 for z in pz if z == 2)
        print()
        if cross > 0:
            print(f"  ★★재현됐다 — 영역2 로 **{cross}개** 넘어갔다."
                  f" 설계면적이 두 영역에 걸친다.")
        else:
            print("  영역을 안 넘었다.")
        # corridor 가 두 영역에 걸치는지 — 사람이 보는 «배관망이 넘어간다» 다.
        w = sess["design"]["got"].get("worst") or {}
        pts = b.pts
        zs = set()
        for (a, cc) in (w.get("loads") or {}):
            for n in (a, cc):
                if 0 <= n < len(pts):
                    zs.add(zone_of((float(pts[n][0]), float(pts[n][1])), rects))
        print(f"  corridor 가 걸친 영역: {sorted(zs)}"
              + ("  ★두 영역에 걸쳤다" if {1, 2} <= zs else ""))
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
