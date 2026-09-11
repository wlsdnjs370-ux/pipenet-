# -*- coding: utf-8 -*-
"""[최불리 배관망 ↔ 아이소] 손질에서 등록한 배관망이 그대로 표로 갔는가.

■ 사용자 지적 (2026-09-11)

  「4. 수리계산의 아이소매트릭에서 2. 찍기에서 수동으로 지정한 배관과 헤드를
   인식 못한다. 그 이전에 최불리 배관망에서 등록했던 배관망을 그대로 가져오는
   것조차 안 되고 있다.」

■ 여태 맞대 본 적이 없는 것

  §2-1 은 **헤드 집합**을 맞췄다(평면 30 == 표 노즐 30). 그런데 사람이 보는
  것은 헤드 점이 아니라 **배관망**이다. 손질 화면의 corridor 는 board 간선
  위에 그려지고, 표/아이소는 «제한 전개» 가 새로 만든 망이다 — 둘이 같은
  헤드를 덮더라도 **관이 같은 길을 지나는지는 따로 봐야 한다.**

  이 계측기는 그 둘을 간선 단위로 맞댄다:

      corridor  = worst["loads"] 의 키  = board 간선 (i, j)
      표(아이소) = got["edge_ref"]       = kfp 배관 -> board 간선 역참조

  · corridor 에 있는데 표에 없는 간선  -> 「등록한 배관망이 안 왔다」
  · 표에 있는데 corridor 에 없는 간선  -> 「내가 안 고른 배관이 생겼다」
  · 역참조가 없는 배관(세로 구간)은 정상 — 도면에 그려진 선이 아니다.

    python scripts/_probe_corridor_vs_iso.py [--k 10] [--zones]
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


def key(e):
    a, b = int(e[0]), int(e[1])
    return (min(a, b), max(a, b))


def mlen(pts, edges):
    t = 0.0
    for a, b in edges:
        if 0 <= a < len(pts) and 0 <= b < len(pts):
            t += math.dist(pts[a][:2], pts[b][:2])
    return t / 1000.0


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=10)
    ap.add_argument("--zones", action="store_true",
                    help="영역 두 곳을 지정해 사용자 상황에 맞춘다")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with isolated_workdir(prefix="corridor_"), srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        with open(DXF, "rb") as f:
            raw = f.read()
        sid = c.post("/api/module-f/open", data={
            "dxf_file": (_io.BytesIO(raw), os.path.basename(str(DXF)))},
            content_type="multipart/form-data").get_json()["sid"]
        wait(c, sid)

        # ── 2. 찍기 — «수동으로 지정» 하는 길 그대로 (배관 색 찍기 + 헤드 찍기)
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
        n_click = 0
        for c_ in cands:
            d = c.post("/api/module-f/pick/click",
                       json={"sid": sid, "x": c_["x"], "y": c_["y"],
                             "max_d": 300}).get_json()
            n_click += 1
            if (d.get("report") or {}).get("동작") == "취소":
                c.post("/api/module-f/pick/click",
                       json={"sid": sid, "x": c_["x"], "y": c_["y"],
                             "max_d": 300})
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        wait(c, sid)

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        b = sess["edit"].board
        print(f"\n■ 최불리 배관망 ↔ 표(아이소) 맞대기")
        print(f"  [2 찍기]  수동 클릭 {n_click} · 손질망 절점 {len(b.pts)} ·"
              f" 간선 {len(b.edges)} · 헤드 {len(b.disks)}")

        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        seg = est["body_groups"][0]["segs"]
        c.post("/api/module-f/edit/mode",
               json={"sid": sid, "mode": "급수시작위치"})
        c.post("/api/module-f/edit/click",
               json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                     "y": (seg[1] + seg[3]) / 2, "max_d": 2000})

        body = {"sid": sid, "k": args.k}
        if args.zones:
            disks = [(float(d[0]), float(d[1])) for d in b.disks]
            xs = [p[0] for p in disks]; ys = [p[1] for p in disks]
            mid = (min(xs) + max(xs)) / 2
            body["zones"] = [[min(xs) - 1, min(ys) - 1, mid, max(ys) + 1],
                             [mid, min(ys) - 1, max(xs) + 1, max(ys) + 1]]
        jw = c.post("/api/module-f/edit/worst", json=body).get_json()
        if not jw.get("ok"):
            print(f"★최불리 실패 — {str(jw)[:200]}")
            return 1
        w = sess["worst"]
        corridor = {key(e) for e in (w.get("loads") or {})}
        print(f"  [3 손질]  최불리 corridor 간선 {len(corridor)} ·"
              f" 총연장 {mlen(b.pts, corridor):.1f} m ·"
              f" 헤드 {len(w.get('heads') or ())}")

        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패")
            return 1
        got = sess["design"]["got"]
        tbl = sess["design"]["tables"]
        ref = got.get("edge_ref") or {}
        in_table = {key(e) for e in ref.values()}
        pipes = got["kfp"].get("pipe_data") or {}
        print(f"  [4 표]    배관 {len(pipes)} · 그중 도면 간선에 역참조되는 것"
              f" {len(ref)} · 세로 구간 {len(pipes) - len(ref)}")
        print(f"            표가 덮는 board 간선 {len(in_table)} ·"
              f" 총연장 {mlen(b.pts, in_table):.1f} m")

        # ── ★기하로 맞댄다. 절점 번호로 맞대면 «노드정리»(일직선 중간 절점
        #    병합) 때문에 멀쩡한 망도 전부 어긋난 것처럼 보인다 — 실측에서
        #    0.09m 짜리 간선 여럿이 3.07m 한 줄로 합쳐져 있었다.
        origin = got.get("origin_mm") or (0.0, 0.0)

        def to_board(n):
            return (float(n.get("x", 0) or 0) + float(origin[0]) - 1000.0,
                    float(n.get("y", 0) or 0) + float(origin[1]) - 1000.0)

        at = {str(n.get("label")): to_board(n) for n in tbl.nodes}
        dsegs = []
        # ★배관 행의 양 끝은 `in`/`out` 이다(PipeTablesG 규약). from/to 로 읽어
        #   «표 선분 0개» 를 냈고, 하마터면 「표에 배관이 하나도 안 왔다」로
        #   보고할 뻔했다 — 자료 구조를 짐작하지 말 것.
        for pr in tbl.pipes:
            a = at.get(str(pr.get("in")))
            z = at.get(str(pr.get("out")))
            if a and z and math.dist(a, z) > 1.0:
                dsegs.append((a, z))

        def seg_dist(p, a, z):
            ax, ay = a; zx, zy = z; px, py = p
            dx, dy = zx - ax, zy - ay
            L2 = dx * dx + dy * dy
            if L2 <= 0:
                return math.dist(p, a)
            t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
            return math.dist(p, (ax + t * dx, ay + t * dy))

        TOL = 300.0          # mm — 노드정리·수직전개의 자리 흔들림 여유

        def uncovered(src_segs, dst_segs):
            out = []
            for (a, z) in src_segs:
                pts_ = [a, z, ((a[0] + z[0]) / 2, (a[1] + z[1]) / 2)]
                if all(any(seg_dist(p, c, d) <= TOL for (c, d) in dst_segs)
                       for p in pts_):
                    continue
                out.append((a, z))
            return out

        csegs = [(tuple(b.pts[e[0]][:2]), tuple(b.pts[e[1]][:2]))
                 for e in corridor]
        g_lost = uncovered(csegs, dsegs)
        g_extra = uncovered(dsegs, csegs)
        tot = lambda ss: sum(math.dist(a, z) for a, z in ss) / 1000.0
        print(f"\n  ── 기하 비교 (문턱 {TOL:.0f}mm) ──")
        print(f"  corridor 선분 {len(csegs)} · {tot(csegs):.1f} m"
              f"  /  표 선분 {len(dsegs)} · {tot(dsegs):.1f} m")
        print(f"  ★표가 안 덮는 corridor 선분 : {len(g_lost)} · {tot(g_lost):.1f} m")
        print(f"  ★corridor 에 없는 표 선분   : {len(g_extra)} · {tot(g_extra):.1f} m")
        for (a, z) in sorted(g_lost, key=lambda s: -math.dist(*s))[:6]:
            print(f"      빠짐 ({a[0]:.0f},{a[1]:.0f})-({z[0]:.0f},{z[1]:.0f})"
                  f" · {math.dist(a, z) / 1000:.2f} m")
        for (a, z) in sorted(g_extra, key=lambda s: -math.dist(*s))[:6]:
            print(f"      덤   ({a[0]:.0f},{a[1]:.0f})-({z[0]:.0f},{z[1]:.0f})"
                  f" · {math.dist(a, z) / 1000:.2f} m")

        lost = corridor - in_table
        extra = in_table - corridor
        print(f"\n  ★corridor 에 있는데 표에 없는 간선 : {len(lost)}"
              f" · {mlen(b.pts, lost):.1f} m")
        print(f"  ★표에 있는데 corridor 에 없는 간선 : {len(extra)}"
              f" · {mlen(b.pts, extra):.1f} m")
        for e in sorted(lost)[:6]:
            p, q = b.pts[e[0]], b.pts[e[1]]
            print(f"      빠짐 {e} ({p[0]:.0f},{p[1]:.0f})-({q[0]:.0f},{q[1]:.0f})"
                  f" · {math.dist(p[:2], q[:2]) / 1000:.2f} m")
        for e in sorted(extra)[:6]:
            p, q = b.pts[e[0]], b.pts[e[1]]
            print(f"      덤   {e} ({p[0]:.0f},{p[1]:.0f})-({q[0]:.0f},{q[1]:.0f})"
                  f" · {math.dist(p[:2], q[:2]) / 1000:.2f} m")

        meta = dict(tbl.meta)
        print(f"\n  [연장] corridor {mlen(b.pts, corridor):.1f} m"
              f" · 표 meta 총연장 {meta.get('총 배관연장 (m)') or meta.get('총연장 (m)') or '?'}")
        print(f"  [노즐] {len(tbl.nozzles)} · [절점] {len(tbl.nodes)}"
              f" · [배관] {len(tbl.pipes)}")
        ok = not g_lost
        print("\n  " + ("★corridor 가 표에 그대로 왔다 (기하 기준)" if ok else
                        "★★corridor 의 일부가 표에 없다 — 사용자가 본 그 증상"))
        return 0 if ok else 2


if __name__ == "__main__":
    raise SystemExit(main())
