# -*- coding: utf-8 -*-
"""[수동 찍기] 사람이 찍은 배관·헤드가 수리계산까지 «몇 개» 살아 오는가.

사용자 지적: 「02.찍기에서 수동으로 정의한 배관이랑 헤드를 04.수리계산에서
반영을 못하는 것 같다」.

★추측하지 않고 **관문마다 센다.** 잃어버리는 자리가 어디인지는 숫자가 말한다::

    ⑴ 찍기      스펙의 재료 묶음 · 헤드 묶음
    ⑵ 세계      그 묶음에 실제로 든 선분 수 · 원 수
    ⑶ 손질판    board.pts/edges/disks — 찍은 선분이 간선으로 섰는가
    ⑷ 표        노즐·배관 — 손질 헤드가 표까지 왔는가

자동(추천 채택)이 아니라 **클릭**으로 찍는다. 두 길이 갈릴 수 있으므로
`--how auto` 로 견줄 수도 있다.

    python scripts/_probe_manual_pick_flow.py [--k 30] [--how click|auto]
"""
from __future__ import annotations

import argparse
import math
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


def wait(c, sid, limit=40000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def _mid(a, b):
    return ((a[0] + b[0]) / 2.0, (a[1] + b[1]) / 2.0)


def _near(pts, q, eps):
    """q 에 eps 안으로 붙는 pts 인덱스 — 없으면 None."""
    best, bi = eps, None
    for i, p in enumerate(pts):
        d = math.hypot(p[0] - q[0], p[1] - q[1])
        if d < best:
            best, bi = d, i
    return bi


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--how", default="click", choices=("click", "auto"))
    ap.add_argument("--eps", type=float, default=60.0)
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    if not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0
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
        wait(c, sid)

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        ps = sess["pick"]
        board = ps.board
        world = sess.get("world") or {}

        print("\n■ ⑴ 찍기 — 사람이 «클릭» 으로 정의한다"
              if args.how == "click" else "\n■ ⑴ 찍기 — 추천 채택")
        if args.how == "auto":
            c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "PIPE"})
            c.post("/api/module-f/pick/mode",
                   json={"sid": sid, "action": "complete"})
            c.post("/api/module-f/pick/auto", json={"sid": sid, "cat": "HEAD"})
        else:
            c.post("/api/module-f/pick/mode", json={"sid": sid,
                                                    "action": "pipe"})
            # 추천 PIPE 묶음의 «실제 선분 중점» 을 하나씩 클릭한다.
            want = [tuple(b["key"]) if b.get("key") else
                    (b.get("layer"), b.get("color"))
                    for b in (world.get("bundles") or [])
                    if b.get("cat") == "PIPE"]
            done = []
            for key in want:
                segs = board.by_bundle.get(key) or []
                if not segs:
                    continue
                mx, my = _mid(*segs[len(segs) // 2])
                rep = (c.post("/api/module-f/pick/click",
                              json={"sid": sid, "x": mx, "y": my})
                       .get_json() or {}).get("report") or {}
                done.append((key, rep.get("픽"), rep.get("동작")))
            print(f"    배관 클릭 {len(done)}회 — "
                  + " · ".join(f"{k[0]}×{k[1]}→{p}" for k, p, _a in done[:4]))
            c.post("/api/module-f/pick/mode",
                   json={"sid": sid, "action": "complete"})
            # 헤드 — 추천 HEAD 묶음의 원 중심을 클릭한다.
            hb = [tuple(b["key"]) if b.get("key") else
                  (b.get("layer"), b.get("color"))
                  for b in (world.get("bundles") or [])
                  if b.get("cat") == "HEAD"]
            hits = 0
            for key in hb:
                cs = [t for t in board._csmall
                      if (t[0], t[1]) == (key[0], key[1])]
                if not cs:
                    continue
                _ly, _co, cx, cy, _r = cs[len(cs) // 2]
                rep = (c.post("/api/module-f/pick/click",
                              json={"sid": sid, "x": cx, "y": cy})
                       .get_json() or {}).get("report") or {}
                if rep:
                    hits += 1
            print(f"    헤드 클릭 {hits}회")

        sp = ps.spec()
        # ★칸 이름은 `material_picks` 다. `materials` 로 읽으면 0 이 나오고,
        #   그 0 을 손실이라 말하면 멀쩡한 것을 고치러 간다(한 번 그랬다).
        mats = [tuple(m) for m in (sp.get("material_picks") or ())]
        heads = list(sp.get("heads") or ())
        print(f"    스펙 — 재료 묶음 {len(mats)} · 헤드 묶음 {len(heads)}")

        # ⑵ 세계 — 그 묶음에 실제로 든 선분·원
        pick_segs = []
        for key in mats:
            pick_segs += list(board.by_bundle.get(key) or ())
        hkeys = {(tuple(h["bundle"])[0], tuple(h["bundle"])[1],
                  round(float(h.get("r") or 0.0), 1))
                 for h in heads if "tri_side" not in h}
        pick_circ = [t for t in board._csmall
                     if (t[0], t[1], round(t[4], 1)) in hkeys]
        print(f"\n■ ⑵ 세계 — 찍힌 선분 {len(pick_segs)} · 찍힌 원"
              f" {len(pick_circ)}")

        # ⑶ 손질판
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★손질 실패 — {j}")
            return 1
        es = sess["edit"]
        b = es.board
        print(f"\n■ ⑶ 손질판 — 절점 {len(b.pts)} · 간선 {len(b.edges)}"
              f" · 헤드 {len(b.disks)}")

        # 찍은 선분의 «양 끝» 이 손질 절점으로 섰는가 (접기·합치기로 중간점은
        # 사라질 수 있다 — 그래서 «끝점» 을 본다).
        miss_seg = 0
        for a, bb in pick_segs:
            if _near(b.pts, a, args.eps) is None or \
                    _near(b.pts, bb, args.eps) is None:
                miss_seg += 1
        print(f"    찍은 선분 중 «끝점이 손질판에 없는» 것 {miss_seg}"
              f" / {len(pick_segs)}")
        miss_h = 0
        for (_ly, _co, cx, cy, _r) in pick_circ:
            if _near([(d[0], d[1]) for d in b.disks], (cx, cy),
                     args.eps) is None:
                miss_h += 1
        print(f"    찍은 원 중 «손질 헤드가 안 선» 것 {miss_h}"
              f" / {len(pick_circ)}")

        # ⑷ 표
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": args.k})
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {j}")
            return 1
        tbl = sess["design"]["tables"]
        sm = (sess.get("design") or {}).get("summary") or {}
        print(f"\n■ ⑷ 표 — 절점 {len(tbl.nodes)} · 배관 {len(tbl.pipes)}"
              f" · 노즐 {len(tbl.nozzles)}")
        print(f"    인계 {sm.get('handoff')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
