# -*- coding: utf-8 -*-
"""[최불리] 기준개수 K 를 넣었는데 표에 노즐이 K개보다 적게 온다 — 몇 개가 왜.

사용자 지적(2026-09-09): 「기준개수 30개로 헤드 배관망을 추출해도 30개 밑으로
헤드가 검출된다」.

★가설: 선정이 «같은 자리» 를 둘로 세는데 표는 하나로 만든다::

    design/worst.py
        ranked = sorted(head_node, …)
        picked = ranked[:k]      # 같은 노드를 문 헤드가 둘이면 둘 다 센다
    convert/planar.py
        head_vid[vid] = (hx, hy) # 같은 노드면 뒤엣것이 덮는다 → 노즐 하나

여기서는 **관문마다** 세어 그 가설을 확인하거나 버린다::

    ⑴ 손질판 헤드            len(board.disks)
    ⑵ 선정이 고른 헤드        len(worst["heads"])
    ⑶ 그중 «서로 다른 자리»   부착 노드 기준 · 좌표 기준
    ⑷ 표에 온 노즐            len(tables.nozzles)

    python scripts/_probe_k_shortfall.py [--k 30]
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from collections import Counter
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


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
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
        rec = ((c.get(f"/api/module-f/recon?sid={sid}").get_json() or {})
               .get("recon") or {})
        c.post("/api/module-f/pick/adopt",
               json={"sid": sid, "materials": True,
                     "heads": {"conf_min": (rec.get("adopt") or {})
                               .get("conf_min")}})
        wait(c, sid)
        c.post("/api/module-f/pick/commit", json={"sid": sid})
        if wait(c, sid).get("state") != "done":
            print("★손질 실패")
            return 1

        from routes.module_f.jobs import _sess
        sess = _sess(sid)
        b = sess["edit"].board
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)
        c.post("/api/module-f/edit/worst", json={"sid": sid, "k": args.k})
        w = sess.get("worst") or {}
        picked = list(w.get("heads") or ())

        print(f"\n■ 대명동 · 기준개수 K={args.k}")
        print(f"  ⑴ 손질판 헤드      {len(b.disks)}")
        print(f"  ⑵ 선정이 고른 헤드  {len(picked)}"
              f"  (닿는 헤드 {w.get('reachable')})")

        # ⑶ 그중 «서로 다른 자리» 인가 — 두 가지 자로 잰다.
        xy = Counter((round(float(b.disks[hi][0]), 1),
                      round(float(b.disks[hi][1]), 1))
                     for hi in picked if hi < len(b.disks))
        node = Counter()
        for hi in picked:
            ns = b.hnodes[hi] if hi < len(b.hnodes) else ()
            node[tuple(sorted(ns))] += 1
        dup_xy = [(p, n) for p, n in xy.items() if n > 1]
        dup_nd = sum(n - 1 for n in node.values() if n > 1)
        print(f"  ⑶ 서로 다른 «좌표»  {len(xy)}"
              f"   (겹친 자리 {len(dup_xy)}곳 · 중복 {sum(n - 1 for _p, n in dup_xy)})")
        print(f"     서로 다른 «노드»  {len(node)}   (중복 {dup_nd})")
        for p, n in dup_xy[:8]:
            print(f"       겹침 {n}개 — {p}")

        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        if wait(c, sid).get("state") != "done":
            print("★표 확정 실패")
            return 1
        tbl = sess["design"]["tables"]
        print(f"  ⑷ 표에 온 노즐      {len(tbl.nozzles)}"
              f"   ← 여기서 {len(picked) - len(tbl.nozzles)}개가 모자란다")
        print(f"\n  가설 검증 — «서로 다른 좌표» {len(xy)} 와 «노즐»"
              f" {len(tbl.nozzles)} 가 같은가: "
              + ("**같다 → 겹침이 원인**" if len(xy) == len(tbl.nozzles)
                 else "다르다 → 겹침 말고 다른 원인도 있다"))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
