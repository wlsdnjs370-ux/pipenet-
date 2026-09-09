# -*- coding: utf-8 -*-
"""[최불리 인계] 손질이 고른 K개를 수리계산이 그대로 받는가 — 지시서 §3-1.

■ 증상 (사용자 화면 · B1F)

  손질에서 영역을 지정해 K개를 골랐는데, 수리계산은 그 선정을 **버리고** 도면
  전체에서 K개를 다시 뽑는다. 밑그림을 켜면 두 망이 정확히 겹치므로 좌표는
  무결하고 **어느 헤드인가만** 다르다.

      routes/module_f/api_design.py:539
          got = select_and_expand(payload, es.board, k=cfg["k"],
                                  selected_source=sel)     # ★only_heads 없음

  `select_and_expand` 는 `only_heads=None` 이면 `cand = wet`(도면 전체의 물닿는
  헤드)에서 K 개를 다시 뽑는다.

■ 인덱스 공간 — 셋이 같은가 (§2-2 · **이것부터 확인한다**)

      손질  w["heads"]      = {hi for hi, d in enumerate(b.disks) …}
      payload  data["hcov"] = [list(d) for d in b.disks]
      planar   wet_head_idx = enumerate(hcov) 의 인덱스

  코드상 셋 다 `board.disks` 순서다. 그래도 **좌표로 재서** 확인한다 —
  어긋난 채 넘기면 엉뚱한 헤드 K개가 조용히 선정된다(지금보다 나쁘다).

    python scripts/_probe_worst_handoff.py [--key 저장본] [--k 20] [--tag 조치전]
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


def wait(c, sid, limit=200000):
    for _ in range(limit):
        j = c.get(f"/api/module-f/job?sid={sid}").get_json()
        if j.get("state") in ("done", "error", "idle"):
            return j
        time.sleep(0.1)
    return {"state": "timeout"}


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--key", default="")
    ap.add_argument("--k", type=int, default=20)
    ap.add_argument("--tag", default="")
    ap.add_argument("--frac", type=float, default=0.55,
                    help="영역 박스가 도면 bbox 의 몇 배인가")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        if args.key:
            r = c.post("/api/module-f/reopen", json={"key": args.key})
            j0 = r.get_json() or {}
            if not j0.get("ok"):
                print(f"★이어서 열기 실패 — {str(j0)[:200]}")
                return 1
            sid = j0["sid"]
            if wait(c, sid).get("state") != "done":
                print("★이어서 열기 잡 실패")
                return 1
        else:
            if not PLAN.is_file():
                print(f"표본 없음: {PLAN}")
                return 0
            with open(PLAN, "rb") as fh:
                r = c.post("/api/module-f/slot/open",
                           data={"dxf_file": (fh, PLAN.name), "kind": "plan"},
                           content_type="multipart/form-data")
            sid = (r.get_json() or {})["sid"]
            wait(c, sid)
            c.post("/api/module-f/slot/read",
                   json={"sid": sid, "method": "manual"})
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
        es = sess["edit"]
        b = es.board
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (st.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/anchor-click",
                   json={"sid": sid, "x": seg[0], "y": seg[1]})
            wait(c, sid)

        disks = [(float(d[0]), float(d[1])) for d in b.disks]
        xs = [p[0] for p in disks]
        ys = [p[1] for p in disks]
        cx, cy = (min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2
        hw = (max(xs) - min(xs)) * args.frac / 2
        hh = (max(ys) - min(ys)) * args.frac / 2
        zone = [cx - hw, cy - hh, cx + hw, cy + hh]
        print(f"\n■ 최불리 인계 계측 {args.tag}"
              f" · 도면 헤드 {len(disks)} · K={args.k}")
        print(f"  영역 박스 ({zone[0]:.0f},{zone[1]:.0f})–"
              f"({zone[2]:.0f},{zone[3]:.0f})")

        r = c.post("/api/module-f/edit/worst",
                   json={"sid": sid, "k": args.k, "zones": [zone]})
        jw = r.get_json() or {}
        if not jw.get("ok"):
            print(f"★최불리 선정 실패 — {str(jw)[:260]}")
            return 1
        w = sess.get("worst") or {}
        heads = [int(i) for i in (w.get("heads") or ())]
        print(f"  [손질] 고른 헤드 {len(heads)}개 · 영역 안 후보"
              f" {w.get('candidates')} · 닿는 헤드 {w.get('reachable')}")
        print(f"         기준 헤드 disk {w.get('worst_head')}"
              f" · 최원 {w.get('far_m')} m"
              f" · corridor 간선 {len(w.get('loads') or {})}")

        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        j = wait(c, sid)
        if j.get("state") != "done":
            print(f"★표 확정 실패 — {str(j)[:200]}")
            return 1
        d = sess["design"]
        tbl, got = d["tables"], d["got"]
        noz = [str(z.get("in")) for z in tbl.nozzles]
        at = {str(n["label"]): (float(n.get("x", 0) or 0),
                                float(n.get("y", 0) or 0)) for n in tbl.nodes}
        origin = got.get("origin_mm")
        print(f"  [설계] 표 헤드(노즐) {len(noz)}개"
              f" · 제외 {got.get('excluded_heads')}"
              f" · 후보 {got.get('candidate_heads')}"
              f" / 도면 {got.get('total_heads')}")

        # ── 좌표 1:1 대응 — 손질 disk ↔ 설계 헤드 노드
        #
        #   표 좌표는 재원점된 mm 다: (board − origin) + 1000
        def to_tbl(p):
            return (p[0] - float(origin[0]) + 1000.0,
                    p[1] - float(origin[1]) + 1000.0)

        want = [to_tbl(disks[i]) for i in heads]
        have = [at[n] for n in noz if n in at]
        used, gaps, miss = set(), [], 0
        for q in want:
            best, bd = None, 1e18
            for k2, p in enumerate(have):
                if k2 in used:
                    continue
                dd = math.dist(p, q)
                if dd < bd:
                    best, bd = k2, dd
            if best is None or bd > 100.0:
                miss += 1
                continue
            used.add(best)
            gaps.append(bd)
        gaps.sort()
        print(f"  ★[인덱스 확인] 손질 disk ↔ 설계 헤드"
              f" **{len(gaps)}/{len(heads)}** 대응"
              + (f" · 최대 오차 {gaps[-1]:.1f} mm"
                 f" (중앙 {gaps[len(gaps) // 2]:.1f})" if gaps else "")
              + (f" · 못 찾음 {miss}" if miss else ""))
        if len(gaps) != len(heads):
            print("     ★★인덱스 공간이 다를 수 있다 — 지시서 §2-2 는 이때"
                  " «멈추고 보고» 하라고 했다.")

        # ── 기준 헤드·최원 경로가 같은가
        wh = w.get("worst_head")
        wh_xy = to_tbl(disks[int(wh)]) if wh is not None else None
        meta = dict(tbl.meta)
        print(f"  [기준 헤드] 손질 disk {wh}"
              f" → 표 좌표 ({wh_xy[0]:.0f},{wh_xy[1]:.0f})" if wh_xy else "")
        near = None
        if wh_xy:
            near = min(((math.dist(at[n], wh_xy), n) for n in noz if n in at),
                       default=(None, None))
            print(f"             설계 표에서 가장 가까운 노즐 노드"
                  f" {near[1]} · {near[0]:.1f} mm"
                  if near[0] is not None else "             (못 찾음)")
        print(f"  [최원] 손질 far_m {w.get('far_m')}"
              f" · 설계 meta 최원 {meta.get('최원 유하거리 (m)')}")
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
