# -*- coding: utf-8 -*-
"""[비용] «붙는 헤드» 를 선정 «전» 에 재면 얼마나 더 드나 — 큰 도면으로.

사용자 지적 둘이 여기서 만난다::

    ⑴ 「사전에 체크한 30개와 거기 연결된 배관만 가져오면 되는 것 아닌가」
       → 고를 때부터 «전개가 붙일 수 있는 헤드» 만 봐야 K 가 진짜 K 가 된다.
    ⑵ 「지하주차장 도면에서 수리계산이 너무 오래 걸린다」
       → 그 «붙는 헤드» 판정이 곧 **전체망 전개 한 번**이다.

  둘은 같은 줄이다. 그래서 «옮기는 것이 이득인가» 를 재야 한다::

      끔(MF_NO_ATTACH_PRESELECT=1)   손질 최불리 빠름 · 수리계산이 전개 1회
      켬(기본)                        손질 최불리가 전개 1회 · 수리계산 즉시

  총합은 같고 **자리만 옮긴다.** 다만 수리계산을 두 번 이상 누르면(관경·K를
  바꿔 가며) 켠 쪽이 유리하다 — 캐시가 판마다 한 번만 재기 때문이다.

    python scripts/_probe_attach_cost.py [--k 30]
"""
from __future__ import annotations

import argparse
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

B1F = "B1F 현장조사 소화설비 평면도"


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
    ap.add_argument("--key", default=B1F)
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    off = os.environ.get("MF_NO_ATTACH_PRESELECT") == "1"
    print(f"\n■ 붙는 헤드 «선정 전» 판정 — {'끔(종전)' if off else '켬(기본)'}"
          f" · {args.key} · K={args.k}")

    with srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        r = c.post("/api/module-f/reopen", json={"key": args.key})
        j0 = r.get_json() or {}
        if not j0.get("ok"):
            print(f"★이어서 열기 실패 — {str(j0)[:160]}")
            return 0
        sid = j0["sid"]
        if wait(c, sid).get("state") != "done":
            print("★열기 잡 실패")
            return 1
        st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        print(f"    손질판 — 절점 {st['counts']['pts']:,}"
              f" · 간선 {st['counts']['edges']:,}"
              f" · 헤드 {st['counts']['heads']:,}")
        if not (st.get("sources") or ()):
            print("    ★급수원이 없어 못 잽니다 — 저장본에 알람밸브가 있어야 합니다.")
            return 0

        t0 = time.perf_counter()
        w = c.post("/api/module-f/edit/worst",
                   json={"sid": sid, "k": args.k}).get_json() or {}
        t_worst = time.perf_counter() - t0
        sm = w.get("summary") or {}
        print(f"    ① 손질 최불리   {t_worst:6.1f}s"
              f"  · 고른 헤드 {sm.get('k')}"
              f" · 후보 {sm.get('candidates')}"
              f" · 못 붙는 헤드 {sm.get('unattached')}")

        t0 = time.perf_counter()
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        jb = wait(c, sid)
        t_dg = time.perf_counter() - t0
        if jb.get("state") != "done":
            print(f"    ★표 확정 실패 — {str(jb)[:160]}")
            return 1
        from routes.module_f.jobs import _sess
        tbl = _sess(sid)["design"]["tables"]
        print(f"    ② 수리계산     {t_dg:6.1f}s  · 표에 온 노즐"
              f" **{len(tbl.nozzles)}** / K={args.k}")

        t0 = time.perf_counter()
        c.post("/api/module-f/design/build", json={"sid": sid, "k": args.k})
        wait(c, sid)
        t_dg2 = time.perf_counter() - t0
        print(f"    ③ 수리계산 재실행 {t_dg2:6.1f}s  (같은 판 — 여기가 캐시 이득)")
        print(f"\n    합계 {t_worst + t_dg:6.1f}s"
              f"  (한 번 더 누르면 +{t_dg2:.1f}s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
