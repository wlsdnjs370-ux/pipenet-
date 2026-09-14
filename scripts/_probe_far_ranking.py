# -*- coding: utf-8 -*-
"""그려지는데 최불리 K 에 안 드는 헤드가 있나 — 순위를 통째로 낸다.

  「가장 먼 헤드가 안 뽑힌다」를 짐작으로 고치지 않기 위한 자다. 도면에 그려진
  헤드 **전부**를 놓고 세 가지를 한 줄씩 낸다:

      · 급수원에서 배관 따라 잰 유하거리 순위 (상위 N)
      · 후보에 **못 든** 헤드 — 자리와 이유 (여기 있으면 K 에 들 수가 없다)
      · 가장 먼 헤드가 상위 K 안에 있나

  ★후보에 못 드는 길은 셋뿐이다:
      ① 영역·도면 장 밖          (사람이 가둔 것 — 규칙대로다)
      ② 급수원에서 배관이 안 닿음  (`hnodes` 중 도달 노드가 없다)
      ③ 다른 헤드와 «같은 자리»   (겹쳐 그림 — 대표 하나로 접힌다)
  그 밖의 이유로 빠지면 그것이 결함이다(`ModuleF_최불리규칙_복원_지시서.md` §0).

    python scripts/_probe_far_ranking.py [--k 30] [--top 25] [--key 저장본]
"""
from __future__ import annotations

import argparse
import math
import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core"), str(ROOT / "cad_project_editor_g"),
           str(ROOT / "tests"), str(ROOT / "scripts")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from _probe_candidate_drop import _Null, setup                # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--k", type=int, default=30)
    ap.add_argument("--top", type=int, default=25)
    ap.add_argument("--key", default="")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    ctx = _Null() if args.key else isolated_workdir(prefix="far_")
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, argparse.Namespace(key=args.key))

        from routes.module_f.jobs import _sess
        from routes.module_f.remote30 import _worst_k_heads
        from routes.module_f.attach import wet_heads
        sess = _sess(sid)
        es = sess["edit"]
        b = es.board
        est = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
        if not (est.get("sources") or ()):
            seg = sorted((g.get("segs") or [] for g in est["body_groups"]),
                         key=len, reverse=True)[0]
            c.post("/api/module-f/edit/mode",
                   json={"sid": sid, "mode": "급수시작위치"})
            c.post("/api/module-f/edit/click",
                   json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                         "y": (seg[1] + seg[3]) / 2, "max_d": 2000})
        src_index = 0 if b.sources else None
        n = len(b.disks)

        full = _worst_k_heads(b.pts, b.edges, b.hnodes, b.sources,
                              k=max(1, n), only_heads=None,
                              source_index=src_index, head_xy=b.disks)
        order = list(full["heads"])
        far = {int(h): float(v) for h, v in (full.get("dists") or {}).items()}
        probe = wet_heads(sess, es)
        reason = probe.get("reason") or {}
        wet = {int(i) for i in (probe.get("wet") or ())}

        def xy(i):
            return float(b.disks[i][0]), float(b.disks[i][1])

        print(f"\n■ 유하거리 순위 · {args.key or '대명동 단위세대'}"
              f" · 도면 헤드 {n}개 · 후보(대표) {len(order)}개 · K={args.k}")
        print(f"\n[상위 {min(args.top, len(order))}]  등수 · disk · 자리(mm)"
              f" · 유하거리 · K 안? · 물닿음?")
        for r, h in enumerate(order[:args.top], 1):
            x, y = xy(h)
            print(f"  {r:>4}  disk {h:<5} ({x:>9.0f}, {y:>10.0f})"
                  f" {far.get(h, 0) / 1000.0:>8.2f} m"
                  f"  {'K안' if r <= args.k else '  —'}"
                  f"  {'물닿음' if h in wet else '★마름'}")

        # ── 후보에 못 든 헤드 — 여기 있으면 K 에 들 길이 없다
        missing = [i for i in range(n) if i not in set(order)]
        print(f"\n[후보 밖]  {len(missing)}개 — 그려는 지는데 순위에 못 든 헤드")
        same = 0
        for i in missing:
            x, y = xy(i)
            twin = next((j for j in order
                         if math.dist((x, y), xy(j)) <= 0.15), None)
            if twin is not None:
                same += 1
                if same <= 4:
                    print(f"    disk {i:<5} ({x:>9.0f}, {y:>10.0f})"
                          f"  ③ 겹쳐 그림 — disk {twin} 과 같은 자리")
            else:
                why = str(reason.get(i) or reason.get(str(i)) or "?")
                print(f"    disk {i:<5} ({x:>9.0f}, {y:>10.0f})"
                      f"  ★② 급수원에서 배관이 안 닿음 · {why}")
        if same > 4:
            print(f"    … ③ 겹쳐 그림 {same}개 (자리 하나로 접힘 — 규칙대로다)")

        # ── 도면에서 «가장 멀리 있는» 헤드가 1등인가
        if order:
            top1 = order[0]
            x, y = xy(top1)
            print(f"\n[1등]      disk {top1} ({x:.0f}, {y:.0f})"
                  f" · {far[top1] / 1000.0:.2f} m")
            print(f"[K 경계]   {args.k}등 {far[order[args.k - 1]] / 1000.0:.2f} m"
                  f"  /  {args.k + 1}등 "
                  + (f"{far[order[args.k]] / 1000.0:.2f} m"
                     if len(order) > args.k else "(없음)"))
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
