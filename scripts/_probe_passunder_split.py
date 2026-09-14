# -*- coding: utf-8 -*-
"""[최불리규칙복원 §2-2] 못 붙은 헤드를 `stage5_split_through_uprights` 가 붙이나.

지시서 §2-2 는 「새 판정을 만들지 말고 그 함수를 다시 부른다」고 한다. 부르기
전에 **그 함수가 그 헤드들을 실제로 집는지** 재 본다 — 그 함수는 제 안에서
`has_near_node`(중심 ≤ARM_CTR **또는 테두리 ±HEAD_TOUCH**)로 «이미 붙은 헤드»
를 건너뛰는데, `pass_under` 는 정의상 **테두리에 노드가 있는** 경우라 그 문에
걸릴 수 있기 때문이다.

    attach_heads_center 의 사유
      no_center   중심에도 테두리에도 «간선 달린» 노드가 없다   ← has_near_node 거짓
      chord_only  테두리 끝이 있었지만 양끝이 다 원 위였다       ← has_near_node 참
      pass_under  테두리에 노드는 있는데 «선이 끝나는» 자리가 아니다 ← has_near_node 참(?)

    python scripts/_probe_passunder_split.py [--key 저장본]
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
    ap.add_argument("--key", default="")
    args = ap.parse_args()
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    ctx = _Null() if args.key else isolated_workdir(prefix="pusp_")
    with ctx, srv.app.test_client() as c:
        with c.session_transaction() as s:
            s["authed"] = True
        sid = setup(c, argparse.Namespace(key=args.key))

        from routes.module_f.jobs import _sess
        from services.cad_import.pipeline import flow as fw
        b = _sess(sid)["edit"].board
        pts, edges, disks = list(b.pts), set(b.edges), list(b.disks)

        why: dict = {}
        p2, e2, centers, n_wire, _multi = fw.attach_heads_center(
            pts, edges, disks, why=why)
        miss = [i for i, ctr in enumerate(centers) if ctr is None]
        by = {}
        for i in miss:
            by[why.get(i, "?")] = by.get(why.get(i, "?"), 0) + 1
        print(f"\n■ 못 붙은 헤드 · {args.key or '대명동'}"
              f" · 헤드 {len(disks)}개")
        print(f"\n[못 붙음]   {len(miss)}개 — "
              + " · ".join(f"{k} {v}" for k, v in sorted(by.items())))

        # ── 그 함수의 문(has_near_node)이 이 헤드들을 어떻게 보나
        used = {n for e in e2 for n in e}

        def near_node(hx, hy, hr):
            best = None
            for n in used:
                d = math.hypot(p2[n][0] - hx, p2[n][1] - hy)
                if best is None or d < best[0]:
                    best = (d, n)
                if d <= fw.ARM_CTR or abs(d - hr) <= fw.HEAD_TOUCH:
                    return True, d
            return False, (best[0] if best else -1)

        def passes_through(hx, hy, hr):
            """원 안(횡이탈 ≤ r)을 지나는 관이 있나 — 그 함수의 둘째 조건."""
            for (i, j) in e2:
                ax, ay = p2[i][:2]
                bx, by = p2[j][:2]
                L = math.hypot(bx - ax, by - ay)
                if L < 1e-9:
                    continue
                t = ((hx - ax) * (bx - ax) + (hy - ay) * (by - ay)) / (L * L)
                if t <= 1e-6 or t >= 1.0 - 1e-6:
                    continue
                px, py = ax + (bx - ax) * t, ay + (by - ay) * t
                if math.hypot(px - hx, py - hy) <= hr:
                    return True
            return False

        stat = {}
        for i in miss[:400]:
            hx, hy, hr = (float(disks[i][0]), float(disks[i][1]),
                          float(disks[i][2]))
            blocked, d = near_node(hx, hy, hr)
            thru = passes_through(hx, hy, hr)
            key = (why.get(i, "?"), blocked, thru)
            stat[key] = stat.get(key, 0) + 1
        print(f"\n[그 함수의 눈]  사유 · has_near_node · 원 안 통과관 → 개수")
        for (w, blocked, thru), n in sorted(stat.items()):
            verdict = ("건너뜀(이미 붙었다고 본다)" if blocked else
                       ("쪼갠다" if thru else "통과관이 없어 못 쪼갬"))
            print(f"    {w:<12} near={str(blocked):<5} thru={str(thru):<5}"
                  f" → {n:>4}개   {verdict}")

        # ── 실제로 불러 본다 (지시서가 시킨 그대로)
        ups = [(float(disks[i][0]), float(disks[i][1]), float(disks[i][2]))
               for i in miss]
        p3, e3, n_split = fw.stage5_split_through_uprights(p2, e2, ups)
        print(f"\n[불러 봄]   stage5_split_through_uprights(못 붙은 {len(ups)}개)"
              f" → 쪼갠 관 {n_split}개")
        if n_split:
            why2: dict = {}
            _p, _e, c2, _w, _m = fw.attach_heads_center(
                p3, e3, disks, why=why2)
            miss2 = [i for i, ctr in enumerate(c2) if ctr is None]
            by2 = {}
            for i in miss2:
                by2[why2.get(i, "?")] = by2.get(why2.get(i, "?"), 0) + 1
            print(f"[다시 붙임] 못 붙음 {len(miss)} → {len(miss2)}개 — "
                  + (" · ".join(f"{k} {v}" for k, v in sorted(by2.items()))
                     or "(없음)"))
        return 0


if __name__ == "__main__":
    raise SystemExit(main())
