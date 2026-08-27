# -*- coding: utf-8 -*-
"""채택이 왜 느려지는가 — 클릭당 비용이 «일정한가 늘어나는가» 를 잰다.

실측이 어긋났다. F-8b 에서 B1F 후보 72개를 0.4s 에 찍었는데(클릭당 5.5ms),
F-8e 에서 3,235개를 1,683s 에 찍었다(클릭당 260ms). 47배다. 후보가 늘면
클릭당 비용도 같이 느는 것 — 즉 어딘가 O(N) 이 클릭 «안» 에 있다는 뜻이다.

여기서는 추측하지 않고 구간별로 잰다: 처음 100개, 다음 100개, … 클릭당
평균이 계단처럼 오르면 누적 상태를 매번 훑는 자리가 있는 것이다.

    python scripts/_probe_adopt_cost.py [도면.dxf] [--n 800] [--step 100]
"""
from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--n", type=int, default=800)
    ap.add_argument("--step", type=int, default=100)
    a = ap.parse_args()

    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    from routes.module_f.common import _boot
    _boot()
    from routes.module_f.adopt import ADOPT_MAX_D_MM
    from routes.module_f.recon import run_recon
    from services.cad_import.pick.session import PickSession

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    t0 = time.perf_counter()
    ps = PickSession.open(str(dxf))
    ps.select_pipe()
    print(f"찍기판 {time.perf_counter() - t0:.1f}s")

    rec = run_recon(dxf, tag="측정")
    cands = rec["heads"]
    print(f"후보 {len(cands):,}개\n")

    # 재료 — 레이어 사전이 PIPE 로 본 묶음의 선분 중점을 찍는다.
    from routes.module_f.common import _layer_category
    got = 0
    for (ly, c), segs in list(ps.board.by_bundle.items()):
        if _layer_category(str(ly)) != "PIPE" or not segs:
            continue
        p, q = segs[0]
        r = ps.click((p[0] + q[0]) / 2.0, (p[1] + q[1]) / 2.0)
        if r and r.get("동작") == "추가":
            got += 1
    print(f"재료 {got}묶음")
    if not ps.complete_pipe():
        print("재료 0 — 중단")
        return 1
    ps.set_slot(ps.head_label)

    print(f"\n{'구간':>12} {'클릭':>7} {'초':>8} {'클릭당ms':>10} "
          f"{'picks':>7} {'추가':>6} {'취소':>6} {'실패':>6}")
    print("-" * 68)
    n = min(a.n, len(cands))
    for lo in range(0, n, a.step):
        hi = min(lo + a.step, n)
        t = time.perf_counter()
        clicks = add = cancel = bad = 0
        for cd in cands[lo:hi]:
            rep = ps.click(cd["x"], cd["y"], max_d=ADOPT_MAX_D_MM)
            clicks += 1
            act = (rep or {}).get("동작")
            if act == "추가":
                add += 1
            elif act == "취소":
                ps.click(cd["x"], cd["y"], max_d=ADOPT_MAX_D_MM)
                clicks += 1
                cancel += 1
            else:
                bad += 1
        el = time.perf_counter() - t
        print(f"{lo:>6,}~{hi:<5,} {clicks:>7,} {el:>8.2f} "
              f"{el / max(1, clicks) * 1000:>10.1f} "
              f"{len(ps.board.heads):>7,} {add:>6} {cancel:>6} {bad:>6}")

    print(f"\n누적 picks {len(ps.board.heads):,} · "
          f"클릭 기록 {len(ps.board.clicks):,}")
    print("클릭당 ms 가 구간마다 오르면 누적 상태를 매번 훑는 자리가 있는 것이다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
