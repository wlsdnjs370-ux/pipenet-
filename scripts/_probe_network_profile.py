# -*- coding: utf-8 -*-
"""배관망 검출 41.7초가 «어디로» 가는가 — cProfile 로 연다.

실측(`_probe_module_f_hot.py`):

    B1F     망 검출  41,670 ms   ← 나머지 전부의 90배
    대명동   망 검출     756 ms

`run_network` 는 A 의 `select_worst30_heads_anchored` 를 **k = 헤드 전부**
(B1F 3,338)로 부른다. A 의 통상 호출은 k=30 이다. k 에 비례하거나 그보다 나쁘게
드는 자리가 있으면 여기서 보인다.

    python scripts/_probe_network_profile.py [도면.dxf] [--k N] [--top 25]
"""
from __future__ import annotations

import argparse
import cProfile
import io
import pstats
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--k", type=int, default=0, help="0 이면 헤드 전부(현행)")
    ap.add_argument("--top", type=int, default=25)
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    from routes.module_f.auto import (detect_head_candidates, head_region_of,
                                      parse_plan, region_around)
    from remote30_prototype import select_worst30_heads_anchored

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    ents, cat, _ = parse_plan(dxf)
    cand = detect_head_candidates(ents, cat)
    hp = [(h["x"], h["y"]) for h in cand]
    cx = sum(q[0] for q in hp) / len(hp)
    cy = sum(q[1] for q in hp) / len(hp)
    al = min(hp, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
    rects = region_around(cand, al)
    region = head_region_of(rects)
    k = a.k or max(1, len(cand))
    print(f"{dxf.name} · 헤드 후보 {len(cand):,} · k={k:,}\n")

    audit: dict = {}
    pr = cProfile.Profile()
    t0 = time.perf_counter()
    pr.enable()
    sel = select_worst30_heads_anchored(
        ents, cat, (float(al[0]), float(al[1])), region, k=k,
        audit_out=audit, load_mode=False)
    pr.disable()
    el = (time.perf_counter() - t0) * 1000
    print(f"■ 전체 {el:,.0f} ms · 뽑힌 헤드 "
          f"{len(getattr(sel, 'heads', None) or ()):,} · 간선 "
          f"{len(getattr(sel, 'edges', None) or ()):,}\n")

    s = io.StringIO()
    st = pstats.Stats(pr, stream=s).sort_stats("cumulative")
    st.print_stats(a.top)
    body = s.getvalue()
    # 헤더의 잡음을 걷고 표만 보여 준다.
    lines = body.splitlines()
    keep = False
    for ln in lines:
        if "ncalls" in ln and "cumtime" in ln:
            keep = True
        if keep and ln.strip():
            print("  " + ln.rstrip()[:150])

    print("\n  tottime 이 큰 자리가 실제로 CPU 를 태우는 곳이다.")
    print("  ncalls 가 k 에 비례해 늘면 «헤드마다 다시 하는» 자리다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
