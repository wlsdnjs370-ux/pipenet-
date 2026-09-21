# -*- coding: utf-8 -*-
"""B1F 밸브->헤드 경로 꼬임 진단 (수정 전/후 비교용).

select_worst30_heads 를 auto 로 돌려 top-K 헤드의
  - path_dist (그래프 최단경로 거리)
  - euclid   (밸브-헤드 직선거리)
  - ratio    (path/euclid ; 1 에 가까울수록 곧게, 클수록 뺑 돌아감)
를 뽑는다. tangling 지표 = 평균/최대 ratio + 총 path 길이.
ASCII-only 출력 (cp949 콘솔 안전).
"""
from __future__ import annotations

import math
import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))

from remote30_prototype import (  # noqa: E402
    parse_dxf_bundle, filter_pipenet_only, select_worst30_heads,
)

DXF = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"


def main():
    assert DXF.is_file(), f"DXF not found: {DXF}"
    print(f"file: {DXF.name} ({DXF.stat().st_size/1e6:.1f} MB)")
    t = time.time()
    bundle = parse_dxf_bundle(DXF)
    layer_categories = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = filter_pipenet_only(bundle)
    print(f"parse+filter: {time.time()-t:.1f}s  pipe_ents={len(pipe_ents):,}")

    MANUAL_SRC = (661506.0, 177357.0)  # 사용자 지정 알람밸브
    t = time.time()
    sel = select_worst30_heads(pipe_ents, layer_categories, k=30,
                               manual_source=MANUAL_SRC)
    print(f"select_worst30 (manual src={MANUAL_SRC}): {time.time()-t:.1f}s")
    if sel.source_pos is None:
        print("NO SOURCE -> abort")
        return
    sx, sy = sel.source_pos
    print(f"source_kind={sel.source_kind} pos=({sx:.0f},{sy:.0f}) "
          f"bridge_dist={sel.source_bridge_dist_mm:.0f} fallback={sel.source_fallback}")
    print(f"heads_selected={len(sel.heads)} subgraph_edges={len(sel.edges)}")

    rows = []
    for h, pd in zip(sel.heads, sel.distances):
        eu = math.hypot(h.pos[0] - sx, h.pos[1] - sy)
        ratio = pd / eu if eu > 1e-6 else float("inf")
        rows.append((ratio, pd, eu, h.pos))
    rows.sort(reverse=True)  # worst tangle first

    ratios = [r for r, _, _, _ in rows if math.isfinite(r)]
    total_path = sum(pd for _, pd, _, _ in rows)
    print(f"\n-- tangling (path/euclid) --")
    print(f"avg_ratio={sum(ratios)/len(ratios):.2f}  max_ratio={max(ratios):.2f}  "
          f"total_path_mm={total_path:,.0f}")
    print(f"worst-head path_dist={sel.distances[0]:,.0f}mm "
          f"(sel.distances[0] = farthest)")
    print("\ntop 8 most-tangled heads:")
    print(f"{'ratio':>6} {'path_mm':>10} {'euclid_mm':>10}  head_pos")
    for r, pd, eu, pos in rows[:8]:
        print(f"{r:6.2f} {pd:10.0f} {eu:10.0f}  ({pos[0]:.0f},{pos[1]:.0f})")
    print("DONE")


if __name__ == "__main__":
    main()
