# -*- coding: utf-8 -*-
"""분기영역(branch_zones) 알고리즘 검증 — B1F 좌측 wing.

목표 확인:
  1) branch_zones 미지정: 영역까지 주배관이 여러 갈래(트리 fan-out)인가?
  2) branch_zones 지정: 영역 밖이 source->영역 단일 corridor 로 붕괴하는가?
     (영역 밖 노드 중 degree>=3 junction 이 사라져야 함)
  3) 영역 안 노드/edge 는 보존되는가?
ASCII-only stdout.
"""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
import remote30_prototype as R  # noqa: E402

ORIG = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"


def in_rect(pt, rect):
    x1, y1, x2, y2 = rect
    lo_x, hi_x = sorted((x1, x2)); lo_y, hi_y = sorted((y1, y2))
    return lo_x <= pt[0] <= hi_x and lo_y <= pt[1] <= hi_y


def analyze(res, region_rect, tag):
    src = res.source_pos
    # 그래프 재구성 (edges = (a,b,L) 3-tuple)
    from collections import defaultdict
    adj = defaultdict(set)
    for a, b, L in res.edges:
        adj[a].add(b); adj[b].add(a)
    # source 로부터 영역까지, 영역 밖 junction(degree>=3) 개수
    out_junc = 0
    in_nodes = 0
    for n in adj:
        deg = len(adj[n])
        inside = in_rect(n, region_rect) if region_rect else False
        if inside:
            in_nodes += 1
        elif deg >= 3:
            out_junc += 1
    heads_in = sum(1 for h in res.heads if in_rect(h.pos, region_rect)) if region_rect else 0
    print(f"[{tag}] heads={len(res.heads)} nodes={len(adj)} edges={len(res.edges)} "
          f"src={tuple(round(v) for v in src) if src else None}", flush=True)
    print(f"   영역밖 junction(deg>=3)={out_junc}  영역안 노드={in_nodes}  영역안 헤드={heads_in}",
          flush=True)


def main():
    bundle = R.parse_dxf_bundle_cached(ORIG)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pe = R.filter_pipenet_only(bundle)

    # 좌측 wing PIT (summary 기준). 우선 head 분포로 영역 잡기 위해 no-zone 실행.
    SRC = (156776.0, 177970.0)
    res0 = R.select_worst30_heads(pe, lc, k=115, manual_source=SRC)
    hxs = [h.pos[0] for h in res0.heads]; hys = [h.pos[1] for h in res0.heads]
    if hxs:
        hbb = (min(hxs), min(hys), max(hxs), max(hys))
        print(f"no-zone head bbox = ({hbb[0]:.0f},{hbb[1]:.0f},{hbb[2]:.0f},{hbb[3]:.0f})", flush=True)
    else:
        print("no heads from left-wing source — source may be wrong", flush=True)
        return

    # 분기영역 = head cluster 를 감싸되 source 는 제외하도록 약간 여유
    rect = (hbb[0] - 2000, hbb[1] - 2000, hbb[2] + 2000, hbb[3] + 2000)
    print(f"branch_zone rect = {tuple(round(v) for v in rect)}\n", flush=True)

    analyze(res0, rect, "WITHOUT branch_zone")
    res1 = R.select_worst30_heads(pe, lc, k=115, manual_source=SRC, branch_zones=[rect])
    analyze(res1, rect, "WITH    branch_zone")
    print("\nDONE", flush=True)


if __name__ == "__main__":
    main()
