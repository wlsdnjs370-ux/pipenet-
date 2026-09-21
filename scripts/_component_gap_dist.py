"""실배관(epsilon-cluster + T분기)만 적용한 뒤 남은 조각 사이의 실제 간격 분포.

간격이 mm~수백mm 면 '부속 기호가 배관선을 끊어놓은 자리'(실접속)이고,
수 m 면 진짜 별개 배관이다. 어디까지가 실배관인지 이 분포로 판단한다.

usage:  python scripts/_component_gap_dist.py [dxf]
"""
from __future__ import annotations

import io
import math
import sys
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import remote30_prototype as rp  # noqa: E402

DEFAULT = ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf"
BUCKETS = [50, 100, 200, 300, 500, 1000, 2000, 5000, 1e18]


def main():
    path = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    bundle = rp.parse_dxf_bundle(path)
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    graph, edge_len = rp._build_graph(ents, node_index=rp._NodeIndex(),
                                      layer_categories=layer_cat)
    rp.collapse_parallel_ladders(graph, edge_len)
    splits = rp._split_tee_branches(graph, edge_len)
    comps = rp._connected_components(graph)
    print(f"=== {path.name} ===")
    print(f"T분기 복원 {splits}건 후 조각 {len(comps)}개 "
          f"(노드 {len(graph)}, 최대 조각 {max(len(c) for c in comps)}노드)")

    owner = {}
    for i, c in enumerate(comps):
        for n in c:
            owner[n] = i
    # 각 조각에서 가장 가까운 다른 조각까지의 거리 (끝점 기준)
    dangling = [n for n, nb in graph.items() if len(nb) <= 2]
    cell = 2000.0
    grid: dict = {}
    for n in dangling:
        grid.setdefault((int(n[0] // cell), int(n[1] // cell)), []).append(n)
    best: dict = {}
    for u in dangling:
        cu = owner[u]
        gx, gy = int(u[0] // cell), int(u[1] // cell)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for v in grid.get((gx + dx, gy + dy), ()):
                    if owner[v] == cu:
                        continue
                    d = math.hypot(u[0] - v[0], u[1] - v[1])
                    pair = (min(cu, owner[v]), max(cu, owner[v]))
                    if d < best.get(pair, (1e18,))[0]:
                        best[pair] = (d, u, v)

    per_comp: dict = {}
    for (a, b), (d, u, v) in best.items():
        for c in (a, b):
            if d < per_comp.get(c, (1e18,))[0]:
                per_comp[c] = (d, u, v)

    gaps = sorted(d for d, _u, _v in per_comp.values())
    print(f"\n조각별 '가장 가까운 다른 조각까지' 거리 — {len(gaps)}개 조각 "
          f"(중앙값 {gaps[len(gaps)//2]:.0f}mm)" if gaps else "인접 조각 없음")
    lo = 0.0
    cum = 0
    for bk in BUCKETS:
        c = sum(1 for x in gaps if lo <= x < bk)
        cum += c
        if c:
            hi = "∞" if bk > 1e17 else f"{bk:.0f}"
            print(f"   {lo:>6.0f}~{hi:>6}mm : {c:4d}   (누적 {cum}/{len(gaps)})")
        lo = bk

    print("\n가장 가까운 조각쌍 상위 12건 (좌표 = 끊긴 자리):")
    for d, u, v in sorted(per_comp.values())[:12]:
        print(f"   {d:8.1f}mm  ({u[0]:.0f},{u[1]:.0f}) ~ ({v[0]:.0f},{v[1]:.0f})")


if __name__ == "__main__":
    main()
