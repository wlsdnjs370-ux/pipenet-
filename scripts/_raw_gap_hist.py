"""클러스터링 이전(eps≈0) 원시 끝점 간격 히스토그램 — 이음매 간격 자동판정 근거.

usage:  python scripts/_raw_gap_hist.py
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
from core import remote30_graph as rg  # noqa: E402

PATHS = [
    ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf",
    ROOT / "samples/dxf/LH306동_평면도.dxf",
    ROOT / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf",
    ROOT / "samples/dxf/LH 지하층배관도.dxf",
]
EDGES = [1, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100, 125, 150, 200, 250, 300, 400, 600, 1000]


def gaps_of(path: Path):
    bundle = rp.parse_dxf_bundle(path)
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    graph, _el = rp._build_graph(ents, node_index=rg._NodeIndex(epsilon_mm=1.0),
                                 layer_categories=layer_cat)
    nodes = list(graph)
    cell = 600.0
    grid: dict = {}
    for n in nodes:
        grid.setdefault((int(n[0] // cell), int(n[1] // cell)), []).append(n)
    out = []
    for u in nodes:
        if len(graph[u]) != 1:
            continue
        nb = graph[u]
        gx, gy = int(u[0] // cell), int(u[1] // cell)
        best = None
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for v in grid.get((gx + dx, gy + dy), ()):
                    if v is u or v in nb:
                        continue
                    d = math.hypot(u[0] - v[0], u[1] - v[1])
                    if best is None or d < best:
                        best = d
        if best is not None:
            out.append(best)
    return sorted(out)


def main():
    for p in PATHS:
        if not p.exists():
            print(f"\n=== {p.name} — 없음 ===")
            continue
        g = gaps_of(p)
        print(f"\n=== {p.name} — 느슨한 끝점 {len(g)}개 ===")
        if not g:
            continue
        lo = 0.0
        cum = 0
        for e in EDGES:
            c = sum(1 for x in g if lo <= x < e)
            cum += c
            if c:
                print(f"   {lo:>6.0f}~{e:>5}mm : {c:5d}  누적 {100*cum/len(g):5.1f}%")
            lo = e
        rest = len(g) - cum
        if rest:
            print(f"   {lo:>6.0f}~   ∞mm : {rest:5d}")
        for q in (50, 75, 90, 95):
            print(f"   p{q} = {g[min(len(g)-1, q*len(g)//100)]:.1f}mm")


if __name__ == "__main__":
    main()
