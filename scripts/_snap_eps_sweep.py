"""끝점 클러스터 허용치(SNAP_TOL_MM) 스윕 — 추정연결 없이 실배관만으로 얼마나 이어지나.

usage:  python scripts/_snap_eps_sweep.py [dxf ...]
"""
from __future__ import annotations

import io
import math
import sys
import time
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import remote30_prototype as rp  # noqa: E402
from core import remote30_graph as rg  # noqa: E402

DEFAULTS = [
    ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf",
    ROOT / "samples/dxf/LH306동_평면도.dxf",
    ROOT / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf",
    ROOT / "samples/dxf/LH 지하층배관도.dxf",
]
EPS = [50, 60, 75, 90, 110, 130, 160, 200, 250, 300]


def run(ents, layer_cat, eps: float, tee_gap: float):
    t0 = time.perf_counter()
    ni = rg._NodeIndex(epsilon_mm=eps)
    graph, edge_len = rp._build_graph(ents, node_index=ni,
                                      layer_categories=layer_cat)
    rp.collapse_parallel_ladders(graph, edge_len)
    splits = rp._split_tee_branches(graph, edge_len, max_gap_mm=tee_gap)
    heads = rp.detect_heads(ents, layer_cat)
    drops = 0
    for h in heads:
        hp = ni.canonical(h.pos[0], h.pos[1])
        if hp in graph:
            continue
        nearest = rp._nearest_graph_node(graph, hp)
        if nearest is None or hp == nearest:
            continue
        d = math.hypot(hp[0] - nearest[0], hp[1] - nearest[1])
        if d > 1e-3 and d <= rp.HEAD_BRIDGE_MAX_MM:
            graph.setdefault(hp, set()).add(nearest)
            graph[nearest].add(hp)
            edge_len[(min(hp, nearest), max(hp, nearest))] = d
            drops += 1
    src, _k = rp._find_source(ents, layer_cat)
    if src is not None:
        src = ni.canonical(src[0], src[1])
        if src not in graph:
            src = rp._nearest_graph_node(graph, src)
    seen = set()
    if src in graph:
        seen, stack = {src}, [src]
        while stack:
            u = stack.pop()
            for v in graph[u]:
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
    reached = sum(1 for h in heads
                  if ni.canonical(h.pos[0], h.pos[1]) in seen)
    comps = rp._connected_components(graph)
    return {
        "조각": len(comps),
        "최대조각": max((len(c) for c in comps), default=0),
        "노드": len(graph),
        "간선": len(edge_len),
        "T분기": splits,
        "헤드drop": drops,
        "헤드도달": f"{reached}/{len(heads)}",
        "총연장m": round(sum(edge_len.values()) / 1000, 1),
        "초": round(time.perf_counter() - t0, 2),
    }


def main():
    paths = [Path(a) for a in sys.argv[1:]] or DEFAULTS
    for p in paths:
        if not p.exists():
            print(f"\n=== {p.name} — 없음 ===")
            continue
        bundle = rp.parse_dxf_bundle(p)
        layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
        ents = rp.filter_pipenet_only(bundle)
        print(f"\n=== {p.name} — 용접·브리지 전부 OFF ===")
        rows = []
        for e in EPS:
            rows.append((f"eps {e}mm", run(ents, layer_cat, e, min(20.0, e * 0.4))))
        cols = list(rows[0][1])
        w0 = max(len(n) for n, _ in rows)
        print(f"{'':{w0}}  " + "  ".join(f"{c:>9}" for c in cols))
        for name, r in rows:
            print(f"{name:{w0}}  " + "  ".join(f"{str(r[c]):>9}" for c in cols))
        t0 = time.perf_counter()
        aud: dict = {}
        chosen = rp.auto_snap_eps(ents, layer_cat, audit_out=aud)
        print(f"  → auto_snap_eps = {chosen:.0f}mm  ({time.perf_counter()-t0:.2f}초, "
              f"후보중 기각 {sum(1 for t in aud['snap_eps']['trials'] if not t['kept'])}건)")


if __name__ == "__main__":
    main()
