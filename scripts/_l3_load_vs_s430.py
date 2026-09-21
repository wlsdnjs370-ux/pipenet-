"""L3 실측 — S430 담당 헤드 수(내부 _subtree_calc)와 L1 compute_edge_load 대조.

usage:  python scripts/_l3_load_vs_s430.py
"""
from __future__ import annotations

import io
import sys
from collections import defaultdict
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import remote30_prototype as rp  # noqa: E402

WEST_UNIT_POLY = [
    (244500.0, -243500.0), (253500.0, -243500.0),
    (253500.0, -221500.0), (244500.0, -221500.0),
]


def _nfpc(n: int) -> int:
    for lim, d in ((2, 25), (3, 32), (5, 40), (10, 50), (30, 65),
                   (60, 80), (80, 90), (100, 100), (160, 125)):
        if n <= lim:
            return d
    return 150


def main() -> None:
    bundle = rp.parse_dxf_bundle(ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf")
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    region = rp.HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(ents, layer_cat, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    audit: dict = {}
    sel = rp.select_worst30_heads_anchored(ents, layer_cat, alarm_xy=(cx, cy),
                                           head_region=region, audit_out=audit)

    print(f"region 승인 헤드 : {audit['heads']['detected_in_region']}")
    print(f"  부착됨         : {audit['heads']['attached']}")
    print(f"  미도달         : {len(audit['heads']['unreachable'])}")
    print(f"선정 헤드(top-K) : {len(sel.heads)}")
    print(f"최종 merged pipe : {len(sel.edges)}")

    # ── S430 현행: rooted_traversal 트리 + 재귀 _subtree_calc, 선정 헤드만 집계
    parent_map, _depth, bfs_order = rp.rooted_traversal(sel.source_pos, sel.edges)
    children_of: dict = defaultdict(list)
    for nd, pr in parent_map.items():
        if pr is not None:
            children_of[pr].append(nd)
    sel_head_set = {h.pos for h in sel.heads}
    subtree: dict = {}

    def _calc(n):
        c = 1 if n in sel_head_set else 0
        for ch in children_of[n]:
            c += _calc(ch)
        subtree[n] = c
        return c

    _calc(sel.source_pos)

    def _cur(a, b) -> int:
        if parent_map.get(b) == a:
            return subtree.get(b, 0)
        if parent_map.get(a) == b:
            return subtree.get(a, 0)
        return 0

    # ── L1: 같은 트리(parents)·같은 헤드 집합으로 compute_edge_load
    g: dict = defaultdict(set)
    el: dict = {}
    for a, b, L in sel.edges:
        g[a].add(b)
        g[b].add(a)
        el[(min(a, b), max(a, b))] = L
    parents = {n: p for n, p in parent_map.items() if p is not None}
    load = rp.compute_edge_load(g, el, sel.source_pos, list(sel_head_set),
                                parents=parents)

    diff = [(a, b, _cur(a, b), load.get((min(a, b), max(a, b)), 0))
            for a, b, _L in sel.edges
            if _cur(a, b) != load.get((min(a, b), max(a, b)), 0)]
    print(f"\n[A] 같은 트리·같은 헤드집합 대조 — 불일치 pipe : {len(diff)} / {len(sel.edges)}")
    for a, b, c1, c2 in diff[:10]:
        print(f"    {a} → {b}  현행 {c1} vs L1 {c2}")
    dia_diff = sum(1 for a, b, _L in sel.edges
                   if _nfpc(_cur(a, b)) != _nfpc(load.get((min(a, b), max(a, b)), 0)))
    print(f"    NFPC 별표1 최소 호칭경 불일치 구간 : {dia_diff}")

    # ── [B] 헤드 집합 차이의 영향: 승인 헤드 전체로 세면?
    all_pos = [h.pos for h in gated]
    load_all = rp.compute_edge_load(g, el, sel.source_pos, all_pos, parents=parents)
    b_diff = [(a, b, load.get((min(a, b), max(a, b)), 0),
               load_all.get((min(a, b), max(a, b)), 0))
              for a, b, _L in sel.edges
              if load.get((min(a, b), max(a, b)), 0)
              != load_all.get((min(a, b), max(a, b)), 0)]
    b_dia = sum(1 for a, b, _L in sel.edges
                if _nfpc(load.get((min(a, b), max(a, b)), 0))
                != _nfpc(load_all.get((min(a, b), max(a, b)), 0)))
    print(f"\n[B] 선정 헤드(top-K) vs 승인 헤드 전체 — 부하 다른 pipe : {len(b_diff)} / {len(sel.edges)}")
    print(f"    NFPC 최소 호칭경이 달라지는 구간 : {b_dia}")
    for a, b, c1, c2 in b_diff[:10]:
        print(f"    {a} → {b}  top-K {c1} → 전체 {c2}  (호칭경 {_nfpc(c1)}→{_nfpc(c2)})")


if __name__ == "__main__":
    main()
