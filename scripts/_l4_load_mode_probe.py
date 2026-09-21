"""L4 실측 — load_mode off/on 대명동 앵커 실행 비교.

usage:  python scripts/_l4_load_mode_probe.py
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


def run(load_mode: bool):
    bundle = rp.parse_dxf_bundle(ROOT / "samples/dxf/대명동201동 단위세대_layer정리.dxf")
    layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = rp.filter_pipenet_only(bundle)
    region = rp.HeadRegion.from_polygon(WEST_UNIT_POLY)
    gated = rp.detect_heads(ents, layer_cat, region=region)
    cx = sum(h.pos[0] for h in gated) / len(gated)
    cy = sum(h.pos[1] for h in gated) / len(gated)
    audit: dict = {}
    sel = rp.select_worst30_heads_anchored(ents, layer_cat, alarm_xy=(cx, cy),
                                           head_region=region, audit_out=audit,
                                           load_mode=load_mode)
    return sel, audit


def final_load(sel):
    g: dict = defaultdict(set)
    el: dict = {}
    for a, b, L in sel.edges:
        g[a].add(b)
        g[b].add(a)
        el[(min(a, b), max(a, b))] = L
    pm, _d, _o = rp.rooted_traversal(sel.source_pos, sel.edges)
    parents = {n: p for n, p in pm.items() if p is not None}
    return rp.compute_edge_load(g, el, sel.source_pos, [h.pos for h in sel.heads],
                                parents=parents)


def main() -> None:
    for mode in (False, True):
        sel, audit = run(mode)
        load = final_load(sel)
        zero = [k for k in load if load[k] == 0]
        print(f"\n=== load_mode={mode} ===")
        print(f"헤드 승인/부착/미도달 : {audit['heads']['detected_in_region']}"
              f"/{audit['heads']['attached']}/{len(audit['heads']['unreachable'])}")
        print(f"선정 헤드 / merged pipe : {len(sel.heads)} / {len(sel.edges)}")
        print(f"최종 부하 0 간선 : {len(zero)} / {len(load)}")
        print(f"총 연장 mm : {sum(L for _a, _b, L in sel.edges):.1f}")
        if "pruned" in audit:
            p = audit["pruned"]
            print(f"prune: dead={p['dead_edge_count']} len={p['dead_len_mm']:.1f}"
                  f" cycle_cut={p['cycle_cut_count']}  residual={audit['residual_cycles']}")


if __name__ == "__main__":
    main()
