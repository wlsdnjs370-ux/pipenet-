# -*- coding: utf-8 -*-
"""B1F 원본 vs '_도면정리_최소'(정답 배관망) 비교 — 추출 수렴 진단.

사용자가 원본에서 배관망을 최소(가장 불리한 헤드망)로 정리한 참조 파일을 만들었다.
목표: 원본 자동추출이 참조와 같아지도록. 먼저 무엇이 다른지 정량화한다.

출력:
  - 레이어별 엔티티 수 (원본 vs 참조) — 어느 레이어가 줄었나
  - filter_pipenet_only 통과 엔티티 수
  - _build_graph 노드/엣지/컴포넌트
  - 두 파일 PIPE 계열 라인 좌표 bbox
ASCII-only 로그.
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
import remote30_prototype as R  # noqa: E402

ORIG = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"
REF = BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"


def layer_entity_counts(bundle):
    """레이어 -> (엔티티종류 Counter). PIPE 계열 위주."""
    per_layer = defaultdict(Counter)
    for e in bundle.entities:
        ly = e.get("l", "?")
        per_layer[ly][e.get("t", "?")] += 1
    return per_layer


def analyze(tag, dxf):
    print(f"\n===== {tag}: {dxf.name} =====", flush=True)
    if not dxf.exists():
        print("  MISSING", flush=True)
        return None
    bundle = R.parse_dxf_bundle_cached(dxf)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    print(f"  total_entities={len(bundle.entities)}  layers={len(bundle.layers)}", flush=True)
    per = layer_entity_counts(bundle)
    # 카테고리별 요약
    cat_ent = Counter()
    for ly, cnt in per.items():
        cat_ent[lc.get(ly, "?")] += sum(cnt.values())
    print("  category entity totals:", dict(cat_ent.most_common()), flush=True)
    # PIPE 계열 레이어 상세
    print("  -- per-layer (category=PIPE or name~pipe/fire/소화/배관) --", flush=True)
    rows = []
    for ly, cnt in per.items():
        cat = lc.get(ly, "?")
        nm = ly.lower()
        if cat == "PIPE" or any(k in nm for k in ("pipe", "fire", "소화", "배관", "sp", "스프")):
            rows.append((ly, cat, sum(cnt.values()), dict(cnt)))
    rows.sort(key=lambda r: -r[2])
    for ly, cat, tot, cnt in rows[:40]:
        print(f"     [{cat:6s}] {ly!r:40s} n={tot:6d} {cnt}", flush=True)
    # 추출
    pe = R.filter_pipenet_only(bundle)
    print(f"  filter_pipenet_only -> {len(pe)} entities", flush=True)
    graph, edge_len = R._build_graph(pe, layer_categories=lc)
    from core.remote30_graph import _connected_components
    print(f"  _build_graph -> nodes={len(graph)} edges={len(edge_len)} comps={len(_connected_components(graph))}", flush=True)
    return {"per": per, "lc": lc, "graph": graph, "edge_len": edge_len, "pe": pe}


def main():
    o = analyze("ORIG", ORIG)
    r = analyze("REF ", REF)
    if o and r:
        print("\n===== DIFF (layer entity totals) =====", flush=True)
        lo = {ly: sum(c.values()) for ly, c in o["per"].items()}
        lr = {ly: sum(c.values()) for ly, c in r["per"].items()}
        allk = set(lo) | set(lr)
        diffs = []
        for k in allk:
            a, b = lo.get(k, 0), lr.get(k, 0)
            if a != b:
                diffs.append((b - a, k, a, b))
        diffs.sort()
        for d, k, a, b in diffs:
            print(f"   {d:+8d}  {k!r:40s} orig={a} ref={b}", flush=True)
    print("\nDONE", flush=True)


if __name__ == "__main__":
    main()
