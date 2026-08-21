# -*- coding: utf-8 -*-
"""MF-304 최불리 배관망 생성이 막히는 지점을 실제 파이프라인으로 짚는다.

«길을 못 찾아서» 인지 확인하려면 그래프가 실제로 어떻게 생겼는지 봐야 한다 —
급수원(알람밸브)을 찾았는지, 그래프가 몇 조각으로 쪼개졌는지, 헤드가 그중
어느 조각에 붙어 있는지. 추측 대신 그 숫자를 뽑는다.
"""
from __future__ import annotations

import collections
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "core"))

DXF = ROOT / "data" / "_genz_dxf" / "MF-304(지상1층 소방시설(기계) 평면도).dxf"


def main():
    from remote30_prototype import (
        _NodeIndex, _build_graph, _find_source, auto_snap_eps, detect_heads,
        filter_pipenet_only, parse_dxf_bundle, select_worst30_heads)

    print(f"■ {DXF.name}")
    t0 = time.perf_counter()
    bundle = parse_dxf_bundle(DXF)
    cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = filter_pipenet_only(bundle)
    print(f"  파싱 {time.perf_counter()-t0:.1f}s · 전체 entity {len(bundle.entities):,}"
          f" · 레이어 {len(cats)} · pipenet {len(ents):,}")
    print("  레이어 분류:", dict(collections.Counter(cats.values())))
    xd = bundle.xref_diagnostics or {}
    print(f"  XREF: {xd.get('xref_count', 0)}건 · 껍데기 시트 판정 "
          f"{xd.get('is_sheet')} · 내용 엔티티 {xd.get('content_entities')}")

    stats: dict = {}
    heads = detect_heads(ents, cats, stats=stats)
    print(f"  헤드 {len(heads)}개 (신호 {stats.get('raw_cues')} · "
          f"결합반경 {stats.get('cluster_r')}mm)")

    # ── 급수원(알람밸브) ────────────────────────────────────────
    src, kind = _find_source(ents, cats)
    print(f"  급수원: {src} · 방식={kind}")

    # ── 그래프 ──────────────────────────────────────────────────
    eps = auto_snap_eps(ents, cats)
    graph, edge_len = _build_graph(
        ents, node_index=_NodeIndex(epsilon_mm=eps), layer_categories=cats)
    print(f"  그래프: 노드 {len(graph):,} · 간선 "
          f"{sum(len(v) for v in graph.values())//2:,} · 이음매 {eps:.0f}mm")

    # 연결성분
    seen, comps = set(), []
    for n in graph:
        if n in seen:
            continue
        stack, comp = [n], []
        seen.add(n)
        while stack:
            u = stack.pop()
            comp.append(u)
            for v in graph.get(u, ()):
                if v not in seen:
                    seen.add(v)
                    stack.append(v)
        comps.append(comp)
    comps.sort(key=len, reverse=True)
    print(f"  연결성분 {len(comps):,}개 · 상위 크기 {[len(c) for c in comps[:8]]}")

    # 헤드가 어느 성분에 붙는가 — 가장 가까운 그래프 노드까지 거리
    import math
    nodes = list(graph)
    comp_of = {}
    for ci, c in enumerate(comps):
        for n in c:
            comp_of[n] = ci
    grid = {}
    CELL = 2000.0
    for n in nodes:
        grid.setdefault((int(n[0] // CELL), int(n[1] // CELL)), []).append(n)
    hit = collections.Counter()
    far = 0
    dists = []
    for h in heads:
        hx, hy = h.pos
        gx, gy = int(hx // CELL), int(hy // CELL)
        best, bn = None, None
        for ax in range(gx - 2, gx + 3):
            for ay in range(gy - 2, gy + 3):
                for n in grid.get((ax, ay), ()):
                    d = math.hypot(hx - n[0], hy - n[1])
                    if best is None or d < best:
                        best, bn = d, n
        if bn is None or best > 5000:
            far += 1
            continue
        dists.append(best)
        hit[comp_of[bn]] += 1
    dists.sort()
    print(f"  헤드→그래프 최근접: 중앙 {dists[len(dists)//2]:.0f}mm" if dists else "  헤드 부착 실패")
    print(f"  헤드가 붙은 성분 상위: {hit.most_common(8)} · 5m 밖 헤드 {far}")

    if src is not None:
        sgx, sgy = int(src[0] // CELL), int(src[1] // CELL)
        best, bn = None, None
        for ax in range(sgx - 3, sgx + 4):
            for ay in range(sgy - 3, sgy + 4):
                for n in grid.get((ax, ay), ()):
                    d = math.hypot(src[0] - n[0], src[1] - n[1])
                    if best is None or d < best:
                        best, bn = d, n
        print(f"  급수원→그래프 최근접 {best:.0f}mm · 성분 "
              f"{comp_of.get(bn) if bn else None} (크기 "
              f"{len(comps[comp_of[bn]]) if bn else 0})")

    # ── 실제 선정 실행 ──────────────────────────────────────────
    print("\n  ── select_worst30_heads 실행 ──")
    t0 = time.perf_counter()
    try:
        res = select_worst30_heads(ents, cats, k=30,
                                   progress_cb=lambda f, m: None)
        print(f"  {time.perf_counter()-t0:.1f}s · 선정 헤드 "
              f"{len(getattr(res, 'heads', []) or [])} · "
              f"경로 노드 {len(getattr(res, 'nodes', []) or [])} · "
              f"간선 {len(getattr(res, 'edges', []) or [])}")
        for f in ("source", "source_kind", "notes", "warnings", "reason"):
            if hasattr(res, f):
                print(f"    {f} = {getattr(res, f)!r}"[:220])
    except Exception as exc:  # noqa: BLE001
        import traceback
        print(f"  !! 실패 {time.perf_counter()-t0:.1f}s: "
              f"{type(exc).__name__}: {exc}")
        traceback.print_exc(limit=6)


if __name__ == "__main__":
    main()
