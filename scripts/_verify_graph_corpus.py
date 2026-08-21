# -*- coding: utf-8 -*-
"""배관망 그래프 코퍼스 회귀 — 이음매·노이즈 컷을 건드릴 때 전후를 비교한다.

`auto_snap_eps` 와 그래프 노이즈 컷은 **모든 도면**의 결과를 바꾼다. 기준 도면
(대명동)이 흔들리지 않는지, 타현장 도면이 실제로 좋아지는지를 같은 표로 본다.

지표
  · 이음매(mm)      — auto_snap_eps 가 고른 값
  · 최대조각        — 가장 큰 연결성분의 노드 수 (클수록 배관망이 하나로 붙음)
  · 성분수          — 연결조각 개수 (작을수록 좋음)
  · 헤드도달        — 가장 큰 조각에 붙은 헤드 수 / 전체 헤드
  · 연장(m)         — 그래프 총 배관 연장 (과대 병합 감시용)

사용:  python scripts/_verify_graph_corpus.py [저장이름]
       인자를 주면 결과를 data/_graph_corpus_<이름>.json 으로 저장하고,
       data/_graph_corpus_before.json 이 있으면 그것과 나란히 보여준다.
"""
from __future__ import annotations

import json
import math
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "core"))

CASES = [
    ("대명동 평면도(기준)", "routes/제출용[최종]/1. 입력도면 대명동 단위세대 평면도.dxf"),
    ("대명동 기계실(기준)", "routes/제출용[최종]/1. 입력도면 대명동 단위세대 기계실.dxf"),
    ("S1 대구오페라", "data/_genz_dxf/S1_대구오페라_지하층소화설비.dxf"),
    ("S2 죽전", "data/_genz_dxf/S2_죽전_지하주차장소화설비.dxf"),
    ("S3 청라포레스트", "data/_genz_dxf/S3_청라포레스트_지하주차장소화설비.dxf"),
    ("S4 대우이안", "data/_genz_dxf/S4_대우이안_지하4층소방시설.dxf"),
    ("S6 대구오페라 세대", "data/_genz_dxf/S6_대구오페라_단위세대.dxf"),
    ("MF-304 청라스타필드", "data/_genz_dxf/MF-304(지상1층 소방시설(기계) 평면도).dxf"),
]


def measure(path: Path) -> dict:
    from remote30_prototype import (
        _NodeIndex, _build_graph, _connected_components, auto_snap_eps,
        collapse_parallel_ladders, detect_heads, filter_pipenet_only,
        parse_dxf_bundle, _split_tee_branches)

    bundle = parse_dxf_bundle(path)
    cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    ents = filter_pipenet_only(bundle)
    heads = detect_heads(ents, cats)
    eps = auto_snap_eps(ents, cats)
    graph, edge_len = _build_graph(
        ents, node_index=_NodeIndex(epsilon_mm=eps), layer_categories=cats)
    collapse_parallel_ladders(graph, edge_len)
    _split_tee_branches(graph, edge_len)
    comps = _connected_components(graph)
    comps = sorted(comps, key=len, reverse=True)
    biggest = set(comps[0]) if comps else set()

    # 가장 큰 조각에 실제로 닿는 헤드 수 — 배관망이 쓸 만한지의 최종 지표.
    cell = 2000.0
    grid: dict[tuple[int, int], list] = {}
    for n in biggest:
        grid.setdefault((int(n[0] // cell), int(n[1] // cell)), []).append(n)
    reach = 0
    for h in heads:
        hx, hy = h.pos
        gx, gy = int(hx // cell), int(hy // cell)
        ok = False
        for ax in (gx - 1, gx, gx + 1):
            for ay in (gy - 1, gy, gy + 1):
                for n in grid.get((ax, ay), ()):
                    if math.hypot(hx - n[0], hy - n[1]) <= 1500.0:
                        ok = True
                        break
                if ok:
                    break
            if ok:
                break
        reach += 1 if ok else 0

    return {
        "eps": round(eps, 0),
        "nodes": len(graph),
        "largest": len(biggest),
        "comps": len(comps),
        "heads": len(heads),
        "reach": reach,
        "len_m": round(sum(edge_len.values()) / 1000.0, 1),
    }


def main():
    tag = sys.argv[1] if len(sys.argv) > 1 else None
    before = {}
    bpath = ROOT / "data" / "_graph_corpus_before.json"
    if tag and tag != "before" and bpath.is_file():
        before = json.loads(bpath.read_text(encoding="utf-8"))

    out = {}
    print(f"{'도면':22s} {'이음매':>6s} {'최대조각':>8s} {'성분수':>7s} "
          f"{'헤드':>6s} {'도달':>6s} {'연장(m)':>9s}")
    print("-" * 78)
    for name, rel in CASES:
        p = ROOT / rel
        if not p.exists():
            print(f"{name:22s}  파일 없음 — 건너뜀")
            continue
        t0 = time.perf_counter()
        r = measure(p)
        out[name] = r
        line = (f"{name:22s} {r['eps']:6.0f} {r['largest']:8,d} {r['comps']:7,d} "
                f"{r['heads']:6,d} {r['reach']:6,d} {r['len_m']:9,.1f}")
        b = before.get(name)
        if b:
            line += (f"   ← 전: 이음매 {b['eps']:.0f} · 최대 {b['largest']:,} · "
                     f"성분 {b['comps']:,} · 도달 {b['reach']:,} · "
                     f"연장 {b['len_m']:,.1f}")
        print(line + f"   [{time.perf_counter()-t0:.0f}s]")

    if tag:
        dst = ROOT / "data" / f"_graph_corpus_{tag}.json"
        dst.write_text(json.dumps(out, ensure_ascii=False, indent=1),
                       encoding="utf-8")
        print(f"\n저장: {dst}")


if __name__ == "__main__":
    main()
