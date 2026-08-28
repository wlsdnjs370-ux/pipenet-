# -*- coding: utf-8 -*-
"""최단경로 나무 재사용이 «같은 답» 을 내는가 — 빠른 것만으론 부족하다.

`_finalize_selection` 은 거리 맵을 한 번 만든 뒤 헤드마다 `_shortest_path` 로
Dijkstra 를 또 돌렸다(B1F 실측 2,206회 · 40.5초). 이제는 그 한 번에서 나무를
같이 받아 되돌아 걷는다.

둘 다 같은 그래프·같은 가중치·같은 힙이라 나무가 같아야 하고, 따라서 경로도
같아야 한다. «해야 한다» 로 끝내지 않고 잰다 — 무승부(같은 길이의 다른 경로)가
갈리면 합집합 간선이 달라질 수 있다.

옛 방식을 `_shortest_path` 로 그대로 재현해 **경로 하나하나를 대조**한다.

    python scripts/_verify_dijkstra_tree.py [도면.dxf ...]
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf",
    ROOT / "samples" / "dxf" / "계통도_LH_306.dxf",
    ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf",
]
GRAB: dict = {}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A
    from remote30_graph import _dijkstra_from, _path_from_prev, _shortest_path
    from routes.module_f.auto import (detect_head_candidates, head_region_of,
                                      parse_plan, region_around)

    # 나무를 만든 그 자리를 붙잡는다 — 실제로 쓰인 graph/edge_len/src 로 대조해야
    # 의미가 있다(합성 그래프로 재면 «이 도면에서» 같은지를 못 잰다).
    real = _dijkstra_from

    def hook(graph, edge_len, src, prev_out=None):
        out = real(graph, edge_len, src, prev_out)
        if prev_out is not None:
            GRAB["g"], GRAB["el"], GRAB["src"] = graph, edge_len, src
            GRAB["prev"] = dict(prev_out)
        return out

    A._dijkstra_from = hook
    try:
        for dxf in [Path(x) for x in sys.argv[1:]] or DEFAULT:
            if not dxf.is_file():
                print(f"■ {dxf.name} — 파일 없음\n")
                continue
            GRAB.clear()
            ents, cat, _ = parse_plan(dxf)
            cand = detect_head_candidates(ents, cat)
            if not cand:
                print(f"■ {dxf.name} — 헤드 0\n")
                continue
            hp = [(h["x"], h["y"]) for h in cand]
            cx = sum(q[0] for q in hp) / len(hp)
            cy = sum(q[1] for q in hp) / len(hp)
            al = min(hp, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
            region = head_region_of(region_around(cand, al))

            t0 = time.perf_counter()
            sel = A.select_worst30_heads_anchored(
                ents, cat, (float(al[0]), float(al[1])), region,
                k=max(1, len(cand)), load_mode=False)
            el = (time.perf_counter() - t0) * 1000

            heads = list(getattr(sel, "heads", None) or ())
            edges = list(getattr(sel, "edges", None) or ())
            print(f"■ {dxf.name}")
            print(f"    {el:>9,.0f} ms · 헤드 {len(heads):,} · 간선 {len(edges):,}")

            g, elen = GRAB.get("g"), GRAB.get("el")
            src, prev = GRAB.get("src"), GRAB.get("prev")
            if not g:
                print("    (나무를 못 잡았다 — 대조 생략)\n")
                continue

            # ★옛 방식과 경로를 하나하나 대조. 헤드 노드는 나무에 든 것 전부에서
            #   고르되, 큰 도면에선 앞에서 400개만 본다(한 개라도 다르면 실패).
            tgts = [n for n in prev][:400]
            diff = miss = 0
            t1 = time.perf_counter()
            for t in tgts:
                new = _path_from_prev(prev, src, t)
                old = _shortest_path(g, elen, src, t)
                if not old and not new:
                    continue
                if not old or not new:
                    miss += 1
                elif old != new:
                    diff += 1
            t_cmp = (time.perf_counter() - t1) * 1000
            ok = (diff == 0 and miss == 0)
            print(f"    대조 {len(tgts):,}경로 · 다름 {diff} · 한쪽만 {miss}"
                  f"   {'[OK]' if ok else '[FAIL]'}   ({t_cmp:,.0f} ms)")
            if not ok:
                print("    ★경로가 갈린다 — 무승부 처리가 다르다는 뜻이다.")
            print()
    finally:
        A._dijkstra_from = real

    print("  «다름 0 · 한쪽만 0» 이어야 최적화다. 하나라도 갈리면 답이 바뀐 것이다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
