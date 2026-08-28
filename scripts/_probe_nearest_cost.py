# -*- coding: utf-8 -*-
"""격자 색인이 «작은 도면» 에서는 손해인가 — 90초의 범인을 가린다.

LH306 자동 추출이 90.2초다. 직전 브라우저 통과 뒤에 A 를 한 번 더 건드렸으므로
(`_NearestNodeIndex`), 내가 느리게 만든 것인지부터 가른다.

색인은 부를 때마다 O(N) 으로 «세운다». 큰 도면에선 그 값이 헤드 × 노드를
지우고도 남지만, 헤드가 몇 개 안 되는 작은 도면에선 세우는 값만 남을 수 있다.
그래서 색인을 쓰는 경로와 예전 선형 탐색 경로를 나란히 재고, 색인을 «몇 번»
세우는지도 센다.

    python scripts/_probe_nearest_cost.py [도면.dxf ...]
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "samples" / "dxf" / "LH306동_평면도.dxf",
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf",
]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A
    import remote30_graph as G
    from routes.module_f.auto import (detect_head_candidates, head_region_of,
                                      parse_plan, region_around)

    real_cls = G._NearestNodeIndex

    class Counting(real_cls):
        n_built = 0
        n_query = 0
        build_ms = 0.0

        def __init__(self, graph):
            t0 = time.perf_counter()
            super().__init__(graph)
            Counting.build_ms += (time.perf_counter() - t0) * 1000
            Counting.n_built += 1
            Counting.n_nodes = len(graph)

        def nearest(self, pt):
            Counting.n_query += 1
            return super().nearest(pt)

    # 예전 경로 재현 — 색인을 세우지 않고 매번 선형으로 훑는다.
    class Linear:
        __slots__ = ("_g",)

        def __init__(self, graph):
            self._g = graph

        def nearest(self, pt):
            return G._nearest_graph_node(self._g, pt)

    for dxf in [Path(x) for x in sys.argv[1:]] or DEFAULT:
        if not dxf.is_file():
            print(f"■ {dxf.name} — 파일 없음\n")
            continue
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
        print(f"■ {dxf.name} · 헤드 후보 {len(cand):,}")

        out = {}
        for tag, cls in (("격자 색인(지금)", Counting), ("선형 탐색(예전)", Linear)):
            Counting.n_built = Counting.n_query = 0
            Counting.build_ms = 0.0
            Counting.n_nodes = 0
            A._NearestNodeIndex = cls
            try:
                t0 = time.perf_counter()
                sel = A.select_worst30_heads_anchored(
                    ents, cat, (float(al[0]), float(al[1])), region, k=30,
                    load_mode=False)
                el = (time.perf_counter() - t0) * 1000
            finally:
                A._NearestNodeIndex = real_cls
            out[tag] = el
            extra = ""
            if cls is Counting:
                extra = (f"   색인 {Counting.n_built}회 세움"
                         f"(절점 {Counting.n_nodes:,} · {Counting.build_ms:.0f}ms) "
                         f"· 물음 {Counting.n_query:,}회")
            print(f"    {tag:<16} {el:>9,.0f} ms   헤드 "
                  f"{len(getattr(sel, 'heads', None) or ()):>4}{extra}")
        a, b = out["격자 색인(지금)"], out["선형 탐색(예전)"]
        print(f"    차이 {a - b:+,.0f} ms "
              + ("★색인이 손해다" if a > b * 1.05 else "(색인이 이득이거나 무해)"))
        print()

    print("  색인 세우는 값이 물음 수보다 크면 작은 도면에서 손해다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
