# -*- coding: utf-8 -*-
"""격자 색인이 선형 탐색과 «같은 노드» 를 고르는가.

`_nearest_graph_node` 는 헤드마다 전 노드를 훑는다(B1F 실측 9,810회 · 10.8초).
`_NearestNodeIndex` 로 대신하려면 빠른 것만으론 부족하다 — **같은 답**이어야
한다. 특히 무승부(같은 거리) 규칙까지 같아야 하는데, 원본은 `d < bestd` 로만
갈아치우므로 «먼저 삽입된 노드» 가 이긴다.

실제 도면의 실제 그래프로 잰다. 합성 그래프로 재면 이 도면에서 같은지를 못
잰다 — 그래서 A 의 그래프 빌드를 가로채 그대로 쓴다.

    python scripts/_verify_nearest_index.py [도면.dxf ...]
"""
from __future__ import annotations

import random
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf",
    ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf",
]
GRAB: list = []


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A
    from remote30_graph import _NearestNodeIndex, _nearest_graph_node
    from routes.module_f.auto import (detect_head_candidates, head_region_of,
                                      parse_plan, region_around)

    real = A._build_graph

    def hook(*a, **kw):
        g, el = real(*a, **kw)
        if len(g) > 200:
            GRAB.append(g)
        return g, el

    A._build_graph = hook
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
            A.select_worst30_heads_anchored(
                ents, cat, (float(al[0]), float(al[1])),
                head_region_of(region_around(cand, al)),
                k=max(1, len(cand)), load_mode=False)

            graph = max(GRAB, key=len)
            print(f"■ {dxf.name}")
            print(f"    그래프 절점 {len(graph):,} · 물어볼 점 {len(hp):,}")

            t0 = time.perf_counter()
            idx = _NearestNodeIndex(graph)
            t_build = (time.perf_counter() - t0) * 1000

            # 헤드 전부 + 그래프 밖 무작위 점(격자 테 확장을 실제로 태운다)
            rnd = random.Random(20260828)
            xs = [n[0] for n in graph]
            ys = [n[1] for n in graph]
            extra = [(rnd.uniform(min(xs), max(xs)), rnd.uniform(min(ys), max(ys)))
                     for _ in range(500)]
            pts = hp + extra

            t0 = time.perf_counter()
            fast = [idx.nearest(p) for p in pts]
            t_fast = (time.perf_counter() - t0) * 1000
            t0 = time.perf_counter()
            slow = [_nearest_graph_node(graph, p) for p in pts]
            t_slow = (time.perf_counter() - t0) * 1000

            bad = [i for i, (f, s) in enumerate(zip(fast, slow)) if f != s]
            ok = not bad
            print(f"    색인 세우기 {t_build:>8,.0f} ms")
            print(f"    선형 탐색   {t_slow:>8,.0f} ms")
            print(f"    격자 색인   {t_fast:>8,.0f} ms   "
                  f"({t_slow / max(1e-9, t_fast):.0f}배)")
            print(f"    다른 답 {len(bad)} / {len(pts):,}   "
                  f"{'[OK]' if ok else '[FAIL]'}")
            if bad:
                for i in bad[:5]:
                    p = pts[i]
                    df = ((fast[i][0] - p[0]) ** 2 + (fast[i][1] - p[1]) ** 2) ** .5
                    ds = ((slow[i][0] - p[0]) ** 2 + (slow[i][1] - p[1]) ** 2) ** .5
                    print(f"      {p} → 색인 {fast[i]} ({df:.3f}) vs "
                          f"선형 {slow[i]} ({ds:.3f})")
            print()
    finally:
        A._build_graph = real

    print("  «다른 답 0» 이어야 갈아끼울 수 있다. 하나라도 다르면 답이 바뀐다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
