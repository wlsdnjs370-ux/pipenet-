# -*- coding: utf-8 -*-
"""[감사] 최불리 선정이 정말 «가장 불리한» 것을 잡는가.

사용자 지적: 「가장 멀리 있는 헤드들은 인식을 못 하고, 오히려 덜 먼 헤드들을
먼저 엮어 버린다」.

말로 다투지 말고 숫자로 본다. 급수원 기준 유하거리로 헤드를 줄 세우고,
뽑힌 K개가 그 줄의 어디에 있는지 잰다.

  · 앵커(기준 헤드)가 정말 1등인가
  · 뽑힌 K개 안에 «가장 먼 K개» 가 몇 개나 들어 있나
  · 안 뽑힌 먼 헤드가 있다면 그 거리는 얼마인가
"""
from __future__ import annotations

import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(_ROOT)
sys.path.insert(0, _ROOT)

from routes.module_f.common import _boot          # noqa: E402

_boot()

from services.cad_import.design.worst import worst_k_heads   # noqa: E402
from services.cad_import.edit.session import EditSession     # noqa: E402


def _dijkstra(pts, edges, seeds):
    import heapq
    adj: dict = {}
    for a, b in edges:
        adj.setdefault(a, []).append(b)
        adj.setdefault(b, []).append(a)
    dist: dict = {}
    pq: list = []
    for s in seeds:
        if isinstance(s, int) and 0 <= s < len(pts):
            dist[s] = 0.0
            heapq.heappush(pq, (0.0, s))
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, float("inf")):
            continue
        for v in adj.get(u, ()):
            nd = d + math.dist(pts[u], pts[v])
            if nd < dist.get(v, float("inf")):
                dist[v] = nd
                heapq.heappush(pq, (nd, v))
    return dist


def main() -> int:
    key = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")
    k = int(os.environ.get("MF_K", "30"))
    es = EditSession.open(key, out_dir=None, load_saved=True, use_cache=True)
    b = es.board
    if not b.sources:
        print(f"«{key}» 에 급수원(알람밸브)이 없다 — 감사 불가")
        return 0

    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=k,
                      source_index=0)
    picked = list(w["heads"])
    anchor = w["worst_head"]

    # 같은 자로 전체 헤드의 유하거리를 다시 잰다.
    src = _dijkstra(b.pts, b.edges, [list(b.sources)[0]])
    far: dict = {}
    for hi, nodes in enumerate(b.hnodes):
        reach = [n for n in nodes if n in src]
        if reach:
            far[hi] = min(src[n] for n in reach)
    order = sorted(far, key=far.get, reverse=True)       # 먼 순서

    print(f"[{key}]  헤드 {len(b.hnodes)} · 도달 {len(far)} · K={k}")
    print(f"  앵커(기준 헤드) = h{anchor} · {far.get(anchor, 0) / 1000:.2f} m")
    rank = order.index(anchor) + 1 if anchor in order else -1
    print(f"  ★앵커의 «먼 순서» 등수: {rank} / {len(order)}"
          + ("  ← 1등이면 정상" if rank == 1 else "  ← 1등이 아니다!"))

    topk = set(order[:k])
    inter = topk & set(picked)
    print(f"\n  뽑힌 {len(picked)}개 · 가장 먼 {k}개와 겹침 {len(inter)}개"
          f" ({len(inter) / max(k, 1) * 100:.0f}%)")
    pf = sorted((far[hi] for hi in picked), reverse=True)
    print(f"  뽑힌 것의 유하거리: 최대 {pf[0] / 1000:.2f} m ·"
          f" 최소 {pf[-1] / 1000:.2f} m")
    print(f"  전체 유하거리    : 최대 {far[order[0]] / 1000:.2f} m ·"
          f" 최소 {far[order[-1]] / 1000:.2f} m")

    missed = [hi for hi in order[:k] if hi not in set(picked)]
    print(f"\n  ★«가장 먼 {k}개» 중 안 뽑힌 것 {len(missed)}개")
    for hi in missed[:12]:
        d = b.disks[hi] if hi < len(b.disks) else (0, 0)
        print(f"      h{hi:<5} 유하 {far[hi] / 1000:>7.2f} m"
              f"  (전체 {order.index(hi) + 1}등)  좌표 ({d[0]:.0f}, {d[1]:.0f})")

    # 뽑힌 것들이 앵커 주위에 뭉쳐 있는가 — 설계면적의 뜻이 그것이다.
    an = _dijkstra(b.pts, b.edges, [min(
        (n for n in b.hnodes[anchor] if n in src), key=lambda n: src[n])])
    spread = sorted(
        an.get(min((n for n in b.hnodes[hi] if n in src),
                   key=lambda n: src[n]), float("inf")) for hi in picked)
    print(f"\n  앵커에서 뽑힌 것까지(배관거리): 중앙 {spread[len(spread) // 2] / 1000:.2f} m"
          f" · 최대 {spread[-1] / 1000:.2f} m")
    print(f"  (산출 span_m = {w['span_m']} m · far_m = {w['far_m']} m)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
