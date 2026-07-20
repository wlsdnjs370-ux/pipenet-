# -*- coding: utf-8 -*-
"""전이 실패 진단 — 실제 DXF 의 끊긴 끝단이 '끝점↔끝점'인가 '끝점↔본관중간(T분기)'인가.

SDF 정답망은 모든 연결이 노드(=끝점↔끝점)다. 그런데 모델이 실제 DXF 에서 연결을
하나도 제안 못 했다. 가설: 실제 CAD 파탄의 참 연결은 가지배관 끝이 본관 **중간**에
닿는 T분기(끝점↔edge)라, 끝점↔끝점만 후보로 보는 현재 방식이 애초에 못 잡는다.

각 dangling 끝단에 대해:
  · 최근접 '다른 끝단' 거리   (끝점↔끝점 후보가 닿을 수 있는 거리)
  · 최근접 '본관 edge' 거리   (자기 자신 edge 제외, 점↔선분)
두 분포를 비교. edge 거리 ≪ 끝단 거리면 T분기가 지배적 → 모델 설계 수정 필요.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_HERE = Path(__file__).resolve().parent
for _p in (str(_ROOT), str(_HERE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from clean_candidate_survey import _raw_graph  # noqa: E402
from remote30_prototype import parse_dxf_for_view  # noqa: E402

DEFAULT = [
    "data/sample_problem/대명동201동 계통도_최소.dxf",
    "samples/dxf/계통도_LH_306.dxf",
]


def _pt_seg_dist(p, a, b):
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    if L2 <= 1e-18:
        return math.hypot(p[0] - ax, p[1] - ay)
    t = ((p[0] - ax) * dx + (p[1] - ay) * dy) / L2
    t = max(0.0, min(1.0, t))
    cx, cy = ax + t * dx, ay + t * dy
    return math.hypot(p[0] - cx, p[1] - cy)


def _pct(vals, ps=(10, 25, 50, 75, 90)):
    if not vals:
        return "n/a"
    s = sorted(vals)
    out = []
    for p in ps:
        k = (len(s) - 1) * p / 100.0
        lo, hi = math.floor(k), math.ceil(k)
        v = s[lo] if lo == hi else s[lo] + (s[hi] - s[lo]) * (k - lo)
        out.append(f"p{p}={v:.2f}")
    return "  ".join(out)


def diag(dxf: Path):
    parsed = parse_dxf_for_view(dxf, include_hidden_layers=True)
    graph, edge_len, scale = _raw_graph(parsed["entities"])
    segs = list(edge_len.keys())  # (a,b) 좌표쌍
    deg = {u: len(v) for u, v in graph.items()}
    dangling = [u for u, d in deg.items() if d == 1]
    med_edge = sorted(edge_len.values())[len(edge_len) // 2] if edge_len else 1.0

    tip_d, edge_d = [], []
    for u in dangling:
        nb = next(iter(graph[u]))  # 자기 edge 의 반대끝
        # 최근접 다른 끝단
        best_t = min((math.hypot(w[0] - u[0], w[1] - u[1])
                      for w in dangling if w is not u), default=float("inf"))
        # 최근접 edge (자기 자신 edge 제외)
        best_e = float("inf")
        for a, b in segs:
            if (a == u and b == nb) or (a == nb and b == u):
                continue
            if a == u or b == u:  # u 가 끝점인 다른 edge 도 제외
                continue
            d = _pt_seg_dist(u, a, b)
            if d < best_e:
                best_e = d
        if math.isfinite(best_t):
            tip_d.append(best_t)
        if math.isfinite(best_e):
            edge_d.append(best_e)

    print("\n" + "=" * 92)
    print(f"진단 — {dxf.name}  (scale={scale:.5f}, med_edge={med_edge:.3f}, dangling={len(dangling)})")
    print("=" * 92)
    print(f"  끝단→최근접 끝단 거리 : {_pct(tip_d)}")
    print(f"  끝단→최근접 edge 거리 : {_pct(edge_d)}")
    # 정규화 비교
    tn = sorted(d / med_edge for d in tip_d)
    en = sorted(d / med_edge for d in edge_d)
    if tn and en:
        mt, me = tn[len(tn) // 2], en[len(en) // 2]
        print(f"  중앙값(정규화): 끝단={mt:.2f}  edge={me:.2f}  → "
              f"{'T분기(끝점↔edge) 지배 — 끝점끝점만으론 못잡음' if me < mt * 0.6 else '끝점↔끝점 우세'}")
        # edge 가 더 가까운 끝단 비율
        closer_edge = sum(1 for u_i in range(min(len(tip_d), len(edge_d)))
                          if edge_d[u_i] < tip_d[u_i])
        # 위 인덱스 정렬이 깨졌으니 재계산
    # 재계산: 끝단별 edge<tip 비율 (정렬 안 한 원본 필요) — 아래서 다시
    return tip_d, edge_d, med_edge


def diag2(dxf: Path):
    """끝단별 edge거리 < 끝단거리 비율 (정렬 영향 없이)."""
    parsed = parse_dxf_for_view(dxf, include_hidden_layers=True)
    graph, edge_len, scale = _raw_graph(parsed["entities"])
    segs = list(edge_len.keys())
    deg = {u: len(v) for u, v in graph.items()}
    dangling = [u for u, d in deg.items() if d == 1]
    closer = 0
    total = 0
    for u in dangling:
        nb = next(iter(graph[u]))
        best_t = min((math.hypot(w[0] - u[0], w[1] - u[1])
                      for w in dangling if w is not u), default=float("inf"))
        best_e = float("inf")
        for a, b in segs:
            if a == u or b == u:
                continue
            d = _pt_seg_dist(u, a, b)
            if d < best_e:
                best_e = d
        if math.isfinite(best_t) and math.isfinite(best_e):
            total += 1
            if best_e < best_t:
                closer += 1
    if total:
        print(f"  끝단 중 'edge 가 더 가까운'(=T분기 후보) 비율: {closer}/{total} "
              f"= {closer / total * 100:.0f}%")


def main(argv):
    targets = argv[1:] if len(argv) > 1 else DEFAULT
    for t in targets:
        p = Path(t)
        if not p.is_absolute():
            p = _ROOT / t
        if not p.is_file():
            print(f"(없음) {t}")
            continue
        diag(p)
        diag2(p)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
