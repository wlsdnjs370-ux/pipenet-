# -*- coding: utf-8 -*-
"""계통도_LH_306.dxf 파탄(fragmentation) 패턴 정량화.

link-prediction 연결복원의 학습데이터를 만들려면, 합성 손상(synthetic corruption)이
'실제 도면이 끊기는 방식'과 닮아야 한다. 이 스크립트는 실제 파탄 도면 한 장을 읽어
다음을 측정한다 — 합성 손상 모델의 파라미터로 직접 쓰기 위함:

  1) 연결성     : connected component 수·크기 분포, dangling(차수1) endpoint 수
  2) 틈 거리    : component 를 가로지르는 최근접 endpoint 쌍의 거리 분포 (gap distance)
  3) 직선성     : 끊긴 끝단의 진행방향 vs 파트너로 향하는 방향의 각도차
                  (실제 배관이 일직선으로 이어지다 끊겼다면 각도차≈0)
  4) 심볼 교차  : 틈 중점이 비-배관 엔티티(TEXT/HEAD/ARC/CIRCLE)에 얼마나 가까운가
                  (배관이 심볼·문자 위를 지나며 끊겼는지)

실행:
    python calibration/lh306_fragmentation.py [dxf경로]
    (인자 없으면 계통도_LH_306.dxf)
"""
from __future__ import annotations

import math
import sys
from collections import Counter, defaultdict
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from remote30_prototype import (  # noqa: E402
    parse_dxf_for_view,
    _categorize_layer,
    _build_graph,
    _connected_components,
    _drawing_scale_ratio,
    _NodeIndex,
    SNAP_TOL_MM,
    MIN_PIPE_EDGE_MM,
)

PIPE_LAYERS_HINT = ("PIPE", "PIPE11")  # 메모리: 계통도_LH_306 배관은 PIPE/PIPE11
SEARCH_RADIUS_MM = 20000.0             # dangling 파트너 탐색 최대 반경
NON_PIPE_CATS = {"HEAD", "TEXT", "ALARM", "ARCH"}


def _percentiles(vals, ps=(0, 10, 25, 50, 75, 90, 100)):
    if not vals:
        return {p: float("nan") for p in ps}
    s = sorted(vals)
    out = {}
    for p in ps:
        if p <= 0:
            out[p] = s[0]
        elif p >= 100:
            out[p] = s[-1]
        else:
            k = (len(s) - 1) * p / 100.0
            lo = math.floor(k)
            hi = math.ceil(k)
            out[p] = s[lo] if lo == hi else s[lo] + (s[hi] - s[lo]) * (k - lo)
    return out


def _fmt_p(d):
    return "  ".join(f"p{p}={d[p]:.1f}" for p in sorted(d))


def main(argv):
    dxf = Path(argv[1]) if len(argv) > 1 else _ROOT / "samples" / "dxf" / "계통도_LH_306.dxf"
    if not dxf.is_file():
        print(f"파일 없음: {dxf}")
        return 1

    print("=" * 96)
    print(f"LH306 파탄 패턴 분석 — {dxf.name}")
    print("=" * 96)

    parsed = parse_dxf_for_view(dxf, include_hidden_layers=True)
    ents = parsed["entities"]
    layers = sorted({e.get("l", "") for e in ents})
    cats = {ly: _categorize_layer(ly) for ly in layers}

    cat_count = Counter(cats[e.get("l", "")] for e in ents)
    print(f"\n총 엔티티 {len(ents)}개 · 레이어 {len(layers)}개")
    print("레이어 카테고리 분포:", dict(cat_count))
    pipe_layers = [ly for ly, c in cats.items() if c == "PIPE"]
    print(f"PIPE 카테고리 레이어: {pipe_layers}")

    # ---- 그래프 빌드 (build_system_graph 와 동일한 스케일 적응, 단 bridge 미적용) ----
    # bridge 는 우리가 link-prediction 으로 대체/보강하려는 휴리스틱이므로, 그 전의
    # raw 파탄 상태를 측정한다.
    pipe_ents = [e for e in ents
                 if e.get("t") in ("L", "PL") and cats.get(e.get("l", ""), "OTHER") == "PIPE"]
    scale_ratio = _drawing_scale_ratio(pipe_ents)
    snap_eps = SNAP_TOL_MM * scale_ratio
    min_edge = MIN_PIPE_EDGE_MM * scale_ratio
    print(f"\nscale_ratio={scale_ratio:.6f} · snap_eps={snap_eps:.3f}mm · "
          f"min_edge={min_edge:.3f}mm · PIPE line/PL 엔티티={len(pipe_ents)}")
    idx = _NodeIndex(snap_eps if scale_ratio < 1.0 else SNAP_TOL_MM)
    graph, edge_len = _build_graph(pipe_ents, node_index=idx, min_edge_mm=min_edge)
    nodes = list(graph.keys())
    n_edges = len(edge_len)
    comps = _connected_components(graph)
    comps.sort(key=len, reverse=True)

    print("\n" + "-" * 96)
    print("[1] 연결성")
    print("-" * 96)
    print(f"노드 {len(nodes)} · 엣지 {n_edges} · component {len(comps)}")
    sizes = [len(c) for c in comps]
    print(f"component 크기 상위10: {sizes[:10]}")
    if sizes:
        biggest = sizes[0]
        print(f"최대 component {biggest}노드 = 전체의 {biggest / max(1, len(nodes)) * 100:.1f}%")
    singletons = sum(1 for s in sizes if s == 1)
    print(f"고립 노드(크기1) component: {singletons}개")

    deg = {u: len(vs) for u, vs in graph.items()}
    dangling = [u for u, d in deg.items() if d == 1]
    isolated = [u for u, d in deg.items() if d == 0]
    print(f"dangling(차수1) endpoint: {len(dangling)} · 차수0 고립: {len(isolated)}")

    # 노드→component id 매핑
    comp_id = {}
    for i, c in enumerate(comps):
        for u in c:
            comp_id[u] = i

    # dangling 노드의 진행방향 (유일 이웃 → 자신 방향, 즉 배관이 뻗어나가던 방향)
    def out_dir(u):
        vs = list(graph.get(u, ()))
        if not vs:
            return None
        v = vs[0]
        dx, dy = u[0] - v[0], u[1] - v[1]
        n = math.hypot(dx, dy)
        return (dx / n, dy / n) if n > 0 else None

    # ---- [2]+[3] 틈 거리 & 직선성 : dangling → 다른 component 최근접 endpoint ----
    print("\n" + "-" * 96)
    print("[2] 틈 거리 + [3] 직선성 (dangling endpoint → 타 component 최근접 endpoint)")
    print("-" * 96)

    # endpoint(차수1) 만 후보로 — 끊긴 관은 끝단끼리 이어져야 자연스러움
    endpoints = dangling
    search_radius = SEARCH_RADIUS_MM * scale_ratio
    gap_dists = []
    angle_diffs = []  # 끊긴 끝단 진행방향 vs 파트너로 향하는 방향
    pairs = []
    for u in endpoints:
        best = None
        bestd = search_radius
        for w in endpoints:
            if w is u or comp_id.get(w) == comp_id.get(u):
                continue
            d = math.hypot(w[0] - u[0], w[1] - u[1])
            if d < bestd:
                bestd = d
                best = w
        if best is None:
            continue
        gap_dists.append(bestd)
        du = out_dir(u)
        if du is not None and bestd > 1e-9:
            tx, ty = (best[0] - u[0]) / bestd, (best[1] - u[1]) / bestd
            cosang = max(-1.0, min(1.0, du[0] * tx + du[1] * ty))
            angle_diffs.append(math.degrees(math.acos(cosang)))
        pairs.append((u, best, bestd))

    print(f"파트너 매칭된 dangling: {len(gap_dists)}/{len(endpoints)}")
    if gap_dists:
        print("틈 거리(mm)  :", _fmt_p(_percentiles(gap_dists)))
        # 엣지 길이 분포와 비교 (틈≈한 세그먼트면 '한 칸씩 끊김')
        elens = list(edge_len.values())
        print("엣지 길이(mm):", _fmt_p(_percentiles(elens)))
    if angle_diffs:
        print("직선성 각도차(도, 0=완벽일직선):", _fmt_p(_percentiles(angle_diffs)))
        collinear = sum(1 for a in angle_diffs if a <= 15.0)
        print(f"  ≤15도(일직선 연장) 비율: {collinear}/{len(angle_diffs)} "
              f"= {collinear / len(angle_diffs) * 100:.0f}%")

    # ---- [4] 심볼/문자 교차 : 틈 중점 ↔ 최근접 비-배관 엔티티 ----
    print("\n" + "-" * 96)
    print("[4] 심볼 교차 (틈 중점 → 최근접 비-배관 엔티티 거리)")
    print("-" * 96)

    sym_pts = []  # 비-배관 엔티티 대표점
    for e in ents:
        c = cats.get(e.get("l", ""), "OTHER")
        if c not in NON_PIPE_CATS:
            continue
        t = e["t"]
        if t == "L":
            x1, y1, x2, y2 = e["p"]
            sym_pts.append(((x1 + x2) / 2, (y1 + y2) / 2))
        elif t == "PL":
            pts = e["p"]
            if pts:
                sym_pts.append((sum(p[0] for p in pts) / len(pts),
                                sum(p[1] for p in pts) / len(pts)))
        elif t in ("A", "C"):
            cc = e.get("c")
            if cc:
                sym_pts.append((cc[0], cc[1]))

    print(f"비-배관 대표점 {len(sym_pts)}개")
    if sym_pts and pairs:
        # 공간 격자로 최근접 근사 (종이축척에 맞춰 셀 크기 조정)
        cell = max(1.0, 2000.0 * scale_ratio)
        grid = defaultdict(list)
        for px, py in sym_pts:
            grid[(int(px // cell), int(py // cell))].append((px, py))

        def nearest_sym(mx, my):
            kx, ky = int(mx // cell), int(my // cell)
            best = float("inf")
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    for px, py in grid.get((kx + dx, ky + dy), ()):
                        d = math.hypot(px - mx, py - my)
                        if d < best:
                            best = d
            return best

        sym_d = []
        for u, w, _g in pairs:
            mx, my = (u[0] + w[0]) / 2, (u[1] + w[1]) / 2
            d = nearest_sym(mx, my)
            if math.isfinite(d):
                sym_d.append(d)
        if sym_d:
            print("틈중점→심볼 거리(mm):", _fmt_p(_percentiles(sym_d)))
            # 종이축척: '심볼 위' 판정은 틈 거리(=한 세그먼트) 규모로. 50mm*scale.
            on_sym = 50.0 * scale_ratio
            near = sum(1 for d in sym_d if d <= on_sym)
            print(f"  ≤{on_sym:.2f}mm(심볼 위에서 끊김 의심) 비율: {near}/{len(sym_d)} "
                  f"= {near / len(sym_d) * 100:.0f}%")

    print("\n" + "=" * 96)
    print("요약 — 합성손상 모델 파라미터로 쓸 값")
    print("=" * 96)
    if gap_dists:
        gp = _percentiles(gap_dists)
        print(f"  · 틈 거리 중앙값 {gp[50]:.0f}mm (p10~p90: {gp[10]:.0f}~{gp[90]:.0f})")
    if angle_diffs:
        ap = _percentiles(angle_diffs)
        coll = sum(1 for a in angle_diffs if a <= 15.0) / len(angle_diffs) * 100
        print(f"  · 직선성: 중앙 각도차 {ap[50]:.0f}도, 일직선(≤15도) {coll:.0f}%")
    print(f"  · 파탄도: {len(comps)} component / {len(nodes)} 노드 "
          f"(최대 component {sizes[0] / max(1, len(nodes)) * 100:.0f}%)" if sizes else "")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
