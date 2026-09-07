# -*- coding: utf-8 -*-
"""[§27 판단 재료] 뽑힌 경로에서 «도면에 없는 다리» 가 몇 %인가.

「허용치 다리를 추정으로 표시할 것인가」는 사람이 정할 문제인데, 그 판단은
**다리가 얼마나 긴가**에 달렸다. 5mm 짜리 붙임이면 점선으로 갈라 봐야 눈만
어지럽고, 미터 단위면 도면에 없는 선을 실측인 양 그리고 있는 것이다.

그래서 두 가지를 잰다:
  ① 다리 길이의 분포 — 몇 mm 짜리인가
  ② 실제로 뽑히는 경로에서 다리가 차지하는 연장 비율
"""
from __future__ import annotations

import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(_ROOT)
sys.path.insert(0, _ROOT)

from routes.module_f.common import _boot          # noqa: E402

_boot()

import remote30_prototype as RP                   # noqa: E402
from routes.module_f.subdrawing import parse_subdrawing   # noqa: E402

TOLS = (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0)


def _key(a, b):
    return (a, b) if a <= b else (b, a)


def _build(ents, layers):
    """`build_system_graph` 와 **같은 순서** 로 세우되 다리를 다 기록한다."""
    all_lines = [en for en in ents if en.get("t") in ("L", "PL")]
    if layers is None:
        matched = RP._auto_pipe_layer_filter(ents)
        lines = [en for en in all_lines if en.get("l") in matched]
    else:
        lines = [en for en in all_lines if en.get("l") in layers]
    ratio = RP._drawing_scale_ratio(lines)
    idx = RP._NodeIndex(epsilon_mm=RP.SNAP_TOL_MM * ratio) if ratio < 1.0 else None
    graph, edge_len = RP._build_graph(
        lines, node_index=idx, min_edge_mm=RP.MIN_PIPE_EDGE_MM * ratio)
    real = {_key(a, b) for a, nbrs in graph.items() for b in nbrs}
    bridges: dict = {}
    for tol in TOLS:
        got: set = set()
        RP._bridge_components(graph, edge_len, max_bridge_mm=tol * ratio,
                              bridge_edges_out=got)
        for k in got:
            kk = _key(k[0], k[1])
            if kk not in real:
                bridges.setdefault(kk, tol)
    forced: set = set()
    RP._bridge_components(graph, edge_len, max_bridge_mm=float("inf"),
                          bridge_edges_out=forced)
    for k in forced:
        kk = _key(k[0], k[1])
        if kk not in real:
            bridges.setdefault(kk, float("inf"))
    return graph, edge_len, real, bridges


def _len(edge_len, a, b):
    return float(edge_len.get((a, b)) or edge_len.get((b, a)) or 0.0)


def _report(name, path, layers=None):
    ents, _ = parse_subdrawing(path)
    graph, edge_len, real, bridges = _build(ents, layers)
    print(f"\n[{name}]  절점 {len(graph)} · 실측 간선 {len(real)}"
          f" · 다리 {len(bridges)}")
    lens = sorted(_len(edge_len, a, b) for (a, b) in bridges)
    if lens:
        n = len(lens)
        print(f"  다리 길이(mm) — 최소 {lens[0]:.0f} · 중앙 {lens[n // 2]:.0f}"
              f" · 최대 {lens[-1]:.0f} · 합계 {sum(lens) / 1000:.2f} m")
        band = {"≤10mm": 0, "≤100mm": 0, "≤1m": 0, "≤10m": 0, "그 이상": 0}
        for v in lens:
            for k, lim in (("≤10mm", 10), ("≤100mm", 100), ("≤1m", 1000),
                           ("≤10m", 10000)):
                if v <= lim:
                    band[k] += 1
                    break
            else:
                band["그 이상"] += 1
        print("  띠별 개수: " + " · ".join(f"{k} {v}" for k, v in band.items()))

    # 실제로 뽑히는 경로에 얼마나 섞이나 — 한 쌍만 보면 운에 좌우된다.
    # 절점 쌍을 골고루 잡아 분포로 본다(사람이 어디를 찍을지 모르므로).
    nodes = list(graph)
    if len(nodes) < 2:
        return
    from remote30_graph import _shortest_path

    step = max(1, len(nodes) // 24)          # 24 × 24 격자꼴 표본
    picks = nodes[::step][:24]
    shares, worst_one, n_with = [], 0.0, 0
    for i, a in enumerate(picks):
        for b in picks[i + 1:]:
            try:
                pn = _shortest_path(graph, edge_len, a, b)
            except Exception:                # noqa: BLE001
                continue
            if not pn or len(pn) < 2:
                continue
            tot = brid = 0.0
            for k in range(len(pn) - 1):
                ln = _len(edge_len, pn[k], pn[k + 1])
                tot += ln
                if _key(pn[k], pn[k + 1]) in bridges:
                    brid += ln
                    worst_one = max(worst_one, ln)
            if tot <= 0:
                continue
            shares.append(brid / tot * 100)
            if brid > 0:
                n_with += 1
    if not shares:
        print("  경로 표본 없음")
        return
    shares.sort()
    n = len(shares)
    print(f"  경로 {n}쌍 — 다리가 섞인 경로 {n_with}개 ({n_with / n * 100:.0f}%)"
          f" · 연장 중 다리 비율 중앙 {shares[n // 2]:.1f}%"
          f" · 90분위 {shares[int(n * 0.9)]:.1f}%"
          f" · 최대 {shares[-1]:.1f}%")
    print(f"  경로에 실제로 쓰인 다리 한 개의 최대 길이: {worst_one:.0f} mm")


if __name__ == "__main__":
    _report("계통도 · 사전 자동",
            "routes/제출용[최종]/1. 입력도면 대명동 단위세대 계통도.dxf")
    _report("기계실 · 사전 자동",
            "routes/제출용[최종]/1. 입력도면 대명동 단위세대 기계실.dxf")
