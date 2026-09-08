# -*- coding: utf-8 -*-
"""[계통도 꼬임 ②] 뽑힌 경로가 «추정 이음» 을 몇 번 밟는지 잰다.

레이어 섞기 말고도 의심 가는 자리가 하나 더 있다: 엔진은 조각난 배관망을
허용오차 다리(200mm~10m)로 잇는데, 그 다리는 라우팅에서 **실배관과 같은
값**이다. 짧은 직선이라 최단경로가 즐겨 밟고, 화면에는 실배관과 똑같이
그려진다 — 도면에 없는 선을 타고 건너뛰니 사람 눈에 «꼬임» 이다.

경로가 다리를 몇 개·몇 m 밟는지, 다리를 아예 금지하면 어떻게 되는지 잰다.
"""
from __future__ import annotations

import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_D = os.path.join(_ROOT, "routes", "제출용[최종]")


def _keys(stats, name):
    out = set()
    for (a, b) in (stats.get(name) or ()):
        ka, kb = (int(a[0]), int(a[1])), (int(b[0]), int(b[1]))
        out.add((min(ka, kb), max(ka, kb)))
    return out


def walk(path, edge_len, tol_keys):
    """경로를 걸으며 (전체 m, 다리 개수, 다리 m)."""
    tot = brg = n = 0.0
    for u, v in zip(path, path[1:]):
        ln = edge_len.get((min(u, v), max(u, v)))
        if ln is None:
            ln = math.hypot(u[0] - v[0], u[1] - v[1])
        tot += ln
        ku = (int(round(u[0])), int(round(u[1])))
        kv = (int(round(v[0])), int(round(v[1])))
        if (min(ku, kv), max(ku, kv)) in tol_keys:
            n += 1
            brg += ln
    return tot, int(n), brg


def probe(label, fn, a_xy, b_xy, layers=None):
    from remote30_prototype import (_nearest_graph_node, _shortest_path,
                                    build_system_graph)
    from routes.module_f.subdrawing import parse_subdrawing
    path_dxf = os.path.join(_D, fn)
    if not os.path.isfile(path_dxf):
        print(f"표본 없음: {path_dxf}")
        return
    entities, _ = parse_subdrawing(path_dxf)
    graph, edge_len, stats = build_system_graph(
        entities, layer_filter=layers, force_connect=True)
    tol = _keys(stats, "tolerance_bridge_edges")
    forced = _keys(stats, "forced_bridge_edges")
    a = _nearest_graph_node(graph, a_xy)
    b = _nearest_graph_node(graph, b_xy)
    if isinstance(a, tuple) and len(a) == 2 and isinstance(a[0], tuple):
        a, b = a[0], b[0]
    print(f"\n=== {label} · 레이어 {layers or '자동(섞음)'}")
    print(f"  절점 {len(graph)} · 허용오차 다리 {len(tol)} · 강제 다리 {len(forced)}")
    p = _shortest_path(graph, edge_len, a, b, penalty_keys=forced)
    if not p:
        print("  경로 없음")
        return
    tot, nb, mb = walk(p, edge_len, tol)
    straight = math.hypot(a[0] - b[0], a[1] - b[1])
    print(f"  경로 절점 {len(p)} · 연장 {tot / 1000:.1f} m"
          f" · 직선 {straight / 1000:.1f} m")
    print(f"  ★그 중 추정 이음 {nb}곳 · {mb / 1000:.1f} m"
          f" ({(mb / tot * 100 if tot else 0):.0f}%)")
    # ★«꼬임» 의 알맹이 — 경로가 계통(레이어) 사이를 몇 번 넘나드는가.
    #   HSP(고층)와 LSP(저층)는 서로 다른 배관인데 한 그래프에 있으면 최단
    #   경로가 둘을 오갈 수 있다. 넘나든 횟수를 그대로 센다.
    from remote30_prototype import _auto_pipe_layer_filter
    own = {}
    for nm in sorted(layers or _auto_pipe_layer_filter(entities)):
        g1, el1, _s1 = build_system_graph(entities, layer_filter={nm},
                                          force_connect=False)
        own[nm] = set(el1)
    seq = []
    for u, v in zip(p, p[1:]):
        k = (min(u, v), max(u, v))
        hit = [nm for nm, ks in own.items() if k in ks]
        seq.append(hit[0] if len(hit) == 1 else ("?" if not hit else "+"))
    runs = [seq[0]] if seq else []
    for s in seq[1:]:
        if s != runs[-1]:
            runs.append(s)
    print(f"  ★경로가 지나는 레이어: {' → '.join(runs)}"
          f"  (넘나듦 {max(0, len(runs) - 1)}회)")
    p2 = _shortest_path(graph, edge_len, a, b, penalty_keys=forced | tol)
    if p2:
        t2, n2, m2 = walk(p2, edge_len, tol)
        print(f"  [다리를 뒤로 미루면] 절점 {len(p2)} · 연장 {t2 / 1000:.1f} m"
              f" · 추정 이음 {n2}곳 {m2 / 1000:.1f} m")


def main() -> int:
    # 화면에서 사람이 찍는 자리와 비슷하게 — 도면 아래쪽(펌프)·위쪽(AV).
    from routes.module_f.subdrawing import parse_subdrawing
    for label, fn in (("계통도", "1. 입력도면 대명동 단위세대 계통도.dxf"),
                      ("기계실", "1. 입력도면 대명동 단위세대 기계실.dxf")):
        p = os.path.join(_D, fn)
        if not os.path.isfile(p):
            continue
        ents, _ = parse_subdrawing(p)
        xs = [c for en in ents for c in (en.get("p") or [])[0::2]
              if isinstance(c, (int, float))]
        ys = [c for en in ents for c in (en.get("p") or [])[1::2]
              if isinstance(c, (int, float))]
        if not xs:
            continue
        cx = (min(xs) + max(xs)) / 2
        lo = (cx, min(ys) + (max(ys) - min(ys)) * 0.18)
        hi = (cx, min(ys) + (max(ys) - min(ys)) * 0.82)
        probe(label, fn, lo, hi)
        from remote30_prototype import _auto_pipe_layer_filter
        auto = sorted(_auto_pipe_layer_filter(ents))
        if auto:
            probe(label, fn, lo, hi, layers={auto[0]})
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
