# -*- coding: utf-8 -*-
"""급수에서 메인 따라가기. 호의 열린 곳으로만 안 간다.

열린 1곳 = 가지 엘보. 열린 2곳 = 가지 크로스.
호 여러 개가 한 점에 있으면 그려진 호가 덮는 각은 열린 곳이 아니다.
엔진·폼·구경산정과 별도. 변환이 부를 때만 쓴다.
"""
from __future__ import annotations

import math
from collections import defaultdict


def _norm360(a):
    a = float(a) % 360.0
    return a + 360.0 if a < 0.0 else a


def _on_drawn(ang, sa, sweep):
    if sweep <= 0.0 or sweep >= 360.0 - 1e-12:
        return True
    d = (_norm360(ang) - _norm360(sa)) % 360.0
    return d <= sweep + 1e-12


def is_open(arcs, ang):
    """그 각이 그려진 호에 안 덮이면 열린 곳."""
    drawn = False
    for a in arcs:
        sa, sw = a.get("sa"), a.get("sweep")
        if sa is None or sw is None:
            continue
        drawn = True
        if _on_drawn(ang, sa, sw):
            return False
    return drawn


def sit_arcs(xy, ho, sit_r):
    """호 → 노드. 한 점에 여러 호를 모은다."""
    node_arcs = defaultdict(list)
    for sp in ho:
        cx, cy = float(sp["cx"]), float(sp["cy"])
        r = float(sp.get("r") or 0.0)
        lim = max(sit_r, r)
        best = None
        for nid, (x, y) in xy.items():
            d = math.hypot(x - cx, y - cy)
            if d <= lim and (best is None or d < best[0]):
                best = (d, nid)
        if best is not None:
            node_arcs[best[1]].append(sp)
    return node_arcs


def snap_seed(xy, adj, src_xy, snap):
    sx, sy = float(src_xy[0]), float(src_xy[1])
    best = None
    for u, vs in adj.items():
        ux, uy = xy[u]
        for v in vs:
            if u > v:
                continue
            vx, vy = xy[v]
            dx, dy = vx - ux, vy - uy
            l2 = dx * dx + dy * dy
            if l2 < 1e-18:
                continue
            t = max(0.0, min(1.0, ((sx - ux) * dx + (sy - uy) * dy) / l2))
            px, py = ux + t * dx, uy + t * dy
            d = math.hypot(sx - px, sy - py)
            if best is None or d < best[0]:
                best = (d, u, v)
    if best is None or best[0] > snap:
        return None
    return (best[1], best[2])


def walk_main(xy, adj, node_arcs, seed):
    """메인 노드·가지 첫 간선. 열린 곳으로 나가지 않는다."""
    main_n, main_e, branch_e = set(), set(), set()
    seen = set()
    todo = [(seed[0], seed[1]), (seed[1], seed[0])]
    main_e.add((seed[0], seed[1]) if seed[0] < seed[1] else (seed[1], seed[0]))
    main_n.add(seed[0])
    main_n.add(seed[1])

    def ek(a, b):
        return (a, b) if a < b else (b, a)

    while todo:
        u, prev = todo.pop()
        if u in seen:
            continue
        seen.add(u)
        main_n.add(u)
        nxt = [v for v in adj.get(u, ()) if v != prev]
        arcs = node_arcs.get(u) or ()
        if len(adj.get(u, ())) <= 2:
            # 접속 배관 2 이하면 갈래가 아니라 메인 위 기호 — 통과 [오너 2026-08-19]
            arcs = ()
        ux, uy = xy[u]
        for v in nxt:
            vx, vy = xy[v]
            e = ek(u, v)
            if arcs and is_open(arcs, math.degrees(math.atan2(vy - uy, vx - ux))):
                branch_e.add((u, v))
                continue
            main_e.add(e)
            todo.append((v, u))
    return {"main_nodes": main_n, "main_edges": main_e,
            "branch_first": branch_e}


def xf_mm_to_m(x, y, minx, miny):
    return ((x - minx) / 1000.0 + 1.0, (y - miny) / 1000.0 + 1.0)


def ho_to_kfp_units(ho, minx, miny):
    out = []
    for sp in ho:
        cx, cy = xf_mm_to_m(sp["cx"], sp["cy"], minx, miny)
        row = dict(sp)
        row["cx"], row["cy"] = cx, cy
        row["r"] = float(sp.get("r") or 0.0) / 1000.0
        out.append(row)
    return out
