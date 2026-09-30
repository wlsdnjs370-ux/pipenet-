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


#: 교차점에 **노드가 둘 겹친** 자리 — 주배관 노드와 가지관 노드가 한 점에 있고 길이
#  0 에 가까운 연결관으로 이어져 있다(B1F 실측: 괄호 자리 연결관 0.04 mm · 회랑 괄호
#  8곳 중 7곳). 그 연결관의 «방향» 은 좌표 잡음이라(실측 167°) 호가 덮은 쪽에 떨어지면
#  가지가 «주배관 계속» 으로 읽혀 안 올라갔다 [오너 2026-09-22]. 이 거리 안의 이웃은
#  «같은 자리» 로 보고, 방향은 그 너머 실제 배관으로 잰다.
JOINT_M = 0.01


def _joint(xy, u, v):
    ux, uy = xy[u]
    vx, vy = xy[v]
    return math.hypot(vx - ux, vy - uy) <= JOINT_M


def _through_angles(xy, adj, u, v):
    """u 에서 겹친 노드 v 를 지나 **실제로 뻗는** 배관들의 방향(도)."""
    ux, uy = xy[u]
    out = []
    for w in adj.get(v, ()):
        if w == u or _joint(xy, u, w):
            continue
        wx, wy = xy[w]
        out.append(math.degrees(math.atan2(wy - uy, wx - ux)))
    return out


def sit_arcs(xy, ho, sit_r, degree=None, tie_m=0.03):
    """호 → 노드. 한 점에 여러 호를 모은다.

    ★같은 자리에 노드가 둘 이상 겹치면(교차점의 갈래 노드 + 그 자리의 통과 노드 —
      B1F 실측: 괄호 호 346자리 중 195자리가 그렇다) 접속이 많은 쪽에 앉힌다.
      거리로만 고르면 «어느 것이 먼저 나오나» 에 따라 갈래 노드를 놓치고, 그러면
      walk_main 이 그 호를 「접속 2개 이하 — 통과」로 버려 가지가 안 올라간다.
      `degree` 를 안 주면 종전대로 거리만 본다.
    """
    node_arcs = defaultdict(list)
    for sp in ho:
        if "connection_node" in sp:
            nid = sp["connection_node"]
            if nid in xy:
                node_arcs[nid].append(sp)
            continue  # an absent/removed proven tee is not a proximity match
        cx, cy = float(sp["cx"]), float(sp["cy"])
        r = float(sp.get("r") or 0.0)
        lim = max(sit_r, r)
        if sp.get("connection_evidence") == "arc_open_branch_unique_through":
            # The source symbol centre is retained. Seat the fitting at the
            # explicitly restored tee, not at the nearby straight branch tip.
            cx, cy = map(float, sp["connection_xy"])
            lim = 1e-6
        best = None
        for nid, (x, y) in xy.items():
            d = math.hypot(x - cx, y - cy)
            if d > lim:
                continue
            if best is None:
                best = (d, nid)
                continue
            if degree is not None and abs(d - best[0]) <= tie_m:
                # 같은 자리 — 접속이 많은 노드(갈래)가 이긴다. 같으면 가까운 쪽.
                if (int(degree.get(nid, 0)), -d) > (int(degree.get(best[1], 0)), -best[0]):
                    best = (d, nid)
                continue
            if d < best[0]:
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


def walk_main(xy, adj, node_arcs, seed, degree=None):
    """메인 노드·가지 첫 간선. 열린 곳으로 나가지 않는다.

    `degree` : {노드: 지우기 전 배관망의 접속 수}. 회랑(최불리 전개)은 갈래를
    잘라내 교차점이 접속 2개로 보이지만 건물에는 그대로 티가 있다 — 그 노드의
    호는 「메인 위 기호」가 아니라 갈래 표시다. 없으면 지금 접속 수를 쓴다.
    """
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
        # 겹친 노드(JOINT_M) 는 한 자리다 — 거기 앉은 호도 이 자리의 호다.
        joint = {v for v in adj.get(u, ()) if _joint(xy, u, v)}
        if joint:
            arcs = list(arcs) + [a for v in sorted(joint, key=str)
                                 for a in (node_arcs.get(v) or ())]
        deg_u = int((degree or {}).get(u, len(adj.get(u, ()))))
        if max(deg_u, len(adj.get(u, ()))) <= 2:
            # 접속 배관 2 이하면 갈래가 아니라 메인 위 기호 — 통과 [오너 2026-08-19]
            arcs = ()
        ux, uy = xy[u]
        for v in nxt:
            vx, vy = xy[v]
            e = ek(u, v)
            if arcs and v in joint:
                # 0 길이 연결관 — 그 너머 배관이 **전부** 열린 쪽이면 가지의 첫 간선.
                #   하나라도 덮인 쪽이면 종전대로 메인으로 걷는다(추측하지 않는다).
                angs = _through_angles(xy, adj, u, v)
                if angs and all(is_open(arcs, a) for a in angs):
                    branch_e.add((u, v))
                    continue
            elif arcs and is_open(arcs, math.degrees(math.atan2(vy - uy, vx - ux))):
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
        if sp.get("connection_xy") is not None:
            row["connection_xy"] = list(xf_mm_to_m(*sp["connection_xy"], minx, miny))
        out.append(row)
    return out
