# -*- coding: utf-8 -*-
"""[회랑 사슬좌표 §2] 바꾸기 전에 재는 네 가지 — 회전각 · 병합 · 최원길이 · 나무.

지시서 `ModuleF_회랑_사슬좌표_지시서.md` §2 · 그림 6·7·8.

  ★그림 6 ② 「병합」은 앞 작업에서 **안 센 값**이다. 앞 프로브의 P1 은 「대응
    하는 노드가 있나」만 봐서, 두 board 노드가 **한** kfp 노드에 같이 대응해도
    둘 다 통과시켰다. 여기서는 «1:1 인가» 를 센다.

  병합을 세는 자는 추측이 아니라 `planar.py` 의 `xform` **그 식**이다::

      mx = (x_mm − minx)/1000 + 1        snap = round(round(mx/GRID_M)*GRID_M, 3)

  같은 (snap_x, snap_y) 로 가는 G_K 노드가 둘 이상이면 그 자리가 병합이다.
"""
from __future__ import annotations

import math
from collections import defaultdict, deque

GRID_M = 0.05
DIRS8 = [(math.cos(math.radians(a)), math.sin(math.radians(a)), a)
         for a in (0, 45, 90, 135, 180, 225, 270, 315)]


def snap_key(x_mm, y_mm, minx, miny):
    """`planar.py` 의 xform 과 **같은 식**. 추측하지 않는다."""
    mx = (float(x_mm) - float(minx)) / 1000.0 + 1.0
    my = (float(y_mm) - float(miny)) / 1000.0 + 1.0
    return (round(round(mx / GRID_M) * GRID_M, 3),
            round(round(my / GRID_M) * GRID_M, 3))


def nearest_dir(dx, dy):
    """8 방향 중 가장 가까운 것 → (ux, uy, 각도, 편차°). 같으면 작은 각."""
    L = math.hypot(dx, dy)
    if L < 1e-12:
        return 1.0, 0.0, 0.0, 0.0
    ux, uy = dx / L, dy / L
    best = None
    for cx, cy, a in DIRS8:
        dev = math.degrees(math.acos(max(-1.0, min(1.0, ux * cx + uy * cy))))
        if best is None or dev < best[3] - 1e-12:
            best = (cx, cy, a, dev)
    return best


def measure_turns(gk_edges, board_pts):
    """[회전각] 회랑 평면 배관이 8 방향에서 얼마나 벗어나 있나."""
    rows = []
    for (u, v) in sorted({(min(a, b), max(a, b)) for a, b in gk_edges}):
        if not (0 <= u < len(board_pts) and 0 <= v < len(board_pts)):
            continue
        a = board_pts[u]
        b = board_pts[v]
        dx, dy = float(b[0]) - float(a[0]), float(b[1]) - float(a[1])
        L = math.hypot(dx, dy)
        if L < 1e-9:
            continue
        _ux, _uy, ang, dev = nearest_dir(dx, dy)
        rows.append({"u": u, "v": v, "L": L, "dev": dev, "dir": ang,
                     "slip": L * math.sin(math.radians(dev))})
    cuts = (0.5, 5.0, 10.0, 22.5)
    dist = {c: sum(1 for r in rows if r["dev"] <= c) for c in cuts}
    return {"n": len(rows), "dist": dist,
            "max": max((r["dev"] for r in rows), default=0.0),
            "rows": sorted(rows, key=lambda r: -r["dev"])}


def head_slip_bound(worst, board_pts, turns):
    """[회전각] 헤드마다 Σ L·sin(편차) — 사슬이 밀릴 수 있는 상한."""
    slip = {(min(r["u"], r["v"]), max(r["u"], r["v"])): r["slip"]
            for r in turns["rows"]}
    adj = defaultdict(list)
    for (a, b) in (worst.get("edges") or ()):
        adj[a].append(b)
        adj[b].append(a)
    # 뿌리 = 최원 경로의 첫 절점 (급수원 쪽)
    wp = list(worst.get("worst_path") or ())
    if not wp:
        return {"max": 0.0, "acc": {}}
    src = wp[0]
    acc, seen = {src: 0.0}, {src}
    q = deque([src])
    while q:
        cur = q.popleft()
        for nxt in adj.get(cur, ()):
            if nxt in seen:
                continue
            seen.add(nxt)
            acc[nxt] = acc[cur] + slip.get((min(cur, nxt), max(cur, nxt)), 0.0)
            q.append(nxt)
    # 헤드가 붙은 절점의 누적값이 그 헤드의 이탈 상한이다(부르는 쪽이 고른다).
    return {"max": max(acc.values(), default=0.0), "acc": acc}


def measure_merge(gk_nodes, board_pts, origin_mm, grid=True):
    """[병합] 서로 다른 G_K 노드가 한 칸으로 눌린 자리 (그림 6 ②).

    ★`grid=False` 는 **격자 스냅을 끈 회랑**을 잴 때다. 그때 눌리는 것은
      키가 완전히 같은 «중복 절점» 뿐이라 정상이다 — 켜진 것처럼 재면
      이미 고친 것을 여전히 깨진 것으로 센다(내가 한 번 그렇게 셌다).
    """
    minx, miny = float(origin_mm[0]), float(origin_mm[1])
    by = defaultdict(list)
    for v in sorted(gk_nodes):
        if not (0 <= v < len(board_pts)):
            continue
        p = board_pts[v]
        by[(snap_key(p[0], p[1], minx, miny) if grid
            else (round(float(p[0]), 6), round(float(p[1]), 6)))].append(v)
    folds = {k: vs for k, vs in by.items() if len(vs) > 1}
    # ★한 덩어리로 세면 안 된다. 완전히 **같은 좌표**의 중복 절점이 눌리는 것은
    #   합치는 게 맞고(도면이 선을 겹쳐 그린 자리), 그림 6 ② 가 말하는 훼손은
    #   «떨어져 있던 두 점» 이 한 칸으로 눌리는 것이다. 실측 대명동: 55개 중
    #   대부분이 같은 좌표였다 — 안 가르면 정상을 훼손으로 센다.
    same, apart = {}, {}
    for k, vs in folds.items():
        far = 0.0
        for i in range(len(vs)):
            for j in range(i + 1, len(vs)):
                a, b = board_pts[vs[i]], board_pts[vs[j]]
                far = max(far, math.hypot(float(a[0]) - float(b[0]),
                                          float(a[1]) - float(b[1])))
        (same if far <= 1.0 else apart)[k] = (vs, far)
    return {"gk": len(gk_nodes), "cells": len(by),
            "merged": sum(len(vs) - 1 for vs in folds.values()),
            "same": same, "apart": apart,
            "n_same": sum(len(v[0]) - 1 for v in same.values()),
            "n_apart": sum(len(v[0]) - 1 for v in apart.values())}


def measure_tree(kfp, nodes_of, pipes_of, ends_of):
    """[나무] |E| = |V| − 1 이고 성분이 하나인가."""
    nd, pr = nodes_of(kfp), pipes_of(kfp)
    adj = defaultdict(set)
    E = 0
    for p in pr.values():
        s, e = ends_of(p)
        if s is None or e is None or s == e:
            continue
        adj[s].add(e)
        adj[e].add(s)
        E += 1
    V = len(nd)
    seen, comp = set(), 0
    for n in nd:
        if n in seen:
            continue
        comp += 1
        q = deque([n])
        seen.add(n)
        while q:
            cur = q.popleft()
            for x in adj.get(cur, ()):
                if x not in seen:
                    seen.add(x)
                    q.append(x)
    return {"V": V, "E": E, "comp": comp, "tree": (E == V - 1 and comp == 1)}


def measure_far(worst, board_pts, kfp, root, nodes_of, pipes_of, ends_of):
    """[최원길이] 손질 far_m(board 거리) ↔ 표의 최원 경로 평면 길이 합."""
    wp = [int(n) for n in (worst.get("worst_path") or ())]
    board_m = 0.0
    for i in range(len(wp) - 1):
        a, b = board_pts[wp[i]], board_pts[wp[i + 1]]
        board_m += math.hypot(float(b[0]) - float(a[0]),
                              float(b[1]) - float(a[1])) / 1000.0
    # kfp 쪽: 뿌리 → 앵커 헤드까지 평면 배관 length_m 합
    nd, pr = nodes_of(kfp), pipes_of(kfp)
    adj = defaultdict(list)
    for pid, p in pr.items():
        s, e = ends_of(p)
        if s is None or e is None:
            continue
        adj[s].append((e, pid))
        adj[e].append((s, pid))
    return {"board_m": board_m, "far_m": float(worst.get("far_m") or 0.0),
            "adj": adj, "nodes": nd, "pipes": pr, "root": root}


def path_len(far_ctx, target, plan, only_plane=True, tol=1.0):
    """뿌리 → target 경로의 length_m 합 (평면 배관만 셀 수 있게)."""
    adj, nd, pr = far_ctx["adj"], far_ctx["nodes"], far_ctx["pipes"]
    root = far_ctx["root"]
    prev = {root: (None, None)}
    q = deque([root])
    while q:
        cur = q.popleft()
        if cur == target:
            break
        for nxt, pid in adj.get(cur, ()):
            if nxt in prev:
                continue
            prev[nxt] = (cur, pid)
            q.append(nxt)
    if target not in prev:
        return None
    tot, cur = 0.0, target
    while prev.get(cur, (None, None))[0] is not None:
        par, pid = prev[cur]
        p = pr.get(pid) or {}
        a = (nd.get(par) or {}).get("coords")
        b = (nd.get(cur) or {}).get("coords")
        horiz = True
        if a and b:
            horiz = math.hypot(plan.mm(a)[0] - plan.mm(b)[0],
                               plan.mm(a)[1] - plan.mm(b)[1]) > tol
        if (not only_plane) or horiz:
            tot += float(p.get("length_m") or 0.0)
        cur = par
    return tot


def measure_merge_actual(gk_nodes, board_pts, ps, origin_mm, tol=35.0):
    """[병합·직접] **만들어진 망에서** 센다 — 가정하지 않는다.

    ★두 번 틀렸다. 처음엔 격자 식으로만 세어 「같은 좌표 중복」까지 훼손으로
      셌고, 다음엔 스냅을 껐다고 **믿고** 정확 좌표로 세다가 되돌림이 걸린
      판에서 0 을 냈다. 자는 «결과» 를 봐야 한다.

    서로 1mm 넘게 떨어진 G_K 노드 둘이 **같은** kfp 노드로 가면 그것이 병합이다.
    """
    from collections import defaultdict as _dd
    hit = _dd(list)
    for v in sorted(gk_nodes):
        if not (0 <= v < len(board_pts)):
            continue
        p = board_pts[v]
        q = (float(p[0]), float(p[1]))
        nid, d = ps.near_node(q)
        if nid is None or d > tol:
            continue
        hit[nid].append(v)
    same, apart = 0, {}
    for nid, vs in hit.items():
        if len(vs) < 2:
            continue
        far = 0.0
        for i in range(len(vs)):
            for j in range(i + 1, len(vs)):
                a, b = board_pts[vs[i]], board_pts[vs[j]]
                far = max(far, math.hypot(float(a[0]) - float(b[0]),
                                          float(a[1]) - float(b[1])))
        if far <= 1.0:
            same += len(vs) - 1
        else:
            apart[nid] = (vs, far)
    return {"n_same": same, "apart": apart,
            "n_apart": sum(len(v[0]) - 1 for v in apart.values())}
