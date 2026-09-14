# -*- coding: utf-8 -*-
"""[아이소·평면보존 §2] P1·P2 판정 — «보존» 을 **기하 집합**으로 잰다.

★자를 한 번 틀렸다. 처음에는 「G_K 노드마다 같은 자리의 G′ **노드**가 있나」로
  쟀는데, 그러면 `build_planar_graph` 의 **일직선 중간 노드 병합**이 전부
  「사라짐」으로 세어진다(실측 대명동: 520 → 286 노드, 끊김 300/519).
  지시서 §1-3 P2 는 그 병합을 **허용**한다 — 「세분·병합만 허용」.

  그래서 판정은 이렇게 읽는다:

      보존 = p(v) 가 π(G′) 라는 **점·선분의 집합** 위에 그대로 있다

  ①노드로 남음  ②간선 위에 남음(중간 노드가 병합된 것 — 허용)
  ③어디에도 없음 = 진짜 사라짐 (d)

  간선도 같다 — 「양 끝만 맞나」가 아니라 **선분 전체가 π(G′) 에 덮이나** 를
  잰다. 덮이면 몇 조각으로 쪼개졌든 상관없다(세분 허용).
"""
from __future__ import annotations

import math
from collections import defaultdict

TOL_MM = 5.0            # 같은 자리로 볼 거리 (GRID_M 스냅은 25mm 까지 나온다)
CELL = 1000.0


def _d(a, b):
    return math.hypot(float(a[0]) - float(b[0]), float(a[1]) - float(b[1]))


def _pt_seg(p, a, b):
    """점 → 선분 거리."""
    vx, vy = b[0] - a[0], b[1] - a[1]
    L2 = vx * vx + vy * vy
    if L2 < 1e-12:
        return _d(p, a)
    t = ((p[0] - a[0]) * vx + (p[1] - a[1]) * vy) / L2
    t = max(0.0, min(1.0, t))
    return math.hypot(p[0] - (a[0] + vx * t), p[1] - (a[1] + vy * t))


class PlanSet:
    """π(G′) — 평면(mm) 으로 내린 G′ 의 점·선분 집합. 격자로 찾는다."""

    def __init__(self, kfp, plan, nodes_of, pipes_of, ends_of):
        nd, pr = nodes_of(kfp), pipes_of(kfp)
        self.pt = {}
        for nid, m in nd.items():
            c = (m or {}).get("coords")
            if c:
                self.pt[nid] = plan.mm(c)
        self.seg = []
        for pid, p in pr.items():
            s, e = ends_of(p)
            if s in self.pt and e in self.pt:
                self.seg.append((self.pt[s], self.pt[e], pid))
        self._gp = defaultdict(list)
        for nid, q in self.pt.items():
            self._gp[(int(q[0] // CELL), int(q[1] // CELL))].append((nid, q))
        self._gs = defaultdict(list)
        for idx, (a, b, _pid) in enumerate(self.seg):
            n = max(1, int(_d(a, b) / CELL) + 1)
            for i in range(n + 1):
                t = i / n
                x, y = a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t
                self._gs[(int(x // CELL), int(y // CELL))].append(idx)

    def near_node(self, q, rings=2):
        gx, gy = int(q[0] // CELL), int(q[1] // CELL)
        best = (None, float("inf"))
        for dx in range(-rings, rings + 1):
            for dy in range(-rings, rings + 1):
                for nid, p in self._gp.get((gx + dx, gy + dy), ()):
                    dd = _d(p, q)
                    if dd < best[1]:
                        best = (nid, dd)
        return best

    def near_seg(self, q, rings=2):
        gx, gy = int(q[0] // CELL), int(q[1] // CELL)
        best = float("inf")
        seen = set()
        for dx in range(-rings, rings + 1):
            for dy in range(-rings, rings + 1):
                for idx in self._gs.get((gx + dx, gy + dy), ()):
                    if idx in seen:
                        continue
                    seen.add(idx)
                    a, b, _p = self.seg[idx]
                    dd = _pt_seg(q, a, b)
                    if dd < best:
                        best = dd
        return best


GRID_HALF_MM = 35.0     # GRID_M=0.05 의 반 칸(25mm) + 대각 여유


def check_P1(gk_nodes, board_pts, ps, tol=TOL_MM):
    """G_K 노드가 π(G′) 위에 그대로 있나.

    ★네 갈래로 가른다. 「있다/없다」로 가르면 **격자 스냅으로 몇 mm 옮겨진 것**
      이 「사라졌다」로 세어진다(실측 대명동: 그렇게 세면 364개가 사라진 것으로
      나오는데, 실제 거리는 5~32mm 였다 — GRID_M 반 칸이다).

        노드로 남음   ≤tol 에 G′ **노드**가 있다
        간선 위로     ≤tol 에 G′ **선분**이 있다 (중간 노드 병합 — 허용)
        격자 스냅     tol < d ≤ 35mm — §6-1 «오너가 정할 자리»
        사라짐        d > 35mm — 진짜 훼손 (d)
    """
    on_node, merged, snapped, lost = [], [], [], []
    for v in sorted(gk_nodes):
        if not (0 <= v < len(board_pts)):
            continue
        q = (float(board_pts[v][0]), float(board_pts[v][1]))
        _nid, dn = ps.near_node(q)
        ds = ps.near_seg(q)
        d = min(dn, ds)
        if dn <= tol:
            on_node.append(dn)
        elif ds <= tol:
            merged.append((v, q, dn, ds))     # 중간 노드 병합 — 허용
        elif d <= GRID_HALF_MM:
            snapped.append((v, q, round(d, 1)))
        else:
            lost.append((v, q, round(d, 1)))
    errs = ([0.0] * len(on_node) + [0.0] * len(merged)
            + [m[2] for m in snapped] + [m[2] for m in lost])
    cuts = (5.0, 50.0, 150.0)
    return {"n": len(gk_nodes), "on_node": len(on_node),
            "merged": merged, "snapped": snapped, "lost": lost,
            "max": max(errs) if errs else 0.0,
            "over": {c: sum(1 for e in errs if e > c) for c in cuts}}


def check_P2(gk_edges, board_pts, ps, tol=TOL_MM, samples=12):
    """G_K 간선이 π(G′) 에 **덮이나** — 몇 조각으로 쪼개졌든 상관없다."""
    broken, strayed = [], []
    devs = []
    for (u, v) in sorted({(min(a, b), max(a, b)) for a, b in gk_edges}):
        if not (0 <= u < len(board_pts) and 0 <= v < len(board_pts)):
            continue
        pu = (float(board_pts[u][0]), float(board_pts[u][1]))
        pv = (float(board_pts[v][0]), float(board_pts[v][1]))
        worst = 0.0
        for i in range(samples + 1):
            t = i / samples
            q = (pu[0] + (pv[0] - pu[0]) * t, pu[1] + (pv[1] - pu[1]) * t)
            worst = max(worst, ps.near_seg(q))
        devs.append(worst)
        if worst > 150.0:
            broken.append((u, v, round(worst, 1)))
        elif worst > tol:
            strayed.append((u, v, round(worst, 1)))
    return {"n": len(devs), "kept": len(devs) - len(broken),
            "broken": broken, "strayed": strayed,
            "max": max(devs) if devs else 0.0}


def check_P3(flat, raised, node_ref, nodes_of):
    """G′ 노드를 셋으로 — G_K 대응 · T_h(세로가 새로 만든 것) · 미분류."""
    f = set(nodes_of(flat))
    r = set(nodes_of(raised))
    ref = set(node_ref or {})
    plain = r & f
    return {"total": len(r), "gk": len(plain & ref), "th": len(r - f),
            "unknown": sorted(plain - ref)}
