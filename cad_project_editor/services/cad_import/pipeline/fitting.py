# -*- coding: utf-8 -*-
"""재료 레이어 작은 원호 부속 자리 잇기. 화면 없음.

접속표시(`pipeline.flow.join_all`/`spot_arms`)와 다른 함수다.
"""
import math
from collections import Counter, defaultdict

from services.cad_import.pipeline.expand import gnear, gput

SMALL_R = 300.0     # 「작은 원호」 문턱 — 본체와 같은 값
ARM_CTR = 5.0       # 팔이 «중심»에 앉았다고 볼 거리
ARM_RIM = 10.0      # 팔이 «원호 위»에 앉았다고 볼 오차
ON_LINE = 5.0       # 남은 팔이 이은 선 «위»에 얹혔다고 볼 거리


def fitting_spots(w, mat_layers):
    """부속 원호 자리 = 재료 레이어의 작은 원호. 중심·반지름이 같으면 한 자리."""
    spots = {}
    for ly, _c, cx, cy, r in w.arcs:
        if ly in mat_layers and 0 < r <= SMALL_R:
            spots.setdefault((round(cx, 1), round(cy, 1), round(r, 1)),
                             (cx, cy, r))
    return list(spots.values())


def join_at_fittings(g, spots):
    """오너 규칙 ①~④. 부속 자리마다 모인 팔을 전부 한 점에 묶는다."""
    # 자유단(차수 1) 뿐 아니라 «모든 노드»를 팔 후보로 본다 — 도면이 팔을
    # 원호 중심까지 밀어 넣은 자리는 이미 다른 관에 얹혀 있을 수 있다.
    ngrid = defaultdict(list)
    for i, (x, y) in enumerate(g.pts):
        gput(ngrid, 400.0, x, y, i)
    adj = g.adj()

    stat = Counter()
    n_join = 0
    for (cx, cy, r) in spots:
        # ---- ① 팔 모으기: 중심(≈0) 또는 원호 위(≈r)에 앉은 노드 ----
        arms = []
        for n in set(gnear(ngrid, 400.0, cx, cy, rings=1)):
            x, y = g.pts[n]
            d = math.hypot(x - cx, y - cy)
            if not (d <= ARM_CTR or abs(d - r) <= ARM_RIM):
                continue
            # 팔의 «뻗는 방향» = 이 노드에 붙은 관이 향하는 쪽
            dirs = []
            for m in adj.get(n, ()):
                vx, vy = g.pts[m][0] - x, g.pts[m][1] - y
                ln = math.hypot(vx, vy)
                if ln > 1.0:
                    dirs.append((vx / ln, vy / ln))
            if dirs:
                arms.append((d, n, dirs))
        stat[f"팔{len(arms)}개"] += 1
        if len(arms) < 2:
            continue

        # ---- ② 마주 보는 팔 둘 찾기 (일직선) ----
        best = None
        for p in range(len(arms)):
            for q in range(p + 1, len(arms)):
                _d1, n1, ds1 = arms[p]
                _d2, n2, ds2 = arms[q]
                if n1 == n2:
                    continue
                for v1 in ds1:
                    for v2 in ds2:
                        if v1[0] * v2[0] + v1[1] * v2[1] > -0.98:
                            continue
                        p1, p2 = g.pts[n1], g.pts[n2]
                        wx, wy = p1[0] - p2[0], p1[1] - p2[1]
                        gap = math.hypot(wx, wy)
                        if gap < 1.0:
                            continue
                        if abs(wx * v2[1] - wy * v2[0]) > ARM_RIM:
                            continue            # 일직선이 아니다
                        if best is None or gap < best[0]:
                            best = (gap, n1, n2)
        if best is None:
            # ---- ④ 엘보 — 팔이 둘뿐이면 방향 안 보고 잇는다 ----
            if len(arms) == 2 and arms[0][1] != arms[1][1]:
                g.add_edge(arms[0][1], arms[1][1])
                n_join += 1
                stat["엘보이음"] += 1
            else:
                stat["짝못찾음"] += 1
            continue

        gap, n1, n2 = best
        p1, p2 = g.pts[n1], g.pts[n2]
        ux, uy = (p1[0] - p2[0]) / gap, (p1[1] - p2[1]) / gap

        # ---- ③ 이은 선 위에 얹힌 나머지 팔을 사슬로 꿴다 ----
        chain = [(0.0, n2), (gap, n1)]
        for _d, n, _ds in arms:
            if n in (n1, n2):
                continue
            x, y = g.pts[n]
            wx, wy = x - p2[0], y - p2[1]
            perp = abs(wx * uy - wy * ux)
            along = wx * ux + wy * uy
            if perp <= ON_LINE and -ARM_RIM <= along <= gap + ARM_RIM:
                chain.append((along, n))
                stat["선위에얹힌팔"] += 1
        chain.sort()
        seen = set()
        order = [n for _a, n in chain if not (n in seen or seen.add(n))]
        for a_, b_ in zip(order, order[1:]):
            if (min(a_, b_), max(a_, b_)) not in g.edges:
                g.add_edge(a_, b_)
                n_join += 1
        stat[f"사슬{len(order)}노드"] += 1
    return n_join, stat
