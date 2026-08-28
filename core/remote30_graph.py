# -*- coding: utf-8 -*-
"""remote30 순수 그래프/기하 원시함수 (Phase2 분할 — leaf 클러스터).

노드 인덱싱(_NodeIndex), 최단경로(_dijkstra_from/_shortest_path), 연결성분,
최근접노드, edge 방향/중점, 삼각형 판정, 점-선분 거리. 다른 top-level 참조 없음.
"""
from __future__ import annotations

import heapq
import math
from collections import defaultdict
from remote30_constants import SNAP_TOL_MM


def _point_to_segment_dist(px: float, py: float,
                            ax: float, ay: float,
                            bx: float, by: float) -> float:
    """점 (px,py) ↔ segment [a,b] 최단 거리."""
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / (dx * dx + dy * dy)))
    cx, cy = ax + t * dx, ay + t * dy
    return math.hypot(px - cx, py - cy)

def _round_pt(x: float, y: float, tol: float = SNAP_TOL_MM) -> tuple[float, float]:
    """격자 정렬 좌표 — HeadCandidate dedup, Counter 키 등 동등성 비교용.

    그래프 노드 키로는 더 이상 사용 안 함 (_NodeIndex 가 raw 좌표 기반 cluster 처리).
    """
    return (round(x / tol) * tol, round(y / tol) * tol)

class _NodeIndex:
    """Grid-bucket 기반 epsilon-tolerant endpoint cluster.

    Snap 격자(round-to-grid)의 대안. raw 좌표를 노드 키로 그대로 보존하면서,
    epsilon 반경 안에 기존 노드가 있으면 그 노드 좌표를 반환 (없으면 신규 등록).

    이점:
      - 노드 좌표 = raw → Stage 3 시각화가 격자 정렬 안 됨 → 비뚤어짐 없음
      - 격자 경계 분리 위험 없음 (cluster 가 9 bucket neighborhood 검색)
      - 같은 raw 좌표는 항상 같은 tuple value → dict hash 일관
    """

    __slots__ = ("eps", "_eps_sq", "_cell", "_bucket")

    def __init__(self, epsilon_mm: float = SNAP_TOL_MM):
        self.eps = epsilon_mm
        self._eps_sq = epsilon_mm * epsilon_mm
        self._cell = epsilon_mm
        self._bucket: dict[tuple[int, int], list[tuple[float, float]]] = defaultdict(list)

    def canonical(self, x: float, y: float) -> tuple[float, float]:
        """epsilon 안에 기존 노드 있으면 그 좌표, 없으면 (x, y) 신규 등록."""
        kx = int(x // self._cell)
        ky = int(y // self._cell)
        best = None
        bestd = float("inf")
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for pt in self._bucket.get((kx + dx, ky + dy), ()):
                    d = (pt[0] - x) ** 2 + (pt[1] - y) ** 2
                    if d < bestd and d <= self._eps_sq:
                        bestd = d
                        best = pt
        if best is not None:
            return best
        new_pt = (x, y)
        self._bucket[(kx, ky)].append(new_pt)
        return new_pt

def _is_triangle_shape(pts: list, tol: float = 2.0) -> bool:
    """HATCH path 의 점 시퀀스가 삼각형 (3 고유 정점) 인지 — closed loop 의 시작/끝 중복 무시."""
    if not pts or len(pts) < 3:
        return False
    unique: list[tuple[float, float]] = []
    for p in pts:
        x, y = float(p[0]), float(p[1])
        if not any(abs(x - u[0]) < tol and abs(y - u[1]) < tol for u in unique):
            unique.append((x, y))
    return len(unique) == 3

def _edge_dir(p: tuple, q: tuple) -> tuple[float, float]:
    dx, dy = q[0] - p[0], q[1] - p[1]
    n = math.hypot(dx, dy)
    if n == 0.0:
        return (0.0, 0.0)
    return (dx / n, dy / n)

def _midpoint(p: tuple, q: tuple) -> tuple[float, float]:
    return ((p[0] + q[0]) / 2, (p[1] + q[1]) / 2)

def _dijkstra_from(graph: dict, edge_len: dict, src: tuple[float, float],
                   prev_out: dict | None = None) -> dict[tuple[float, float], float]:
    """단순 Dijkstra — 모든 노드까지의 거리.

    prev_out: 주면 «최단경로 나무»(각 노드의 직전 노드)를 여기 채운다. 그러면
        같은 src 에서 나가는 경로를 노드마다 다시 풀 필요가 없다 —
        `_path_from_prev` 로 되돌아 걷기만 하면 된다. 안 주면 예전과 똑같이
        거리만 돌려주므로 기존 호출자는 영향이 없다.

        이 인자가 없던 시절, 호출자는 거리 맵을 한 번 만든 뒤 헤드마다
        `_shortest_path` 로 Dijkstra 를 **또** 돌렸다(B1F 실측: 2,206회 ·
        40.5초 · 배관망 검출 전체의 절반). 같은 나무를 2,206번 다시 세운 셈이다.
    """
    dist: dict[tuple[float, float], float] = {src: 0.0}
    pq: list[tuple[float, tuple[float, float]]] = [(0.0, src)]
    while pq:
        d, u = heapq.heappop(pq)
        if d > dist.get(u, float("inf")):
            continue
        for v in graph.get(u, ()):
            key = (min(u, v), max(u, v))
            w = edge_len.get(key, math.hypot(v[0] - u[0], v[1] - u[1]))
            nd = d + w
            if nd < dist.get(v, float("inf")):
                dist[v] = nd
                if prev_out is not None:
                    prev_out[v] = u
                heapq.heappush(pq, (nd, v))
    return dist


def _path_from_prev(prev: dict, src: tuple[float, float],
                    tgt: tuple[float, float]) -> list[tuple[float, float]]:
    """`_dijkstra_from(prev_out=…)` 이 채운 나무에서 src → tgt 경로를 되찾는다.

    반환 규약은 `_shortest_path` 와 똑같다 — 같은 자리에 그대로 끼울 수 있게.
    닿지 않으면 빈 목록, src==tgt 면 [src].
    """
    if src == tgt:
        return [src]
    if tgt not in prev:
        return []
    out = [tgt]
    while out[-1] in prev:
        out.append(prev[out[-1]])
    out.reverse()
    return out if out and out[0] == src else []

def _shortest_path(graph: dict, edge_len: dict, src: tuple[float, float], tgt: tuple[float, float],
                   penalty_keys: set | None = None, penalty_mm: float = 1.0e9) -> list[tuple[float, float]]:
    """src → tgt 최단 경로 (vertex 시퀀스).

    penalty_keys: 거대 가중치를 더할 edge 키 집합 (rounded-int 좌표쌍 (min,max)).
        force_connect 가 만든 추정 bridge(도면에 없는 직선 wormhole)를 여기에 넣으면
        Dijkstra 가 실측 배관 경로를 우선하고, 추정 bridge 는 다른 대안이 전혀 없을 때만
        (최소 개수로) 사용한다 → 도면을 가로지르는 "엉뚱한 경로" 방지. 단, 작은 gap 을
        메우는 단계 bridge(≤tolerance)는 실배관 연속으로 보고 penalty 대상에서 제외.
    """
    if src == tgt:
        return [src]
    penalty_keys = penalty_keys or set()
    dist = {src: 0.0}
    prev: dict = {}
    pq = [(0.0, src)]
    while pq:
        d, u = heapq.heappop(pq)
        if u == tgt:
            break
        if d > dist.get(u, float("inf")):
            continue
        for v in graph.get(u, ()):
            key = (min(u, v), max(u, v))
            w = edge_len.get(key, math.hypot(v[0] - u[0], v[1] - u[1]))
            if penalty_keys:
                ru = (int(round(u[0])), int(round(u[1])))
                rv = (int(round(v[0])), int(round(v[1])))
                if (min(ru, rv), max(ru, rv)) in penalty_keys:
                    w += penalty_mm
            nd = d + w
            if nd < dist.get(v, float("inf")):
                dist[v] = nd
                prev[v] = u
                heapq.heappush(pq, (nd, v))
    if tgt not in prev and tgt != src:
        return []
    # backtrack
    out = [tgt]
    while out[-1] in prev:
        out.append(prev[out[-1]])
    out.reverse()
    return out if out and out[0] == src else []

def _nearest_graph_node(graph: dict, pt: tuple[float, float]) -> tuple[float, float] | None:
    """그래프 노드 중 pt 와 가장 가까운 노드. 같은 좌표면 그대로."""
    if pt in graph:
        return pt
    best = None
    bestd = float("inf")
    for n in graph:
        d = (n[0] - pt[0]) ** 2 + (n[1] - pt[1]) ** 2
        if d < bestd:
            bestd = d
            best = n
    return best

class _NearestNodeIndex:
    """`_nearest_graph_node` 를 여러 번 물을 때 쓰는 격자 색인.

    같은 그래프에 헤드마다 물으면 선형 탐색이 헤드 × 노드로 든다 — B1F 실측
    9,810회 · 10.8초(배관망 검출의 26%). 격자에 한 번 담아 두고 가까운 칸부터
    넓혀 가며 찾으면 그 값이 사라진다.

    ★답은 선형 탐색과 «똑같다». 거리뿐 아니라 **무승부 규칙**까지 맞춘다 —
      원본은 `for n in graph` 를 돌며 `d < bestd` 일 때만 갈아치우므로, 같은
      거리면 먼저 나온(먼저 삽입된) 노드가 이긴다. 그래서 후보를 모은 뒤
      `(거리, 삽입순번)` 으로 고른다. 여기를 대충 맞추면 «더 빠른데 답이 다른»
      것이 되어 최적화가 아니다.
    """

    __slots__ = ("_rank", "_cell", "_grid", "_empty", "_lo", "_hi")

    def __init__(self, graph: dict) -> None:
        nodes = list(graph)
        self._rank = {n: i for i, n in enumerate(nodes)}
        self._empty = not nodes
        if self._empty:
            self._cell, self._grid = 1.0, {}
            self._lo = self._hi = (0, 0)
            return
        xs = [n[0] for n in nodes]
        ys = [n[1] for n in nodes]
        w = max(xs) - min(xs)
        h = max(ys) - min(ys)
        # 노드당 대략 한 칸. 한 점에 몰려 있으면(w=h=0) 칸 하나로 떨어진다.
        span = max(w, h)
        self._cell = max(span / max(1.0, len(nodes) ** 0.5), 1e-9)
        grid: dict = {}
        for n in nodes:
            grid.setdefault((int(n[0] // self._cell), int(n[1] // self._cell)),
                            []).append(n)
        self._grid = grid
        ks = grid.keys()
        self._lo = (min(k[0] for k in ks), min(k[1] for k in ks))
        self._hi = (max(k[0] for k in ks), max(k[1] for k in ks))

    def _scan_all(self, pt):
        """선형 탐색 — 원본과 글자 그대로 같은 규칙(먼저 삽입된 노드가 이김)."""
        best, bestd = None, float("inf")
        for n in self._rank:
            d = (n[0] - pt[0]) ** 2 + (n[1] - pt[1]) ** 2
            if d < bestd:
                bestd, best = d, n
        return best

    def nearest(self, pt: tuple[float, float]) -> tuple[float, float] | None:
        if pt in self._rank:            # 원본의 이른 반환과 같다
            return pt
        if self._empty:
            return None
        cell = self._cell
        gx, gy = int(pt[0] // cell), int(pt[1] // cell)
        lo, hi = self._lo, self._hi
        # ★구름 «밖» 의 점은 격자로 찾으면 안 된다. 테를 구름에 닿을 때까지
        #   넓혀야 하는데, 멀리 있는 점은 그 횟수가 노드 수와 무관하게 커진다
        #   (실측: LH306 42헤드 · 절점 201 에서 90초 — 선형 탐색은 179ms).
        #   밖이면 한 번 훑는 편이 언제나 싸다. 답은 어차피 같다.
        if not (lo[0] <= gx <= hi[0] and lo[1] <= gy <= hi[1]):
            return self._scan_all(pt)
        # 안쪽이면 테를 넓히며 찾는다. 테 «둘레» 만 훑는다 — 정사각형을 통째로
        # 돌면 테마다 O(r²) 이 되어 넓힐수록 제곱으로 는다.
        span = max(hi[0] - lo[0], hi[1] - lo[1]) + 1
        best, bestd = None, float("inf")
        r = 0
        while True:
            if r == 0:
                cells = ((gx, gy),)
            else:
                cells = []
                for i in range(gx - r, gx + r + 1):
                    cells.append((i, gy - r))
                    cells.append((i, gy + r))
                for j in range(gy - r + 1, gy + r):
                    cells.append((gx - r, j))
                    cells.append((gx + r, j))
            for key in cells:
                for n in self._grid.get(key, ()):
                    d = (n[0] - pt[0]) ** 2 + (n[1] - pt[1]) ** 2
                    if d < bestd or (d == bestd and best is not None
                                     and self._rank[n] < self._rank[best]):
                        bestd, best = d, n
            # 찾았어도 한 테 더 봐야 한다 — 대각선 이웃 칸이 더 가까울 수 있다.
            if best is not None and (r * cell) ** 2 >= bestd:
                return best
            r += 1
            if r > span:                # 격자를 다 훑었다 — 남은 것은 선형뿐
                return best if best is not None else self._scan_all(pt)


def _connected_components(graph: dict) -> list[set]:
    """그래프의 connected component 들."""
    seen = set()
    comps = []
    for start in graph:
        if start in seen:
            continue
        stack = [start]
        comp = set()
        while stack:
            u = stack.pop()
            if u in seen:
                continue
            seen.add(u); comp.add(u)
            for v in graph.get(u, ()):
                if v not in seen:
                    stack.append(v)
        comps.append(comp)
    return comps


class HeadRegion:
    """헤드 영역 표현 통일(W4) — rect union 또는 임의 다각형 + 팽창 margin.

    # [문서정합] 작업지시서 W4는 "shapely Polygon 기반 — extractor가 이미 shapely
    # 의존"이라 명시하나 실제로는 shapely 미설치·저장소 무의존 (BLOCKED.md #1).
    # 동일 API(from_rects/from_polygon/contains/dilate)의 순수 파이썬 구현.

    - rect 는 (x1,y1,x2,y2). 생성 시 min/max 정규화 — 기존 branch_zones in_region
      판정(경계 포함 <=)과 비트동일.
    - dilate(mm) 는 margin 누적한 새 인스턴스 반환 (rect: 변 팽창, polygon:
      경계까지 거리 <= margin 판정 — Minkowski 원판 팽창).
    - pts: 정점 리스트 노출 (_AnchorWindow convex hull 입력용).
    """

    __slots__ = ("_rects", "_poly", "_margin")

    def __init__(self, rects=None, poly=None, margin_mm: float = 0.0):
        self._rects = [
            (min(float(x1), float(x2)), min(float(y1), float(y2)),
             max(float(x1), float(x2)), max(float(y1), float(y2)))
            for (x1, y1, x2, y2) in (rects or [])
        ]
        self._poly = [(float(x), float(y)) for x, y in (poly or [])]
        self._margin = float(margin_mm)

    @classmethod
    def from_rects(cls, rects: list) -> "HeadRegion":
        """사각형 union 승격 — rect = (x1,y1,x2,y2), 순서 무관."""
        return cls(rects=rects)

    @classmethod
    def from_polygon(cls, pts: list) -> "HeadRegion":
        """임의 다각형(L자형 등) — 정점 [(x,y), ...]."""
        return cls(poly=pts)

    @property
    def pts(self) -> list:
        """정점 리스트 — convex hull(작업창 W) 입력용."""
        if self._poly:
            return list(self._poly)
        out = []
        for (x1, y1, x2, y2) in self._rects:
            out += [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
        return out

    def __bool__(self) -> bool:
        return bool(self._rects or self._poly)

    def dilate(self, mm: float) -> "HeadRegion":
        return HeadRegion(rects=self._rects, poly=self._poly,
                          margin_mm=self._margin + float(mm))

    def _poly_inside(self, x: float, y: float) -> bool:
        p = self._poly
        inside = False
        for i in range(len(p)):
            x1, y1 = p[i]
            x2, y2 = p[(i + 1) % len(p)]
            if (y1 > y) != (y2 > y):
                xin = (x2 - x1) * (y - y1) / (y2 - y1) + x1
                if x < xin:
                    inside = not inside
        return inside

    def contains(self, pt) -> bool:
        x, y = float(pt[0]), float(pt[1])
        m = self._margin
        for (x1, y1, x2, y2) in self._rects:
            if x1 - m <= x <= x2 + m and y1 - m <= y <= y2 + m:
                return True
        if len(self._poly) >= 3:
            if self._poly_inside(x, y):
                return True
            if m > 0:
                p = self._poly
                return any(
                    _point_to_segment_dist(x, y, p[i][0], p[i][1],
                                           p[(i + 1) % len(p)][0],
                                           p[(i + 1) % len(p)][1]) <= m
                    for i in range(len(p))
                )
        return False
