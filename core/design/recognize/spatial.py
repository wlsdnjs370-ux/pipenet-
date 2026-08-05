# -*- coding: utf-8 -*-
"""인식 셸이 공유하는 공간 인덱스.

[문서정합] 지시서 §2 의 파일 표에는 없다. C130 평행쌍·C150 중심선 쌍·C160 끝점
후보·C170 교차 분할이 전부 같은 인덱스를 필요로 하는데, 네 곳에 각각 두면 공차
규약이 서로 어긋난 채 조용히 갈라진다.

끝점 병합은 **격자 반올림이 아니라 epsilon 군집**이다. 격자로 반올림하면 셀
경계를 사이에 둔 두 점이 공차 안인데도 서로 다른 노드가 된다. 격자는 후보를
좁히는 데만 쓰고 판정은 실제 거리로 한다. 좌표는 원본을 그대로 보존한다.
"""
from __future__ import annotations

import math

Point = tuple[float, float]
Segment = tuple[float, float, float, float]


class NodeIndex:
    """공차 안의 점을 같은 노드로 묶는다. 좌표는 처음 들어온 값을 유지한다.

    점마다 공차가 다를 수 있다 — C150 의 `snap_tol` 은 벽 두께에 걸려 있다
    (§3.3). 그래서 `tol` 은 격자 크기이자 **상한**이고, 실제 병합 판정은 두 점의
    공차 중 작은 쪽으로 한다. 작은 쪽을 쓰지 않으면 어느 점을 먼저 넣었느냐에
    따라 결과가 달라진다.
    """

    def __init__(self, tol: float):
        if tol <= 0:
            raise ValueError("공차는 양수여야 합니다")
        self.tol = float(tol)
        self.points: list[Point] = []
        self.tols: list[float] = []
        self._cells: dict[tuple[int, int], list[int]] = {}

    def _key(self, x: float, y: float) -> tuple[int, int]:
        return (math.floor(x / self.tol), math.floor(y / self.tol))

    def find(self, x: float, y: float, tol: float | None = None) -> int | None:
        limit = self.tol if tol is None else min(float(tol), self.tol)
        cx, cy = self._key(x, y)
        best: int | None = None
        best_d = limit
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for nid in self._cells.get((cx + dx, cy + dy), ()):
                    px, py = self.points[nid]
                    d = math.hypot(px - x, py - y)
                    if d <= min(best_d, self.tols[nid]):
                        best, best_d = nid, d
        return best

    def add(self, x: float, y: float, tol: float | None = None) -> int:
        nid = self.find(x, y, tol)
        if nid is not None:
            return nid
        nid = len(self.points)
        self.points.append((x, y))
        self.tols.append(self.tol if tol is None else min(float(tol), self.tol))
        self._cells.setdefault(self._key(x, y), []).append(nid)
        return nid

    def __len__(self) -> int:
        return len(self.points)


class SegmentGrid:
    """선분이 **지나는 셀 전부**에 등록하는 격자.

    중점만 넣으면 나란히 놓인 두 장선이 서로 다른 셀에 떨어져 평행쌍을 놓친다.
    조회는 3x3 이웃까지 훑는다 — 셀 대각을 스쳐 지나는 경우가 있다.

    셀 계산(`walk`)과 등록/조회를 나눠 둔 이유는, 호출자가 같은 선분을 각도별로
    나눈 여러 격자에 물어보기 때문이다. 매번 다시 걸으면 그게 비용이 된다.
    """

    def __init__(self, cell: float):
        if cell <= 0:
            raise ValueError("셀 크기는 양수여야 합니다")
        self.cell = float(cell)
        self._cells: dict[tuple[int, int], list[int]] = {}

    def walk(self, seg: Segment) -> tuple[tuple[int, int], ...]:
        x1, y1, x2, y2 = seg
        cell = self.cell
        # 표본 간격을 셀의 절반으로 잡아야 셀을 건너뛰지 않는다.
        steps = int(math.hypot(x2 - x1, y2 - y1) / (cell * 0.5)) + 1
        out: set[tuple[int, int]] = set()
        for i in range(steps + 1):
            t = i / steps
            out.add((math.floor((x1 + (x2 - x1) * t) / cell),
                     math.floor((y1 + (y2 - y1) * t) / cell)))
        return tuple(out)

    def add(self, idx: int, cells) -> None:
        for key in cells:
            self._cells.setdefault(key, []).append(idx)

    def lookup(self, cells) -> set[int]:
        out: set[int] = set()
        get = self._cells.get
        for cx, cy in cells:
            for dx in (-1, 0, 1):
                for dy in (-1, 0, 1):
                    out.update(get((cx + dx, cy + dy), ()))
        return out


def angle_bucketed_pairs(segments, *, bucket_deg: float, cell: float):
    """근접한 평행 후보 쌍만 흘려보낸다. `(i, j, angle_i, angle_j)`, `i < j`.

    격자를 각도 버킷별로 따로 둔다. 하나로 합치면 가구 레이어처럼 짧은 선이
    수만 개 몰린 셀에서 후보가 그대로 O(n^2) 로 불어나 실 도면에서만 멈춘다.
    각도로 먼저 가르면 대부분의 후보가 조회 전에 사라진다.

    C130 평행쌍(§3.1)과 C150 중심선 쌍(§3.3)이 버킷 폭만 다르고 같은 탐색을
    한다. 각자 두면 한쪽만 고쳐져 조용히 갈라진다.
    """
    n = len(segments)
    if n < 2:
        return
    angles = [angle_deg(s) for s in segments]
    n_buckets = max(1, int(round(180.0 / bucket_deg)))

    grids: dict[int, SegmentGrid] = {}
    cells: list[tuple] = []
    buckets: list[int] = []
    for i, seg in enumerate(segments):
        b = int(angles[i] / bucket_deg) % n_buckets
        grid = grids.get(b)
        if grid is None:
            grid = grids[b] = SegmentGrid(cell)
        walked = grid.walk(seg)
        grid.add(i, walked)
        cells.append(walked)
        buckets.append(b)

    for i in range(n):
        # 이웃 버킷까지 봐야 179.9° 와 0.1° 처럼 버킷 경계를 사이에 둔 짝을 놓치지 않는다.
        candidates: set[int] = set()
        for d in (-1, 0, 1):
            grid = grids.get((buckets[i] + d) % n_buckets)
            if grid is not None:
                candidates |= grid.lookup(cells[i])
        for j in candidates:
            if j > i:
                yield i, j, angles[i], angles[j]


def angle_deg(seg: Segment) -> float:
    """0 이상 180 미만. 선분에는 방향이 없으므로 180 으로 접는다."""
    x1, y1, x2, y2 = seg
    return math.degrees(math.atan2(y2 - y1, x2 - x1)) % 180.0


def angle_diff(a: float, b: float) -> float:
    d = abs(a - b) % 180.0
    return min(d, 180.0 - d)


def length(seg: Segment) -> float:
    x1, y1, x2, y2 = seg
    return math.hypot(x2 - x1, y2 - y1)


def perp_offset(seg: Segment, px: float, py: float) -> float:
    """`seg` 를 무한직선으로 볼 때 점까지의 수직거리."""
    x1, y1, x2, y2 = seg
    dx, dy = x2 - x1, y2 - y1
    norm = math.hypot(dx, dy)
    if norm == 0.0:
        return math.hypot(px - x1, py - y1)
    return abs((px - x1) * dy - (py - y1) * dx) / norm


def overlap_ratio(a: Segment, b: Segment) -> float:
    """`a` 방향으로 투영했을 때 겹치는 길이 / 짧은 쪽 길이."""
    ax1, ay1, ax2, ay2 = a
    dx, dy = ax2 - ax1, ay2 - ay1
    la = math.hypot(dx, dy)
    lb = length(b)
    if la == 0.0 or lb == 0.0:
        return 0.0
    ux, uy = dx / la, dy / la
    t1 = (b[0] - ax1) * ux + (b[1] - ay1) * uy
    t2 = (b[2] - ax1) * ux + (b[3] - ay1) * uy
    lo, hi = (t1, t2) if t1 <= t2 else (t2, t1)
    return max(0.0, min(la, hi) - max(0.0, lo)) / min(la, lb)


def polygon_area(points: list) -> float:
    """부호 없는 면적. 좌표 단위의 제곱으로 돌려준다."""
    n = len(points)
    if n < 3:
        return 0.0
    total = 0.0
    for i in range(n):
        x1, y1 = points[i][0], points[i][1]
        x2, y2 = points[(i + 1) % n][0], points[(i + 1) % n][1]
        total += x1 * y2 - x2 * y1
    return abs(total) * 0.5


def centroid(points: list) -> Point:
    return (sum(p[0] for p in points) / len(points),
            sum(p[1] for p in points) / len(points))
