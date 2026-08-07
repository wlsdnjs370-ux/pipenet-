# -*- coding: utf-8 -*-
"""모듈 D — D4 라벨 배치.

값이 맞아도 숫자가 겹쳐 읽히지 않으면 도면은 쓸모가 없다. 라벨마다 후보 자리를
가까운 곳부터 훑어 비어 있는 첫 자리에 놓고, 멀리 밀려난 것에는 지시선을 붙여
어느 관로의 값인지 남긴다.

전역 최적화는 하지 않는다. 라벨 하나가 움직여 온 도면이 재배치되면 시각 회귀
비교가 성립하지 않고, 지시서가 준 파일당 3 초 예산도 지킬 수 없다.
"""
from __future__ import annotations

import math
import time
from dataclasses import dataclass
from typing import Callable, Iterable, Sequence

Point = tuple[float, float]
Extent = tuple[float, float, float, float]  # x0, y0, x1, y1

# 상자 높이의 배수로 적는다 — 글자 크기가 바뀌어도 같은 여백감이 유지된다.
GAP = 0.35          # 기준점에서 상자까지 최소 틈
RING_STEP = 0.9     # 한 고리 밀려날 때마다 늘어나는 거리
MAX_RINGS = 26      # 이보다 멀면 어느 관로의 값인지 알아볼 수 없다
LEADER_AFTER = 1.4  # 이보다 멀리 밀면 지시선을 붙인다
# 잰 상자를 이만큼 부풀려 잡는다. 글자 잉크와 글꼴 한 줄 상자는 세로 기준이 조금
# 어긋나 있어서, 그 오차를 미리 얹어 두지 않으면 상자가 안 닿아도 글자는 닿는다.
PAD = 0.30

# 12 방위. 첫 후보(관 직각·노드 우상)에서 막히면 이 순서로 돈다.
_RING_DIRS: tuple[Point, ...] = tuple(
    (math.cos(math.radians(a)), math.sin(math.radians(a))) for a in range(0, 360, 30))

_SQ = math.sqrt(0.5)
_FIRST_DIRS: dict[str, tuple[Point, ...]] = {
    "node": ((_SQ, _SQ), (1.0, 0.0)),
    # 노즐 값은 방출단 아래에 붙인다 — 위쪽은 노즐이 매달린 가지배관 자리다.
    "nozzle": ((0.0, -1.0), (_SQ, -_SQ)),
}


@dataclass(frozen=True)
class LabelRequest:
    key: str
    text: str
    anchor: Point
    angle: float = 0.0
    kind: str = "link"


@dataclass(frozen=True)
class PlacedLabel:
    key: str
    text: str
    x: float
    y: float
    angle: float
    kind: str
    leader: tuple[Point, Point] | None = None


@dataclass(frozen=True)
class LayoutReport:
    placed: int = 0
    leaders: int = 0
    # 후보 자리를 다 훑고도 빈 자리가 없어 그리지 않은 라벨. 겹친 채 두지 않는다.
    dropped: tuple[str, ...] = ()
    seconds: float = 0.0


class _Grid:
    """균등 격자 색인. 라벨이 수천 개가 되어도 충돌 질의가 전수 비교로 늘지 않는다."""

    __slots__ = ("_cell", "_cells")

    def __init__(self, cell: float) -> None:
        self._cell = max(cell, 1e-9)
        self._cells: dict[tuple[int, int], list[Extent]] = {}

    def _span(self, box: Extent) -> tuple[int, int, int, int]:
        c = self._cell
        return (math.floor(box[0] / c), math.floor(box[1] / c),
                math.floor(box[2] / c), math.floor(box[3] / c))

    def hits(self, box: Extent) -> bool:
        i0, j0, i1, j1 = self._span(box)
        for i in range(i0, i1 + 1):
            for j in range(j0, j1 + 1):
                for o in self._cells.get((i, j), ()):
                    if box[0] < o[2] and o[0] < box[2] and box[1] < o[3] and o[1] < box[3]:
                        return True
        return False

    def add(self, box: Extent) -> None:
        i0, j0, i1, j1 = self._span(box)
        for i in range(i0, i1 + 1):
            for j in range(j0, j1 + 1):
                self._cells.setdefault((i, j), []).append(box)

    def count(self, box: Extent) -> int:
        i0, j0, i1, j1 = self._span(box)
        return sum(len(self._cells.get((i, j), ()))
                   for i in range(i0, i1 + 1) for j in range(j0, j1 + 1))


def box_of(measure: Callable[[str], tuple[float, float]], text: str, angle: float
           ) -> tuple[float, float]:
    """회전한 글자를 감싸는 축평행 상자. 겹침 판정은 이 상자로 한다 — 실제 글자보다
    넉넉해서 판정이 후하게 나오는 일은 없다."""
    w, h = measure(text)
    w, h = w + h * PAD, h * (1.0 + PAD)
    c, s = abs(math.cos(math.radians(angle))), abs(math.sin(math.radians(angle)))
    return w * c + h * s, w * s + h * c


def _directions(req: LabelRequest) -> tuple[Point, ...]:
    if req.kind == "link":
        th = math.radians(req.angle)
        px, py = -math.sin(th), math.cos(th)
        first = ((px, py), (-px, -py))
    else:
        first = _FIRST_DIRS.get(req.kind, _FIRST_DIRS["node"])
    return first + _RING_DIRS


def lay_out(
    requests: Iterable[LabelRequest],
    measure: Callable[[str], tuple[float, float]],
    *,
    obstacles: Sequence[Extent] = (),
    budget_s: float = 3.0,
) -> tuple[list[PlacedLabel], LayoutReport]:
    """겹치지 않는 자리에만 라벨을 놓는다.

    ``measure`` 는 회전하지 않은 글자의 (너비, 높이) 를 도면 좌표로 돌려준다.
    ``obstacles`` 는 이미 자리가 정해져 움직일 수 없는 글자(도면 주기 등)다.
    """
    started = time.perf_counter()
    items = [(r, *box_of(measure, r.text, r.angle)) for r in requests if r.text]
    if not items:
        return [], LayoutReport(seconds=time.perf_counter() - started)

    cell = max(sorted(w for _, w, _ in items)[len(items) // 2], 1e-9)
    grid = _Grid(cell)
    for box in obstacles:
        grid.add(box)

    # 붐비는 자리부터 놓는다. 한산한 곳을 먼저 채우면 정작 자리가 없는 쪽이 밀려난다.
    crowd = _Grid(cell)
    for req, _, _ in items:
        crowd.add((req.anchor[0], req.anchor[1], req.anchor[0], req.anchor[1]))
    order = sorted(range(len(items)), key=lambda i: (-_crowding(crowd, *items[i]),
                                                     items[i][0].key))

    out: list[PlacedLabel] = []
    dropped: list[str] = []
    leaders = 0
    for n, i in enumerate(order):
        req, w, h = items[i]
        if n % 64 == 0 and time.perf_counter() - started > budget_s:
            dropped.extend(items[j][0].key for j in order[n:])
            break
        spot = _place(grid, req, w, h)
        if spot is None:
            dropped.append(req.key)
            continue
        (cx, cy), far = spot
        box = (cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)
        grid.add(box)
        leader = None
        if far > h * LEADER_AFTER:
            leader = (req.anchor, (min(max(req.anchor[0], box[0]), box[2]),
                                   min(max(req.anchor[1], box[1]), box[3])))
            leaders += 1
        out.append(PlacedLabel(req.key, req.text, cx, cy, req.angle, req.kind, leader))

    return out, LayoutReport(len(out), leaders, tuple(sorted(dropped)),
                             time.perf_counter() - started)


def _crowding(crowd: _Grid, req: LabelRequest, w: float, h: float) -> int:
    x, y = req.anchor
    return crowd.count((x - w, y - h, x + w, y + h))


def _place(grid: _Grid, req: LabelRequest, w: float, h: float
           ) -> tuple[Point, float] | None:
    ax, ay = req.anchor
    dirs = _directions(req)
    for ring in range(MAX_RINGS):
        dist = h * (GAP + ring * RING_STEP)
        for ux, uy in dirs:
            reach = dist + abs(ux) * w / 2 + abs(uy) * h / 2
            cx, cy = ax + ux * reach, ay + uy * reach
            if not grid.hits((cx - w / 2, cy - h / 2, cx + w / 2, cy + h / 2)):
                return (cx, cy), dist
    return None


def overlapping_pairs(placed: Sequence[PlacedLabel],
                      measure: Callable[[str], tuple[float, float]]
                      ) -> list[tuple[str, str]]:
    """놓인 라벨끼리 겹친 쌍. 수용 기준(겹침 0 건)을 배치기 밖에서 다시 재는 자다."""
    boxes = []
    for lab in placed:
        w, h = box_of(measure, lab.text, lab.angle)
        boxes.append((lab.key, lab.x - w / 2, lab.y - h / 2, lab.x + w / 2, lab.y + h / 2))
    boxes.sort(key=lambda b: b[1])
    bad = []
    for i, a in enumerate(boxes):
        for b in boxes[i + 1:]:
            if b[1] >= a[3]:
                break
            if a[2] < b[4] and b[2] < a[4]:
                bad.append((a[0], b[0]))
    return bad
