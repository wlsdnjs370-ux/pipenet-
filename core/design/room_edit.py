# -*- coding: utf-8 -*-
"""실 편집의 기하 — 합치기·자르기 (지시서 §12.6).

인식기가 뽑은 face 가 늘 맞지는 않는다. 문틀이 안 닫혀 두 실이 하나로 붙거나,
도면에 없는 칸막이 때문에 한 실이 둘로 갈린다. 그 수정은 GATE 에서만 할 수 있고,
여기서 만든 폴리곤이 그대로 C4 의 헤드 배치 면적이 된다.

**모르면 만들지 않는다.** 합치기는 실들이 실제로 변을 맞대고 있을 때만, 자르기는
선이 경계를 정확히 두 번 지날 때만 된다. 어느 쪽도 아니면 그럴듯한 폴리곤을
지어내는 대신 이유를 붙여 거절한다 — 지어낸 면적은 헤드 개수를 바꾸고 그 오차는
아무도 다시 보지 않는다.
"""
from __future__ import annotations

from .recognize.spatial import polygon_area, signed_area

Point = tuple[float, float]

_MM2_PER_M2 = 1_000_000.0

# C170 이 폴리곤을 0.1mm 로 반올림해 내보낸다. 같은 눈금으로 꼭짓점을 묶어야
# 맞닿은 face 의 공유 변이 같은 변으로 인식된다.
_GRID = 10.0

# 자르는 선에서 이만큼 안쪽이면 어느 쪽인지 정할 수 없다고 본다. 0.1mm 눈금의
# 열 배 — 반올림 오차는 넘어서고 사람이 의도한 위치는 넘지 않는다.
_ON_LINE_MM = 1.0


class RoomEditError(ValueError):
    """편집을 할 수 없는 이유. 사람에게 그대로 보여줄 문장을 담는다."""


def _key(point) -> tuple[int, int]:
    return (round(float(point[0]) * _GRID), round(float(point[1]) * _GRID))


def _ring_edges(polygon: list):
    n = len(polygon)
    for i in range(n):
        yield polygon[i], polygon[(i + 1) % n]


def merge_polygons(polygons: list) -> list:
    """맞닿은 폴리곤들의 바깥 경계.

    face 는 같은 평면 그래프에서 나오므로 이웃한 두 face 는 변을 정확히 공유한다.
    그래서 두 번 나온 변을 지우면 남는 것이 곧 합쳐진 실의 경계다. 근사가 아니라
    유도다 — 대신 변을 공유하지 않는 실끼리는 애초에 합쳐지지 않는다.
    """
    seen: dict[tuple, int] = {}
    coords: dict[tuple[int, int], Point] = {}
    for polygon in polygons:
        if len(polygon) < 3:
            raise RoomEditError("꼭짓점이 3개 미만인 실은 합칠 수 없습니다.")
        for a, b in _ring_edges(polygon):
            ka, kb = _key(a), _key(b)
            if ka == kb:
                continue
            coords[ka] = (float(a[0]), float(a[1]))
            coords[kb] = (float(b[0]), float(b[1]))
            edge = (ka, kb) if ka < kb else (kb, ka)
            seen[edge] = seen.get(edge, 0) + 1

    border = [e for e, n in seen.items() if n == 1]
    if not border:
        raise RoomEditError("합칠 실들의 바깥 경계를 찾지 못했습니다.")

    adjacency: dict[tuple[int, int], list] = {}
    for ka, kb in border:
        adjacency.setdefault(ka, []).append(kb)
        adjacency.setdefault(kb, []).append(ka)
    odd = [k for k, nbrs in adjacency.items() if len(nbrs) != 2]
    if odd:
        raise RoomEditError(
            "경계가 한 줄로 이어지지 않습니다. 꼭짓점만 맞닿았거나 가운데가 뚫린 "
            "조합입니다.")

    start = border[0][0]
    ring = [start]
    prev, node = None, start
    while True:
        first, second = adjacency[node]
        nxt = first if first != prev else second
        if nxt == start:
            break
        ring.append(nxt)
        prev, node = node, nxt
    if len(ring) != len(border):
        raise RoomEditError("서로 맞닿지 않은 실이 섞여 있습니다. 붙어 있는 실만 "
                            "골라 합치세요.")

    points = [coords[k] for k in ring]
    return points if signed_area(points) > 0 else points[::-1]


def split_polygon(polygon: list, p1, p2) -> tuple[list, list]:
    """무한 직선으로 실을 둘로 가른다. 경계를 정확히 두 번 지나야 한다.

    갈래가 셋 이상 나오는 ㄷ 자 실이나, 선이 꼭짓점을 지나는 경우는 어느 조각이
    어느 실인지 사람도 화면도 같게 읽는다는 보장이 없어 거절한다.
    """
    if len(polygon) < 3:
        raise RoomEditError("꼭짓점이 3개 미만인 실은 자를 수 없습니다.")
    ax, ay = float(p1[0]), float(p1[1])
    bx, by = float(p2[0]), float(p2[1])
    dx, dy = bx - ax, by - ay
    span = (dx * dx + dy * dy) ** 0.5
    if span < _ON_LINE_MM:
        raise RoomEditError("자르는 선이 너무 짧습니다. 실을 가로질러 그으세요.")

    # 직선까지의 부호 있는 수직거리(mm). 부호가 곧 어느 쪽 조각인지다.
    dist = [((px - ax) * dy - (py - ay) * dx) / span
            for px, py in ((float(p[0]), float(p[1])) for p in polygon)]
    if any(abs(d) < _ON_LINE_MM for d in dist):
        raise RoomEditError("자르는 선이 실의 꼭짓점을 지납니다. 조금 옮겨 다시 "
                            "그으세요.")

    n = len(polygon)
    crossings = sum(1 for i in range(n) if dist[i] * dist[(i + 1) % n] < 0)
    if crossings != 2:
        raise RoomEditError(f"선이 실 경계를 {crossings}번 지납니다. 정확히 두 번 "
                            "지나도록 그으세요.")

    left: list = []
    right: list = []
    for i in range(n):
        j = (i + 1) % n
        pi = (float(polygon[i][0]), float(polygon[i][1]))
        (left if dist[i] > 0 else right).append(pi)
        if dist[i] * dist[j] >= 0:
            continue
        pj = (float(polygon[j][0]), float(polygon[j][1]))
        t = dist[i] / (dist[i] - dist[j])
        cut = (pi[0] + (pj[0] - pi[0]) * t, pi[1] + (pj[1] - pi[1]) * t)
        left.append(cut)
        right.append(cut)

    if len(left) < 3 or len(right) < 3:
        raise RoomEditError("선이 실을 가르지 못했습니다.")
    return left, right


def polygon_area_m2(polygon: list) -> float:
    """mm 좌표의 폴리곤 면적을 ㎡ 로. 면적이 곧 헤드 개수라 단위를 한 곳에 둔다."""
    return polygon_area(polygon) / _MM2_PER_M2


def perimeter(polygon: list) -> float:
    total = 0.0
    for a, b in _ring_edges(polygon):
        total += ((float(b[0]) - float(a[0])) ** 2
                  + (float(b[1]) - float(a[1])) ** 2) ** 0.5
    return total
