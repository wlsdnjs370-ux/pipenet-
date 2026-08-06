# -*- coding: utf-8 -*-
"""C170 — 폐합 영역 → 실 폴리곤 (지시서 §3.5).

중심선(C150)과 가상 간선(C160)을 평면 그래프로 만들고 face 를 뽑는다. 여기서
나온 폴리곤이 실이고, 실 면적이 헤드 개수를 정하므로 잘못 뽑히면 그 뒤가 전부
틀린다.

[문서정합] §3.5 5항은 "면적이 음수인 face(= 외곽) **1개**는 버린다" 고 하지만,
외곽 face 는 연결 요소마다 하나씩 나온다. 실 도면의 벽 그래프는 통으로 이어져
있지 않으므로 음수 face 를 **전부** 버린다. 하나만 버리면 나머지 외곽이 실로
둔갑해 건물 바깥이 방이 된다.

[문서정합] face 필터표의 "가상 간선 비율" 이 개수인지 길이인지 적혀 있지 않다.
같은 절의 신뢰도 식이 길이 비율(`가상 간선 길이 / 전체 둘레`)을 쓰므로 길이로
읽는다. 짧은 문 간극 여러 개보다 긴 벽 하나가 통째로 추정인 쪽이 더 위험하다.

[문서정합] 작업 규칙 9 에 따라 face 마다 `confidence` 와 `provenance` 를 단다.
버린 face 의 개수도 결과에 싣는다 — 실이 안 나온 이유를 검수자가 알아야 한다.
"""
from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass, field

from . import params as P
from .spatial import NodeIndex, SegmentGrid, centroid, signed_area

Point = tuple[float, float]

_MM2_PER_M2 = 1_000_000.0

SUSPICIOUS_COMPLEXITY = "suspicious_complexity"
MOSTLY_VIRTUAL = "mostly_virtual"


@dataclass
class RoomFace:
    """실 하나의 후보 폴리곤. 좌표는 mm, 면적은 ㎡."""

    polygon: list
    node_cycle: list
    area_m2: float
    perimeter_mm: float
    virtual_ratio: float
    unpaired_ratio: float
    confidence: float
    center: Point
    flags: list = field(default_factory=list)
    provenance: list = field(default_factory=list)

    @property
    def edge_count(self) -> int:
        return len(self.node_cycle)

    def to_dict(self) -> dict:
        return {
            "polygon": [[round(x, 1), round(y, 1)] for x, y in self.polygon],
            "area_m2": round(self.area_m2, 3),
            "perimeter_mm": round(self.perimeter_mm, 1),
            "edge_count": self.edge_count,
            "virtual_ratio": round(self.virtual_ratio, 3),
            "unpaired_ratio": round(self.unpaired_ratio, 3),
            "center": [round(self.center[0], 1), round(self.center[1], 1)],
            "confidence": round(self.confidence, 2),
            "flags": list(self.flags),
            "provenance": list(self.provenance),
        }


@dataclass
class FaceResult:
    faces: list
    dropped: dict
    provenance: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "faces": [f.to_dict() for f in self.faces],
            "dropped": dict(self.dropped),
            "provenance": list(self.provenance),
        }


def build_faces(result, closure=None, *, bbox_area_mm2: float | None = None) -> FaceResult:
    """중심선 + 가상 간선에서 실 폴리곤을 뽑는다.

    `result` 는 C150 의 `CenterlineResult`, `closure` 는 C160 의 `ClosureResult`.
    가상 간선은 C150 이 매긴 노드 번호를 그대로 쓰므로 좌표를 다시 맞출 필요가 없다.
    """
    points = list(result.nodes)
    edges: list[tuple] = []
    for c in result.centerlines:
        edges.append((c.n1, c.n2, False, bool(c.unpaired)))
    for e in getattr(closure, "virtual_edges", ()) or ():
        edges.append((e.n1, e.n2, True, False))

    provenance: list[str] = []
    dropped = {"outer": 0, "too_small": 0, "too_large": 0, "degenerate": 0}
    if not edges:
        provenance.append("간선이 없어 face 를 뽑지 못했다")
        return FaceResult(faces=[], dropped=dropped, provenance=provenance)

    points, edges, n_cuts = _split_at_crossings(points, edges)
    if n_cuts:
        provenance.append(f"교차점 {n_cuts}곳에서 간선을 잘랐다")

    if bbox_area_mm2 is None:
        bbox_area_mm2 = _bbox_area(points)
    area_max = bbox_area_mm2 * P.FACE_AREA_MAX_BBOX_RATIO / _MM2_PER_M2

    faces: list[RoomFace] = []
    for cycle in _walk_faces(points, edges):
        polygon = [points[u] for u, _idx in cycle]
        area = signed_area(polygon)
        if area <= 0.0:
            dropped["outer" if area < 0.0 else "degenerate"] += 1
            continue
        face = _score(polygon, cycle, edges, points, area / _MM2_PER_M2)
        if face.area_m2 < P.FACE_AREA_MIN_M2:
            dropped["too_small"] += 1
            continue
        if area_max > 0.0 and face.area_m2 > area_max:
            dropped["too_large"] += 1
            continue
        faces.append(face)

    faces.sort(key=lambda f: -f.area_m2)
    provenance.append(
        f"face {len(faces)}개 채택 — 외곽 {dropped['outer']}, "
        f"면적 미달 {dropped['too_small']}, 과대 {dropped['too_large']}, "
        f"퇴화 {dropped['degenerate']} 버림")
    flagged = sum(1 for f in faces if f.flags)
    if flagged:
        provenance.append(f"플래그가 붙은 face {flagged}개 — 검수 우선순위 상위")
    return FaceResult(faces=faces, dropped=dropped, provenance=provenance)


def _score(polygon, cycle, edges, points, area_m2: float) -> RoomFace:
    perimeter = 0.0
    virtual_len = 0.0
    unpaired = 0
    for i, (u, idx) in enumerate(cycle):
        v = cycle[(i + 1) % len(cycle)][0]
        seg = math.dist(points[u], points[v])
        perimeter += seg
        if edges[idx][2]:
            virtual_len += seg
        if edges[idx][3]:
            unpaired += 1

    virtual_ratio = (virtual_len / perimeter) if perimeter else 0.0
    unpaired_ratio = unpaired / len(cycle)
    many_edges = len(cycle) > P.FACE_EDGE_COUNT_PENALTY_MIN
    conf = (P.FACE_CONF_BASE
            - P.FACE_CONF_VIRTUAL_PENALTY * virtual_ratio
            - P.FACE_CONF_UNPAIRED_PENALTY * unpaired_ratio
            - (P.FACE_CONF_MANY_EDGES_PENALTY if many_edges else 0.0))

    flags = []
    if len(cycle) > P.FACE_EDGE_COUNT_SUSPICIOUS:
        flags.append(SUSPICIOUS_COMPLEXITY)
    if virtual_ratio > P.FACE_VIRTUAL_RATIO_SUSPICIOUS:
        flags.append(MOSTLY_VIRTUAL)

    provenance = [f"둘레 {perimeter:.0f}mm 중 가상 간선 {virtual_len:.0f}mm "
                  f"({virtual_ratio:.0%})"]
    if unpaired:
        provenance.append(f"두께 미상 중심선 {unpaired}/{len(cycle)}변")
    if many_edges:
        provenance.append(f"변 {len(cycle)}개 — 형상이 복잡하다")

    return RoomFace(
        polygon=polygon, node_cycle=[u for u, _ in cycle], area_m2=area_m2,
        perimeter_mm=perimeter, virtual_ratio=virtual_ratio,
        unpaired_ratio=unpaired_ratio, confidence=max(0.0, conf),
        center=centroid(polygon), flags=flags, provenance=provenance)


def _split_at_crossings(points, edges):
    """X 자로 겹친 두 간선을 교차점에서 자른다 (§3.5 1항).

    자르지 않으면 두 벽이 서로를 통과해 지나가고, 그 벽들이 갈라 놓은 두 실이
    하나로 뽑힌다. 평행하게 겹치는 경우는 여기서 보지 않는다 — 그건 C150 이
    한 쌍으로 접었어야 할 몫이다.
    """
    segs = [(points[a][0], points[a][1], points[b][0], points[b][1])
            for a, b, _v, _u in edges]
    lengths = sorted(math.dist((s[0], s[1]), (s[2], s[3])) for s in segs)
    median = lengths[len(lengths) // 2] if lengths else 0.0
    grid = SegmentGrid(max(P.FACE_SPLIT_CELL_MIN_MM, median))

    cells = []
    for i, seg in enumerate(segs):
        walked = grid.walk(seg)
        grid.add(i, walked)
        cells.append(walked)

    cuts: dict[int, list] = defaultdict(list)
    for i, seg in enumerate(segs):
        li = math.dist((seg[0], seg[1]), (seg[2], seg[3]))
        for j in grid.lookup(cells[i]):
            if j <= i:
                continue
            hit = _crossing(seg, segs[j])
            if hit is None:
                continue
            t, u, point = hit
            lj = math.dist((segs[j][0], segs[j][1]), (segs[j][2], segs[j][3]))
            if _interior(t, li):
                cuts[i].append((t, point))
            if _interior(u, lj):
                cuts[j].append((u, point))

    if not cuts:
        return points, edges, 0

    index = NodeIndex(P.FACE_SPLIT_TOL_MM)
    remap = [index.add(x, y) for x, y in points]
    out: list[tuple] = []
    for i, (a, b, virtual, unpaired) in enumerate(edges):
        chain = [remap[a]]
        for _t, (x, y) in sorted(cuts.get(i, ())):
            chain.append(index.add(x, y))
        chain.append(remap[b])
        for lo, hi in zip(chain, chain[1:]):
            if lo != hi:
                out.append((lo, hi, virtual, unpaired))
    return list(index.points), out, sum(len(v) for v in cuts.values())


def _interior(t: float, seg_len: float) -> bool:
    """간선 안쪽에서 잘리는가. 끝점 근처는 이미 노드 스냅이 처리했다."""
    return (P.FACE_SPLIT_TOL_MM < t * seg_len
            and P.FACE_SPLIT_TOL_MM < (1.0 - t) * seg_len)


def _crossing(a, b):
    """두 선분의 교차. 평행하거나 선분 밖이면 None."""
    x1, y1, x2, y2 = a
    x3, y3, x4, y4 = b
    dxa, dya = x2 - x1, y2 - y1
    dxb, dyb = x4 - x3, y4 - y3
    denom = dxa * dyb - dya * dxb
    if denom == 0.0:
        return None
    t = ((x3 - x1) * dyb - (y3 - y1) * dxb) / denom
    u = ((x3 - x1) * dya - (y3 - y1) * dxa) / denom
    if not (0.0 <= t <= 1.0 and 0.0 <= u <= 1.0):
        return None
    return t, u, (x1 + dxa * t, y1 + dya * t)


def _walk_faces(points, edges):
    """half-edge 순회 (§3.5 2~4항). `[(노드, 간선 번호), ...]` 순환을 낸다.

    간선 `(u→v)` 로 들어왔으면 `v` 에서 `(v→u)` 의 **바로 다음 시계방향** 간선을
    택한다. 그러면 내부 face 는 반시계(면적 양수), 외곽 face 는 시계(음수)로 나온다.
    """
    adj: dict[int, list] = defaultdict(list)
    for idx, (a, b, _v, _u) in enumerate(edges):
        if a == b:
            continue
        ax, ay = points[a]
        bx, by = points[b]
        adj[a].append((math.atan2(by - ay, bx - ax), idx, b))
        adj[b].append((math.atan2(ay - by, ax - bx), idx, a))
    for lst in adj.values():
        lst.sort()
    slot = {(node, idx): i
            for node, lst in adj.items()
            for i, (_ang, idx, _other) in enumerate(lst)}

    visited: set[int] = set()
    for start in range(2 * len(edges)):
        if start in visited or edges[start >> 1][0] == edges[start >> 1][1]:
            continue
        cycle = []
        half = start
        while half not in visited:
            visited.add(half)
            idx = half >> 1
            a, b, _v, _u = edges[idx]
            u, v = (a, b) if half & 1 == 0 else (b, a)
            cycle.append((u, idx))
            lst = adj[v]
            _ang, nidx, _other = lst[(slot[(v, idx)] - 1) % len(lst)]
            half = 2 * nidx + (0 if edges[nidx][0] == v else 1)
        if cycle:
            yield cycle


def _bbox_area(points) -> float:
    if not points:
        return 0.0
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return (max(xs) - min(xs)) * (max(ys) - min(ys))
