# -*- coding: utf-8 -*-
"""C160 — 개구부 간극 가상 폐합 (지시서 §3.4).

지시서가 **최대 위험 지점**이라 부른 곳이다. 여기서 간극을 잘못 이으면 두 실이
하나가 되고, 면적이 두 배가 되고, 용도 판정이 틀리고, 헤드 개수가 틀린다. 그래서
이 모듈은 "닫는 데 성공했는가" 가 아니라 **무엇을 근거로 닫았는가**를 낸다.

§3.4 절대 규칙 — 모든 가상 간선은 `is_virtual: true` 를 달고 캔버스에서 점선 +
다른 색으로 그린다. 검수자가 여기를 먼저 봐야 한다.

[문서정합] §3.4 의 출력은 `virtual_edges[]` 뿐이지만 `ClosureResult` 로 감싸
`provenance` 를 함께 낸다. 후보를 자른 사실(성능 상한)이나 완화 모드로 돌았다는
사실은 간선 목록에 실을 자리가 없는데, 그게 없으면 **왜 이 간극이 안 닫혔는지**를
검수자가 알 방법이 없다. 안 닫힌 간극은 실이 하나로 합쳐지는 그 사고다.

[문서정합] "양끝이 같은 직선(공선) 위" 를 각도차만으로 보면 나란히 떨어진 두 벽도
공선이 된다. 두 중심선의 각도차와 **간극 자신의 각도**를 함께 본다.
"""
from __future__ import annotations

import heapq
import math
from dataclasses import dataclass, field

from . import params as P
from .spatial import angle_deg, angle_diff

Point = tuple[float, float]

DOOR = "door"
OPENING = "opening"
INFERRED = "inferred"


@dataclass
class DoorArc:
    """문 호. `anchors` 는 중심(경첩)과 두 끝점이다.

    [문서정합] §3.4 는 "간극 양끝 300mm 이내에 DOOR ARC 의 **끝점**" 이라고 쓰지만,
    같은 행의 "ARC 반경이 간극 폭의 0.8~1.3배" 가 그 읽기를 부정한다. 문 호는
    경첩이 한쪽 문설주(=중심)이고 닫힌 위치의 끝점이 반대쪽 문설주다. 두 **끝점**을
    양쪽 문설주로 놓으면 90° 호에서 반경이 간극의 0.71배가 되어 표의 대역 밖으로
    나간다. 중심과 끝점을 함께 앵커로 본다 — 서로 다른 앵커여야 한다.
    """

    center: Point
    p1: Point
    p2: Point
    radius_mm: float

    @property
    def anchors(self) -> tuple:
        return (self.center, self.p1, self.p2)


@dataclass
class VirtualEdge:
    p1: Point
    p2: Point
    n1: int
    n2: int
    kind: str
    confidence: float
    gap_mm: float
    evidence: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "p1": [round(self.p1[0], 2), round(self.p1[1], 2)],
            "p2": [round(self.p2[0], 2), round(self.p2[1], 2)],
            "nodes": [self.n1, self.n2],
            "kind": self.kind,
            "confidence": round(self.confidence, 2),
            "gap_mm": round(self.gap_mm, 1),
            "evidence": list(self.evidence),
            "is_virtual": True,
        }


@dataclass
class ClosureResult:
    virtual_edges: list
    open_endpoints: int
    relaxed: bool
    provenance: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "virtual_edges": [e.to_dict() for e in self.virtual_edges],
            "open_endpoints": self.open_endpoints,
            "relaxed": self.relaxed,
            "provenance": list(self.provenance),
        }


def door_arcs(entities, *, unit_to_mm: float = 1.0) -> list:
    """DOOR 로 판정된 레이어의 ARC 에서 양 끝점과 반경을 뽑는다."""
    out = []
    for ent in entities:
        if ent.get("t") != "A":
            continue
        cx, cy = ent["c"]
        radius = float(ent["r"]) * unit_to_mm
        start, end = ent["a"]
        cx, cy = float(cx) * unit_to_mm, float(cy) * unit_to_mm
        out.append(DoorArc(
            center=(cx, cy),
            p1=_on_arc(cx, cy, radius, start),
            p2=_on_arc(cx, cy, radius, end),
            radius_mm=radius))
    return out


def _on_arc(cx: float, cy: float, r: float, deg: float) -> Point:
    rad = math.radians(float(deg))
    return (cx + r * math.cos(rad), cy + r * math.sin(rad))


def close_openings(result, doors=(), *, unit_to_mm: float = 1.0) -> ClosureResult:
    """중심선의 열린 끝점 사이 간극을 근거를 달아 잇는다.

    `result` 는 C150 의 `CenterlineResult`. 이미 다른 중심선과 이어진 끝점
    (스냅 군집 크기 ≥ 2)은 건드리지 않는다.
    """
    relaxed = bool(getattr(result, "relaxed", False))
    provenance: list[str] = []
    if relaxed:
        provenance.append("wall_repr=single_line — 공차 완화 모드 (간극 폭 범위는 그대로)")

    free = [nid for nid, deg in enumerate(result.node_degree) if deg == 1]
    heading = _free_headings(result)
    arcs = door_arcs(doors, unit_to_mm=unit_to_mm)
    arc_cells = _arc_endpoint_cells(arcs)

    grid = _endpoint_grid(result.nodes, free)
    # 한 중심선의 두 끝을 잇는 것은 벽을 자기 자신으로 되돌리는 짓이다. 짧은 벽
    # 하나가 통째로 실 하나가 된다.
    own = {frozenset((c.n1, c.n2)) for c in result.centerlines}
    truncated = 0
    candidates = []
    for e in free:
        near = _neighbors(grid, result.nodes[e])
        if len(near) > P.GAP_CANDIDATE_MAX:
            near = heapq.nsmallest(
                P.GAP_CANDIDATE_MAX, near,
                key=lambda f: _dist(result.nodes[e], result.nodes[f]))
            truncated += 1
        for f in near:
            if f <= e or frozenset((e, f)) in own:
                continue
            pe, pf = result.nodes[e], result.nodes[f]
            gap = _dist(pe, pf)
            if gap < P.GAP_MIN_MM or gap > P.GAP_SEARCH_RADIUS_MM:
                continue
            verdict = _classify_gap(pe, pf, gap, heading.get(e), heading.get(f),
                                    arcs, arc_cells, relaxed)
            if verdict is None:
                continue
            kind, confidence, evidence = verdict
            candidates.append((-confidence, gap, e, f, kind, evidence))

    if truncated:
        provenance.append(
            f"끝점 {truncated}곳에서 후보를 가까운 {P.GAP_CANDIDATE_MAX}개로 잘랐다")

    used: set[int] = set()
    edges = []
    for _neg, gap, e, f, kind, evidence in sorted(candidates, key=lambda c: c[:4]):
        if e in used or f in used:
            continue
        used.add(e)
        used.add(f)
        edges.append(VirtualEdge(
            p1=result.nodes[e], p2=result.nodes[f], n1=e, n2=f, kind=kind,
            confidence=-_neg, gap_mm=gap, evidence=evidence))

    provenance.append(
        f"열린 끝점 {len(free)}개 중 {len(used)}개를 가상 간선 {len(edges)}개로 이었다")
    return ClosureResult(virtual_edges=edges, open_endpoints=len(free),
                         relaxed=relaxed, provenance=provenance)


def _classify_gap(pe, pf, gap, ang_e, ang_f, arcs, arc_cells, relaxed):
    """§3.4 증거표. 강한 근거부터 본다."""
    door = _door_evidence(pe, pf, gap, arcs, arc_cells, relaxed)
    if door is not None:
        return DOOR, P.CONF_VE_DOOR, door

    if ang_e is None or ang_f is None:
        return None

    tol = (P.RELAXED_COLLINEAR_TOL_DEG if relaxed
           else P.OPENING_COLLINEAR_TOL_DEG)
    gap_angle = angle_deg((pe[0], pe[1], pf[0], pf[1]))
    collinear = (angle_diff(ang_e, ang_f) <= tol
                 and angle_diff(ang_e, gap_angle) <= tol
                 and angle_diff(ang_f, gap_angle) <= tol)

    if collinear and P.OPENING_GAP_MIN_MM <= gap <= P.OPENING_GAP_MAX_MM:
        return OPENING, P.CONF_VE_OPENING, [
            f"간극 폭 {gap:.0f}mm 가 개구부 범위",
            f"양끝 중심선이 공선 (각도차 ≤{tol:.0f}°)"]

    if not collinear and P.INFERRED_GAP_MIN_MM <= gap <= P.INFERRED_GAP_MAX_MM:
        return INFERRED, P.CONF_VE_INFERRED, [
            f"간극 폭 {gap:.0f}mm",
            f"양끝 중심선 방향이 다르다 ({angle_diff(ang_e, ang_f):.0f}°) — 코너 미접합 추정"]

    return None


def _door_evidence(pe, pf, gap, arcs, arc_cells, relaxed):
    if not arcs:
        return None
    lo = (P.RELAXED_DOOR_RADIUS_MIN_RATIO if relaxed
          else P.DOOR_GAP_RADIUS_MIN_RATIO)
    hi = (P.RELAXED_DOOR_RADIUS_MAX_RATIO if relaxed
          else P.DOOR_GAP_RADIUS_MAX_RATIO)
    near_e = _anchors_near(arcs, arc_cells, pe)
    near_f = _anchors_near(arcs, arc_cells, pf)
    for idx in sorted(near_e.keys() & near_f.keys()):
        arc = arcs[idx]
        if not (gap * lo <= arc.radius_mm <= gap * hi):
            continue
        # 같은 앵커 하나가 간극 양끝에 다 걸리면 근거가 아니다 — 짧은 간극에서 난다.
        if not (near_e[idx] - near_f[idx]) or not (near_f[idx] - near_e[idx]):
            continue
        return [f"DOOR ARC 의 서로 다른 앵커가 간극 양끝 "
                f"{P.DOOR_GAP_ENDPOINT_TOL_MM:.0f}mm 이내",
                f"ARC 반경 {arc.radius_mm:.0f}mm 가 간극 폭 {gap:.0f}mm 의 "
                f"{arc.radius_mm / gap:.2f}배"]
    return None


def _arc_endpoint_cells(arcs) -> dict:
    cell = P.DOOR_GAP_ENDPOINT_TOL_MM
    out: dict = {}
    for idx, arc in enumerate(arcs):
        for slot, (x, y) in enumerate(arc.anchors):
            out.setdefault((math.floor(x / cell), math.floor(y / cell)),
                           []).append((idx, slot))
    return out


def _anchors_near(arcs, arc_cells, point) -> dict:
    """공차 안에 있는 `{arc 번호: {앵커 번호}}`. 격자는 후보만 좁힌다."""
    cell = P.DOOR_GAP_ENDPOINT_TOL_MM
    cx, cy = math.floor(point[0] / cell), math.floor(point[1] / cell)
    out: dict = {}
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for idx, slot in arc_cells.get((cx + dx, cy + dy), ()):
                if _dist(arcs[idx].anchors[slot], point) <= cell:
                    out.setdefault(idx, set()).add(slot)
    return out


def _free_headings(result) -> dict:
    """열린 끝점마다 그 끝점을 가진 중심선의 방향. 차수 1 이라 하나뿐이다."""
    out: dict = {}
    for c in result.centerlines:
        for near, far in ((c.n1, c.n2), (c.n2, c.n1)):
            if result.node_degree[near] != 1:
                continue
            (x1, y1), (x2, y2) = result.nodes[near], result.nodes[far]
            if (x1, y1) != (x2, y2):
                out[near] = angle_deg((x1, y1, x2, y2))
    return out


def _endpoint_grid(nodes, free) -> dict:
    cell = P.GAP_SEARCH_RADIUS_MM
    grid: dict = {}
    for nid in free:
        x, y = nodes[nid]
        grid.setdefault((math.floor(x / cell), math.floor(y / cell)), []).append(nid)
    return grid


def _neighbors(grid, point) -> list:
    """반경 안에 있을 수 있는 열린 끝점. 실제 거리 판정은 호출자가 한다."""
    cell = P.GAP_SEARCH_RADIUS_MM
    cx, cy = math.floor(point[0] / cell), math.floor(point[1] / cell)
    out: list = []
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            out.extend(grid.get((cx + dx, cy + dy), ()))
    return out


def _dist(a, b) -> float:
    return math.hypot(a[0] - b[0], a[1] - b[1])
