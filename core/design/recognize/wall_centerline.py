# -*- coding: utf-8 -*-
"""C150 — 벽 중심선화 (지시서 §3.3).

평행한 두 벽선을 하나의 중심선으로 접고, 두께를 함께 낸다. C170 이 실 폴리곤을
뽑을 때 쓰는 골격이 여기서 나온다.

[문서정합] §3.3 4항의 "중심선 = 두 선의 중점을 잇는 선분" 은 글자 그대로 읽으면
두 중점을 잇는 **벽을 가로지르는 짧은 토막**이 된다(길이 = 벽 두께). 벽의 축을
뜻한 것으로 읽고, 겹치는 구간에 대해 두 선 사이 한가운데를 지나는 선분을 낸다.
긴 선과 짧은 선이 짝지어지면 짧은 쪽이 끝나는 데서 중심선도 끝난다 — 남는 간극은
C160 이 볼 몫이지 여기서 늘려 채울 것이 아니다.

[문서정합] §3.3 의 출력 스키마에는 `confidence` 와 `provenance` 가 없는데 작업
규칙 9 는 인식 셸 결과에 둘 다 요구한다. 스키마의 네 필드는 그대로 두고 덧붙인다.
짝을 찾은 중심선과 미짝 단선은 같은 무게로 다룰 수 없다.

[문서정합] 끝점 스냅 공차가 중심선마다 다르다(`max(30, 두께*0.3)`). `NodeIndex`
가 점별 공차를 받도록 넓혔고, 병합 판정은 두 점의 공차 중 **작은 쪽**으로 한다.
그렇게 하지 않으면 어느 끝점을 먼저 넣었느냐에 따라 실이 붙었다 떨어졌다 한다.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

from . import params as P
from .spatial import (
    NodeIndex, angle_bucketed_pairs, angle_diff, overlap_ratio,
)

Point = tuple[float, float]
Segment = tuple[float, float, float, float]

WALL_REPR_DOUBLE = "double_line"
WALL_REPR_SINGLE = "single_line"


@dataclass
class Centerline:
    """벽 하나의 축. 좌표는 mm."""

    p1: Point
    p2: Point
    thickness_mm: float | None
    source_pair: tuple[int, int] | None
    unpaired: bool
    confidence: float
    provenance: list = field(default_factory=list)
    n1: int = -1
    n2: int = -1

    def to_dict(self) -> dict:
        return {
            "p1": [round(self.p1[0], 2), round(self.p1[1], 2)],
            "p2": [round(self.p2[0], 2), round(self.p2[1], 2)],
            "thickness_mm": (None if self.thickness_mm is None
                             else round(self.thickness_mm, 1)),
            "source_pair": (None if self.source_pair is None
                            else list(self.source_pair)),
            "unpaired": self.unpaired,
            "confidence": round(self.confidence, 2),
            "provenance": list(self.provenance),
            "nodes": [self.n1, self.n2],
        }


@dataclass
class CenterlineResult:
    centerlines: list
    nodes: list
    node_degree: list
    wall_repr: str
    paired_ratio: float
    provenance: list = field(default_factory=list)

    @property
    def relaxed(self) -> bool:
        """단선 표기 도면이면 C160/C170 이 완화 모드로 돈다 (§3.3 실패 신호)."""
        return self.wall_repr == WALL_REPR_SINGLE

    def to_dict(self) -> dict:
        return {
            "centerlines": [c.to_dict() for c in self.centerlines],
            "nodes": [[round(x, 2), round(y, 2)] for x, y in self.nodes],
            "node_degree": list(self.node_degree),
            "wall_repr": self.wall_repr,
            "paired_ratio": round(self.paired_ratio, 3),
            "relaxed": self.relaxed,
            "provenance": list(self.provenance),
        }


def build_centerlines(lines, *, offset_peaks_mm=(), unit_to_mm: float = 1.0) -> CenterlineResult:
    """WALL 계열 LINE 을 중심선으로 접는다.

    `offset_peaks_mm` 은 C130 이 낸 벽 두께 후보다. 비어 있으면 어떤 쌍도
    성립하지 않는다 — 두께를 모르는 채 짝지으면 가구 선 두 개도 벽이 된다.
    """
    segs: list[Segment] = []
    origin: list[int] = []
    for idx, ln in enumerate(lines):
        x1, y1, x2, y2 = (float(v) * unit_to_mm for v in ln[:4])
        if x1 == x2 and y1 == y2:
            continue
        segs.append((x1, y1, x2, y2))
        origin.append(idx)

    peaks = [float(p) for p in offset_peaks_mm]
    provenance: list[str] = []
    if not peaks:
        provenance.append("C130 오프셋 peak 이 없어 쌍 판정을 하지 못했다")

    pairs = _select_pairs(segs, peaks) if peaks else []
    paired_lines: set[int] = set()
    for entry in pairs:
        paired_lines.add(entry[1])
        paired_lines.add(entry[2])
    paired_ratio = (len(paired_lines) / len(segs)) if segs else 0.0

    wall_repr = (WALL_REPR_SINGLE
                 if paired_ratio < P.SINGLE_LINE_PARALLEL_MAX_RATIO
                 else WALL_REPR_DOUBLE)
    provenance.append(
        f"짝지은 선 {len(paired_lines)}/{len(segs)} ({paired_ratio:.2f}) "
        f"→ wall_repr={wall_repr}")

    out: list[Centerline] = []
    for _key, i, j, p1, p2, thickness in pairs:
        out.append(Centerline(
            p1=p1, p2=p2, thickness_mm=thickness,
            source_pair=(origin[i], origin[j]), unpaired=False,
            confidence=P.CONF_CENTERLINE_PAIRED,
            provenance=[f"평행쌍 오프셋 {thickness:.0f}mm 가 C130 peak 과 일치"]))
    for i, seg in enumerate(segs):
        if i in paired_lines:
            continue
        out.append(Centerline(
            p1=(seg[0], seg[1]), p2=(seg[2], seg[3]), thickness_mm=None,
            source_pair=(origin[i], origin[i]), unpaired=True,
            confidence=P.CONF_CENTERLINE_UNPAIRED,
            provenance=["짝을 못 찾은 단선 — 두께 미상"]))

    nodes, degree = snap_endpoints(out)
    return CenterlineResult(centerlines=out, nodes=nodes, node_degree=degree,
                            wall_repr=wall_repr, paired_ratio=paired_ratio,
                            provenance=provenance)


def _select_pairs(segs: list, peaks: list) -> list:
    """한 선은 한 쌍에만 들어간다. 좋은 쌍부터 확정한다.

    벽선 옆에 해치선이 나란히 놓이면 한 선에 후보가 여럿 생긴다. 전부 받으면
    같은 벽에서 중심선이 두 개 나와 C170 이 벽 속에 실을 하나 만든다.
    """
    candidates = []
    for i, j, ai, aj in angle_bucketed_pairs(
            segs, bucket_deg=P.CENTERLINE_ANGLE_BUCKET_DEG,
            cell=P.PARALLEL_OFFSET_MAX_MM):
        if angle_diff(ai, aj) > P.CENTERLINE_ANGLE_TOL_DEG:
            continue
        geom = _pair_geometry(segs[i], segs[j])
        if geom is None:
            continue
        p1, p2, thickness = geom
        deviation = min((abs(thickness - pk) for pk in peaks), default=None)
        if deviation is None or deviation > P.CENTERLINE_OFFSET_MATCH_TOL_MM:
            continue
        overlap = overlap_ratio(segs[i], segs[j])
        if overlap < P.CENTERLINE_OVERLAP_MIN_RATIO:
            continue
        candidates.append(((round(deviation, 6), -round(overlap, 6), i, j),
                           i, j, p1, p2, thickness))

    used: set[int] = set()
    chosen = []
    for entry in sorted(candidates, key=lambda c: c[0]):
        _key, i, j = entry[0], entry[1], entry[2]
        if i in used or j in used:
            continue
        used.add(i)
        used.add(j)
        chosen.append(entry)
    return chosen


def _pair_geometry(a: Segment, b: Segment):
    """겹치는 구간에 대한 중심축과 두께. 겹치지 않으면 None."""
    ax1, ay1, ax2, ay2 = a
    dx, dy = ax2 - ax1, ay2 - ay1
    la = math.hypot(dx, dy)
    if la == 0.0:
        return None
    ux, uy = dx / la, dy / la

    ts = [0.0, la,
          (b[0] - ax1) * ux + (b[1] - ay1) * uy,
          (b[2] - ax1) * ux + (b[3] - ay1) * uy]
    lo = max(min(ts[0], ts[1]), min(ts[2], ts[3]))
    hi = min(max(ts[0], ts[1]), max(ts[2], ts[3]))
    if hi <= lo:
        return None

    bmx, bmy = (b[0] + b[2]) * 0.5, (b[1] + b[3]) * 0.5
    signed = (bmx - ax1) * (-uy) + (bmy - ay1) * ux
    hx, hy = -uy * signed * 0.5, ux * signed * 0.5
    p1 = (ax1 + ux * lo + hx, ay1 + uy * lo + hy)
    p2 = (ax1 + ux * hi + hx, ay1 + uy * hi + hy)
    return p1, p2, abs(signed)


def snap_tol(thickness_mm) -> float:
    """§3.3 6항. 두꺼운 벽일수록 접합부가 벌어져 있다."""
    if thickness_mm is None:
        return P.CENTERLINE_SNAP_TOL_MIN_MM
    return max(P.CENTERLINE_SNAP_TOL_MIN_MM,
               thickness_mm * P.CENTERLINE_SNAP_TOL_RATIO)


def snap_endpoints(centerlines: list):
    tol_max = max([P.CENTERLINE_SNAP_TOL_MIN_MM]
                  + [snap_tol(c.thickness_mm) for c in centerlines])
    index = NodeIndex(tol_max)
    for c in centerlines:
        tol = snap_tol(c.thickness_mm)
        c.n1 = index.add(c.p1[0], c.p1[1], tol)
        c.n2 = index.add(c.p2[0], c.p2[1], tol)
    degree = [0] * len(index)
    for c in centerlines:
        degree[c.n1] += 1
        degree[c.n2] += 1
    return list(index.points), degree
