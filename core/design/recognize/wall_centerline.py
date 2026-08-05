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

[문서정합] §3.3 6항의 스냅만으로는 **직교 접합부가 절대 닫히지 않는다**. 두께 t
인 두 벽이 만나면 각 중심선은 상대 벽의 안쪽 선에서 끝나므로 두 끝점이 t/√2 만큼
떨어진다. 스냅 공차는 0.3·t 이고 0.3 < 0.707 이라 두께와 무관하게 항상 모자란다.
C160 이 받아 주지도 않는다 — `inferred` 하한이 200mm 이라 두께 283mm 미만 벽의
코너는 전부 열린 채 남는다. 실 도면 두 장에서 실이 잘게 부서진 원인이 이것이다.

그래서 스냅 앞에 **접합점 연장**을 넣었다. 비평행한 두 중심선의 축이 만나는 점이
곧 벽 모서리이므로, 끝점을 그 교점까지 늘린다. 이건 추정이 아니라 유도다 — 그래서
가상 간선(0.45)이 아니라 실제 중심선으로 남는다.
"""
from __future__ import annotations

import math
from collections import defaultdict
from dataclasses import dataclass, field

from . import params as P
from .spatial import (
    NodeIndex, angle_bucketed_pairs, angle_deg, angle_diff, overlap_ratio,
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

    extended = extend_to_junctions(out)
    if extended:
        provenance.append(f"접합점까지 늘린 끝점 {extended}개 — 두 축의 교점이 모서리다")

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


def extend_to_junctions(centerlines: list) -> int:
    """비평행한 두 중심선의 끝을 두 축의 교점까지 늘린다. 늘린 끝점 개수를 낸다.

    직교 접합부에서 두 중심선은 서로 상대 벽 두께의 절반만큼 모자란 데서 끝난다.
    교점은 그 벽 모서리의 실제 위치이므로 여기까지 늘리는 것은 추정이 아니다.
    비스듬한 접합은 필요한 연장이 1/sin(각) 로 커지는데, 그건 벽 모서리가 아니라
    멀리서 스쳐 지나는 두 선일 수 있어 `JUNCTION_EXTEND_MAX_MM` 로 자른다.
    """
    segs = [(c.p1[0], c.p1[1], c.p2[0], c.p2[1]) for c in centerlines]
    if len(segs) < 2:
        return 0
    angles = [angle_deg(s) for s in segs]
    reaches = [_junction_reach(c.thickness_mm) for c in centerlines]

    # 두 중심선 **모두** 교점 가까이에 끝점이 있어야 하므로 두 끝점 사이 거리는
    # 2·reach 안이다. 선분이 아니라 끝점만 격자에 넣으면 되고, 긴 벽선이 격자를
    # 수백 칸씩 지나가는 비용이 사라진다. 칸 크기는 **가장 짧은** reach 에 맞춘다 —
    # 가장 긴 것에 맞추면 두께 미상 단선(reach 30mm)이 매번 큰 칸을 훑는다.
    cell = 2.0 * P.CENTERLINE_SNAP_TOL_MIN_MM
    points = [(i, which, seg[which * 2], seg[which * 2 + 1])
              for i, seg in enumerate(segs) for which in (0, 1)]
    buckets: dict[tuple[int, int], list] = defaultdict(list)
    for entry in points:
        buckets[(math.floor(entry[2] / cell), math.floor(entry[3] / cell))].append(entry)

    moves: dict[tuple[int, int], tuple[float, Point]] = {}
    for i, wi, xi, yi in points:
        span = int(math.ceil(2.0 * reaches[i] / cell))
        kx, ky = math.floor(xi / cell), math.floor(yi / cell)
        for dx in range(-span, span + 1):
            for dy in range(-span, span + 1):
                for j, wj, xj, yj in buckets.get((kx + dx, ky + dy), ()):
                    if j <= i:
                        continue
                    reach = min(reaches[i], reaches[j])
                    if math.dist((xi, yi), (xj, yj)) > 2.0 * reach:
                        continue
                    if angle_diff(angles[i], angles[j]) < P.JUNCTION_MIN_ANGLE_DEG:
                        continue
                    hit = _axis_intersection(segs[i], segs[j])
                    if hit is None:
                        continue
                    di = math.dist((xi, yi), hit)
                    dj = math.dist((xj, yj), hit)
                    if di > reach or dj > reach:
                        continue
                    if (not _is_nearest_end(segs[i], wi, hit)
                            or not _is_nearest_end(segs[j], wj, hit)):
                        continue
                    _offer(moves, i, wi, di, hit)
                    _offer(moves, j, wj, dj, hit)

    for (idx, which), (_dist, point) in moves.items():
        line = centerlines[idx]
        if which == 0:
            line.p1 = point
        else:
            line.p2 = point
    return len(moves)


def _offer(moves, idx: int, which: int, dist: float, point: Point) -> None:
    """한 끝점이 여러 교점에 걸리면 가장 가까운 것만 쓴다."""
    key = (idx, which)
    if key not in moves or dist < moves[key][0]:
        moves[key] = (dist, point)


def _junction_reach(thickness_mm) -> float:
    """이 중심선이 모서리까지 모자랄 수 있는 최대 거리.

    상대 벽 두께를 모르므로 자기 두께로 대신한다. 직교 접합에서 필요한 연장은
    상대 두께의 절반이라 같은 굵기의 벽끼리는 넉넉하다.
    """
    if thickness_mm is None:
        return P.CENTERLINE_SNAP_TOL_MIN_MM
    return min(P.JUNCTION_EXTEND_MAX_MM, max(P.CENTERLINE_SNAP_TOL_MIN_MM,
                                             float(thickness_mm)))


def _is_nearest_end(seg: Segment, which: int, point: Point) -> bool:
    """교점 쪽을 향한 끝인가.

    반대쪽 끝이 더 가까우면 교점은 선분 한가운데 — T 접합의 관통하는 쪽이다.
    늘릴 것이 없고, C170 의 교차 분할이 그 자리를 자른다.
    """
    d1 = math.dist((seg[0], seg[1]), point)
    d2 = math.dist((seg[2], seg[3]), point)
    return d1 <= d2 if which == 0 else d2 <= d1


def _axis_intersection(a: Segment, b: Segment):
    """두 축(무한직선)의 교점. 평행이면 None."""
    x1, y1, x2, y2 = a
    x3, y3, x4, y4 = b
    r_x, r_y = x2 - x1, y2 - y1
    s_x, s_y = x4 - x3, y4 - y3
    denom = r_x * s_y - r_y * s_x
    if denom == 0.0:
        return None
    t = ((x3 - x1) * s_y - (y3 - y1) * s_x) / denom
    return (x1 + r_x * t, y1 + r_y * t)


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
