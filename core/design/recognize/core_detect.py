# -*- coding: utf-8 -*-
"""C190 — 코어(계단·승강기·샤프트) 판별 (지시서 §3.7).

§3.7 이 제일 먼저 하는 말은 **단층만으로는 확정하지 마라** 다. 샤프트의 결정적
증거는 여러 층에서 같은 자리에 같은 크기로 반복되는 것이고, 단층 도면에서 이름만
보고 확정하면 헤드를 달지 말아야 할 곳에 달거나 그 반대가 된다.

[문서정합] §3.7 은 "다른 층에서 … 폴리곤 **개수** n" 을 세고 `n ≥ (층수-1)*0.6`
과 비교한다. 개수를 그대로 세면 한 층에 나란히 붙은 샤프트 둘이 두 번 세어져,
2층 도면에서 한 층만 보고도 문턱을 넘는다. 문턱이 층수로 쓰여 있으므로 **층 수**를
센다 — 한 층에서 몇 개가 맞든 그 층은 한 표다.

[문서정합] 이름 힌트 표가 §3.7 에는 없다(§3.6 의 USE_HINTS 는 용도용이다). 단층
도면에서 후보를 내려면 필요하므로 `CORE_HINTS` 를 여기 둔다. 전부 conf 0.40 이고
GATE 확정이 필수라, 짧은 약어가 헛짚어도 사람이 거르는 자리에서 멈춘다.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

from . import params as P

Point = tuple[float, float]

SHAFT = "shaft"
STAIR = "stair"
ELEVATOR = "elevator"

CORE_HINTS = {
    SHAFT: ["PS", "AD", "DS", "TPS", "EPS", "샤프트", "덕트", "파이프", "배관"],
    STAIR: ["계단", "STAIR", "피난"],
    ELEVATOR: ["ELEV", "E/V", "EV", "승강기", "엘리베이터"],
}


@dataclass
class CoreCandidate:
    """코어 후보 하나. 확정은 언제나 GATE 가 한다."""

    floor: int
    face_index: int
    kind: str
    confidence: float
    center: Point
    area_m2: float
    evidence: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "floor": self.floor,
            "face_index": self.face_index,
            "kind": self.kind,
            "confidence": round(self.confidence, 2),
            "center": [round(self.center[0], 1), round(self.center[1], 1)],
            "area_m2": round(self.area_m2, 3),
            "evidence": list(self.evidence),
            # §3.7 — 단층이든 다층이든 사람이 확정한다.
            "needs_confirm": True,
        }


@dataclass
class CoreResult:
    candidates: list
    floor_count: int
    provenance: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "candidates": [c.to_dict() for c in self.candidates],
            "floor_count": self.floor_count,
            "provenance": list(self.provenance),
        }


def detect_cores(floors, names=None) -> CoreResult:
    """층별 실 폴리곤에서 코어 후보를 낸다.

    `floors` 는 층마다의 `RoomFace` 목록이고, `names` 를 주면 같은 모양의 실명
    목록으로 읽는다(C180 의 결과). 층이 하나뿐이면 이름 힌트만으로 후보를 낸다.
    """
    floors = [list(f) for f in floors]
    small = [_small_indices(faces) for faces in floors]
    provenance: list[str] = []

    if len(floors) < 2:
        provenance.append("층 도면이 하나뿐이다 — 반복 증거를 볼 수 없다")
        candidates = _by_name(floors, small, names, provenance)
        return CoreResult(candidates=candidates, floor_count=len(floors),
                          provenance=provenance)

    grids = [_center_grid(faces, idxs) for faces, idxs in zip(floors, small)]
    need = (len(floors) - 1) * P.CORE_FLOOR_SHARE_MIN
    candidates: list[CoreCandidate] = []
    claimed: set[tuple[int, int]] = set()

    for f, idxs in enumerate(small):
        for i in idxs:
            face = floors[f][i]
            hits = [g for g in range(len(floors))
                    if g != f and _matches(floors[g], grids[g], face)]
            if len(hits) < need:
                continue
            claimed.add((f, i))
            candidates.append(CoreCandidate(
                floor=f, face_index=i, kind=SHAFT,
                confidence=P.CONF_SHAFT_MULTIFLOOR,
                center=face.center, area_m2=face.area_m2,
                evidence=[f"{len(hits)}개 층에서 같은 자리에 같은 크기로 반복 "
                          f"(문턱 {need:.1f}층)",
                          f"면적 {face.area_m2:.2f}㎡, 중심 허용 "
                          f"{P.CORE_CENTER_DIST_MAX_MM:.0f}mm"]))

    provenance.append(
        f"{len(floors)}개 층에서 소형 폴리곤 {sum(len(s) for s in small)}개를 대조해 "
        f"{len(candidates)}개를 SHAFT 로 봤다")
    candidates += _by_name(floors, small, names, provenance, skip=claimed)
    return CoreResult(candidates=candidates, floor_count=len(floors),
                      provenance=provenance)


def _by_name(floors, small, names, provenance, skip=frozenset()) -> list:
    """이름 힌트만으로 낸 후보 — §3.7 은 conf 0.40 에 GATE 확정 필수라고 못 박는다."""
    if not names:
        return []
    out = []
    for f, idxs in enumerate(small):
        floor_names = names[f] if f < len(names) else ()
        for i in idxs:
            if (f, i) in skip:
                continue
            name = floor_names[i] if i < len(floor_names) else None
            kind = kind_of(name) if name else None
            if kind is None:
                continue
            face = floors[f][i]
            out.append(CoreCandidate(
                floor=f, face_index=i, kind=kind,
                confidence=P.CONF_SHAFT_SINGLE,
                center=face.center, area_m2=face.area_m2,
                evidence=[f"실명 '{name}' 이 {kind} 힌트와 맞는다 — 반복 증거는 없다"]))
    if out:
        provenance.append(f"이름 힌트만으로 낸 후보 {len(out)}개 — GATE 확정 필수")
    return out


def kind_of(name: str) -> str | None:
    upper = str(name).upper()
    for kind, hints in CORE_HINTS.items():
        if any(hint in upper for hint in hints):
            return kind
    return None


def _small_indices(faces) -> list:
    """§3.2 SHAFT 행과 같은 소형 범위. 이보다 크면 실이고 작으면 틈이다."""
    return [i for i, f in enumerate(faces)
            if P.SHAFT_AREA_MIN_M2 <= f.area_m2 <= P.SMALL_CLOSED_AREA_MAX_M2]


def _center_grid(faces, idxs) -> dict:
    cell = P.CORE_CENTER_DIST_MAX_MM
    grid: dict = {}
    for i in idxs:
        x, y = faces[i].center
        grid.setdefault((math.floor(x / cell), math.floor(y / cell)), []).append(i)
    return grid


def _matches(faces, grid, face) -> bool:
    cell = P.CORE_CENTER_DIST_MAX_MM
    cx, cy = math.floor(face.center[0] / cell), math.floor(face.center[1] / cell)
    for dx in (-1, 0, 1):
        for dy in (-1, 0, 1):
            for i in grid.get((cx + dx, cy + dy), ()):
                other = faces[i]
                if math.dist(other.center, face.center) > P.CORE_CENTER_DIST_MAX_MM:
                    continue
                if not face.area_m2:
                    continue
                ratio = other.area_m2 / face.area_m2
                if P.CORE_AREA_RATIO_MIN <= ratio <= P.CORE_AREA_RATIO_MAX:
                    return True
    return False
