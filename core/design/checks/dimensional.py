# -*- coding: utf-8 -*-
"""차원 무결성 검사 — 지시서 §13.5.

정답지가 없어도 **자기모순**은 잡힌다. 같은 평면이 반복되는 층에서 헤드 수가
다르면 둘 중 하나는 틀렸다 — 어느 쪽이 틀렸는지는 몰라도 틀렸다는 것은 안다.
그래서 기준층 교차검증이 최우선이고, 이 파일에서 가장 먼저 나온다.

검사 결과에는 `pass` / `flag` 말고 **`unverified`** 가 있다. 재지 못한 것을
`pass` 로 적으면 검사표가 초록으로 가득 차고, 그 초록은 아무 의미가 없다.
"""
from __future__ import annotations

import math

from ..recognize.spatial import point_in_polygon, polygon_area

# 층 겹침을 재는 격자. C190 이 층 사이 폴리곤을 같은 코어로 볼 때 쓰는 중심 거리
# 공차와 같은 값이다(§3.7 `CORE_CENTER_DIST_MAX_MM`). 이보다 잘게 재면 인식이
# 이미 "같은 자리" 로 판정한 층이 여기서만 다른 평면으로 갈린다.
_FLOOR_CELL_MM = 500.0
# 격자가 이보다 커지면 셀을 두 배로 키운다. 큰 건물에서 검사가 배치보다 오래
# 걸리면 아무도 안 돌린다.
_FLOOR_CELL_BUDGET = 200_000

_TYPICAL_FLOOR_IOU_MIN = 0.95      # §13.5 명시값
_ROOM_AREA_TOLERANCE = 0.05        # §13.5 명시값
_HEAD_COUNT_TOLERANCE = 0.10       # §13.5 명시값
_EXEMPT_AREA_MAX_RATIO = 0.30      # §13.5 명시값
# 이론값이 이보다 적으면 헤드 하나 차이가 이미 10% 를 넘는다. 그런 실에서는
# 검사가 항상 플래그를 내므로, 통과할 수 없는 검사를 돌리지 않는다.
_HEAD_COUNT_MIN_SAMPLE = 10


def _record(code: str, status: str, message: str, **extra) -> dict:
    return {"code": code, "status": status, "message": message, **extra}


# ────────────────────────────────────────────────────────────────────────────
# 기준층 교차검증 (최우선)
# ────────────────────────────────────────────────────────────────────────────

def _polygon(room) -> list:
    poly = getattr(room, "polygon", None) or []
    return [(float(p[0]), float(p[1])) for p in poly]


def _raster_cell(rooms) -> float:
    """모든 층이 **같은 격자**를 써야 IoU 가 뜻을 갖는다. 셀은 한 번만 정한다."""
    cell = _FLOOR_CELL_MM
    while True:
        total = 0.0
        for poly in rooms:
            if len(poly) < 3:
                continue
            xs = [p[0] for p in poly]
            ys = [p[1] for p in poly]
            total += ((max(xs) - min(xs)) / cell + 1) * ((max(ys) - min(ys)) / cell + 1)
        if total <= _FLOOR_CELL_BUDGET:
            return cell
        cell *= 2.0


def _cells(polygons: list, cell: float) -> set:
    out: set = set()
    for poly in polygons:
        if len(poly) < 3:
            continue
        xs = [p[0] for p in poly]
        ys = [p[1] for p in poly]
        for i in range(math.floor(min(xs) / cell), math.floor(max(xs) / cell) + 1):
            for j in range(math.floor(min(ys) / cell), math.floor(max(ys) / cell) + 1):
                if (i, j) not in out and point_in_polygon(
                        ((i + 0.5) * cell, (j + 0.5) * cell), poly):
                    out.add((i, j))
    return out


def iou(a: set, b: set) -> float:
    union = len(a | b)
    return len(a & b) / union if union else 0.0


def group_floors_by_similarity(rooms, *, iou_min: float = _TYPICAL_FLOOR_IOU_MIN):
    """평면이 같은 층끼리 묶는다. `[[층, 층], ...]` — 혼자인 층은 나오지 않는다.

    폴리곤 불리언 대신 격자로 재는 이유는 그것으로 충분하기 때문이다. 0.95 를
    가르는 데 필요한 분해능은 셀 하나보다 훨씬 크고, 불리언은 자기교차 폴리곤에서
    조용히 틀린다 — C170 이 내놓는 실 폴리곤이 늘 단순하다는 보장은 없다.
    """
    by_floor: dict[str, list] = {}
    for room in rooms:
        poly = _polygon(room)
        if len(poly) >= 3:
            by_floor.setdefault(str(getattr(room, "floor", "") or ""), []).append(poly)
    if len(by_floor) < 2:
        return []

    cell = _raster_cell([p for polys in by_floor.values() for p in polys])
    raster = {floor: _cells(polys, cell) for floor, polys in by_floor.items()}

    groups: list[list[str]] = []
    for floor in sorted(raster):
        for group in groups:
            if iou(raster[floor], raster[group[0]]) >= iou_min:
                group.append(floor)
                break
        else:
            groups.append([floor])
    return [g for g in groups if len(g) > 1]


def check_typical_floor_consistency(rooms, head_counts: dict) -> list:
    """같은 평면인데 층별 헤드 수가 다르면 플래그. 정답 없이 도는 공짜 검증기다."""
    floor_of = {str(getattr(r, "id", "")): str(getattr(r, "floor", "") or "")
                for r in rooms}
    per_floor: dict[str, int] = {}
    for room_id, count in head_counts.items():
        per_floor[floor_of.get(room_id, "")] = \
            per_floor.get(floor_of.get(room_id, ""), 0) + count

    groups = group_floors_by_similarity(rooms)
    if not groups:
        return [_record("TYPICAL_FLOOR_MISMATCH", "unverified",
                        "평면이 같은 층이 둘 이상 없어 교차검증할 수 없습니다.")]

    out = []
    for group in groups:
        counts = [per_floor.get(f, 0) for f in group]
        if len(set(counts)) > 1:
            out.append(_record(
                "TYPICAL_FLOOR_MISMATCH", "flag",
                f"평면이 같은 층 {', '.join(group)} 의 헤드 수가 다릅니다: "
                f"{', '.join(str(n) for n in counts)}.",
                floors=list(group), counts=counts))
        else:
            out.append(_record(
                "TYPICAL_FLOOR_MISMATCH", "pass",
                f"평면이 같은 층 {', '.join(group)} 의 헤드 수가 {counts[0]}개로 같습니다.",
                floors=list(group), counts=counts))
    return out


# ────────────────────────────────────────────────────────────────────────────
# 면적·개수
# ────────────────────────────────────────────────────────────────────────────

def _area_m2(room) -> float:
    area = float(getattr(room, "area_m2", 0.0) or 0.0)
    if area > 0.0:
        return area
    poly = _polygon(room)
    return polygon_area(poly) / 1e6 if len(poly) >= 3 else 0.0


def check_room_area(rooms, gross_floor_area_m2: float | None) -> dict:
    """실 폴리곤 면적 합 vs 건축개요 연면적."""
    total = sum(_area_m2(r) for r in rooms)
    if not gross_floor_area_m2:
        return _record("ROOM_AREA_MISMATCH", "unverified",
                       f"실 면적 합은 {total:,.1f}㎡ 입니다. 대조할 연면적이 없습니다.",
                       room_area_m2=round(total, 1))
    gross = float(gross_floor_area_m2)
    ratio = abs(total - gross) / gross
    status = "flag" if ratio > _ROOM_AREA_TOLERANCE else "pass"
    return _record(
        "ROOM_AREA_MISMATCH", status,
        f"실 면적 합 {total:,.1f}㎡ / 연면적 {gross:,.1f}㎡ — 차이 {ratio * 100:.1f}%.",
        room_area_m2=round(total, 1), gross_floor_area_m2=round(gross, 1),
        deviation=round(ratio, 4))


def check_head_count(rooms, head_counts: dict, constraints) -> dict:
    """헤드 총수 vs 방호면적 ÷ (S×L).

    설치제외 실은 분자에서도 분모에서도 뺀다. 한쪽만 빼면 제외가 많은 건물이
    무조건 미달로 읽힌다.
    """
    exempt = {str(getattr(r, "id", "")) for r in rooms
              if getattr(r, "head_exempt", False)}
    protected = sum(_area_m2(r) for r in rooms if str(getattr(r, "id", "")) not in exempt)
    spacing = float(constraints.head_spacing_square_m)
    theoretical = protected / (spacing * spacing) if spacing > 0 else 0.0
    actual = sum(n for rid, n in head_counts.items() if rid not in exempt)

    if theoretical < _HEAD_COUNT_MIN_SAMPLE:
        return _record(
            "HEAD_COUNT_DEVIATION", "unverified",
            f"이론값이 {theoretical:.1f}개뿐이라 ±10% 로는 판정할 수 없습니다.",
            head_count=actual, theoretical=round(theoretical, 1))

    ratio = (actual - theoretical) / theoretical
    status = "flag" if abs(ratio) > _HEAD_COUNT_TOLERANCE else "pass"
    return _record(
        "HEAD_COUNT_DEVIATION", status,
        f"헤드 {actual}개 / 이론값 {theoretical:.1f}개 (간격 {spacing:.3f}m) — "
        f"{ratio * 100:+.1f}%.",
        head_count=actual, theoretical=round(theoretical, 1),
        deviation=round(ratio, 4))


def check_exempt_area(rooms, gross_floor_area_m2: float | None) -> dict:
    """마스킹(설치제외) 면적 비율. 제외가 과하면 배치가 아니라 판단이 틀린 것이다."""
    exempt = sum(_area_m2(r) for r in rooms if getattr(r, "head_exempt", False))
    basis = float(gross_floor_area_m2) if gross_floor_area_m2 else sum(
        _area_m2(r) for r in rooms)
    if basis <= 0:
        return _record("EXEMPT_AREA_EXCESS", "unverified",
                       "면적이 0 이라 제외 비율을 낼 수 없습니다.")
    ratio = exempt / basis
    status = "flag" if ratio > _EXEMPT_AREA_MAX_RATIO else "pass"
    return _record(
        "EXEMPT_AREA_EXCESS", status,
        f"설치제외 {exempt:,.1f}㎡ / "
        f"{'연면적' if gross_floor_area_m2 else '실 면적 합'} {basis:,.1f}㎡ "
        f"— {ratio * 100:.1f}%.",
        exempt_area_m2=round(exempt, 1), ratio=round(ratio, 4),
        basis="gross_floor_area" if gross_floor_area_m2 else "room_area_sum")


# ────────────────────────────────────────────────────────────────────────────

def head_counts_of(layouts) -> dict:
    """`{실 id: 헤드 수}`. 세션은 JSON 을 왕복하므로 dict 도 그대로 받는다."""
    out: dict[str, int] = {}
    for layout in layouts or ():
        if isinstance(layout, dict):
            out[str(layout.get("room_id"))] = len(layout.get("heads") or [])
        else:
            out[str(layout.room_id)] = len(layout.heads)
    return out


def run_checks(rooms, layouts, constraints, *,
               gross_floor_area_m2: float | None = None) -> dict:
    """§13.5 검사표. `flags` 는 `checks` 중 `flag` 인 것만 추린 것이다."""
    counts = head_counts_of(layouts)
    checks = [
        *check_typical_floor_consistency(rooms, counts),
        check_room_area(rooms, gross_floor_area_m2),
        check_head_count(rooms, counts, constraints),
        check_exempt_area(rooms, gross_floor_area_m2),
        # 배관 총연장 ÷ 헤드 수는 C5 가 끝나야 잴 수 있고, 대조할 "과거 도면
        # 회귀 범위" 가 아직 없다. 범위를 지어내면 그 범위가 근거가 되어 버린다.
        _record("PIPE_LENGTH_PER_HEAD", "unverified",
                "배관 총연장은 C5 이후에만 잴 수 있고, 대조할 회귀 범위가 아직 없습니다."),
    ]
    return {"checks": checks,
            "flags": [c for c in checks if c["status"] == "flag"]}
