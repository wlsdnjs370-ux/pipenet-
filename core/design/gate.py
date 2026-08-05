# -*- coding: utf-8 -*-
"""GATE — 인간 확정 게이트의 결손 판정과 반영 — 지시서 §4.

**인식기가 채운 값은 제안이지 확정이 아니다.** 신뢰도가 0.95 여도 자동 확정하지
않는다(§3.6). 그래서 결손 판정은 값의 유무가 아니라 `provenance` 로 한다 —
`GATE`(사람) 또는 `default`(근거를 함께 남긴 기본값)만 확정으로 친다.

질의는 한 지점에 모은다. 파이프라인 중간에서 사람을 계속 부르면 자동화가 아니라
방해이므로, C1 이 끝나면 결손 전체를 한 번에 내주고 이후는 무인으로 끝까지 돈다.

[문서정합] §4.3 응답 예시는 그룹의 대상 목록 키를 `rooms` 라고 적었으나, 건물
사실·코어 확정처럼 실이 아닌 대상이 섞이므로 `targets` 로 통일하고 `scope` 를
함께 실었다. 소비자(PR-3 화면)는 아직 없어 계약을 깨지 않는다.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Callable

from .deterministic import nftc_tables as T
from .schema import BuildingDraft, Room

PROV_GATE = "GATE"
PROV_DEFAULT = "default"
_CONFIRMED_PROVENANCE = (PROV_GATE, PROV_DEFAULT)

OBSTACLE_STATUSES = ("none", "partial", "complete")
STRUCTURES = ("내화구조", "비내화구조")
SPECIAL_HAZARDS = ("무대부", "랙크식창고", "특수가연물", "전기차충전구역")


class GateNotPassed(Exception):
    """C2 이후 단계의 진입 차단. UI 잠금은 시각 효과일 뿐 여기가 실제 방어선이다."""

    def __init__(self, unresolved: list[str]):
        self.unresolved = unresolved
        super().__init__("인간 확정 게이트를 통과해야 진행할 수 있습니다.")


@dataclass(frozen=True)
class GateField:
    field: str
    label: str
    scope: str                      # room | building | obstacles | core
    required: bool
    options: tuple | None = None
    bulk_apply: tuple = ()
    note: str = ""


ROOM_FIELDS: tuple[GateField, ...] = (
    GateField("use", "용도", "room", True, T.OCCUPANCY_USES, ("floor",),
              "기준개수·R·헤드종류를 전부 결정하는 단일 실패점"),
    GateField("ceiling.has_finish", "반자 유무", "room", True, (True, False),
              ("floor", "use"), "도면에 없다. 헤드 종류와 신축배관을 가른다"),
    GateField("ceiling.finish_height_mm", "반자고(mm)", "room", True, None,
              ("floor", "use"), "반자가 있을 때만 필수"),
    GateField("ceiling.slab_height_mm", "천장고(슬래브, mm)", "room", True, None,
              ("floor", "use"), "기준개수 8m 판정의 대리 지표. 층고를 기본값으로 쓴다"),
    GateField("ambient_temp_max_c", "최고 주위온도(℃)", "room", False, None,
              ("floor", "use"), "표시온도 구간을 정한다"),
    GateField("special_hazard", "특수 위험", "room", False, (None,) + SPECIAL_HAZARDS,
              ("floor", "use"), "무대부·랙크식·특수가연물·EV충전"),
)

BUILDING_FIELDS: tuple[GateField, ...] = (
    GateField("building.floors_total", "건물 층수", "building", True, None, (),
              "수원 20/40/60분 분기"),
    GateField("building.structure", "구조(내화 여부)", "building", True, STRUCTURES, (),
              "R 2.1 vs 2.3"),
    GateField("building.use", "대상물 용도", "building", True, T.OCCUPANCY_USES, (),
              "기준개수 표 2.1.1.1 의 분기 입력"),
    GateField("obstacles.status", "장애물 상태", "obstacles", True, OBSTACLE_STATUSES, (),
              "'미확보(none)' 를 고르는 것도 확정이다"),
    GateField("confirmed", "코어 확정", "core", True, (True, False), (),
              "입상관 후보. 단층 도면만으로는 확정할 수 없다"),
)

_ALL_FIELDS = {f.field: f for f in ROOM_FIELDS + BUILDING_FIELDS}


# ────────────────────────────────────────────────────────────────────────────
# 값 읽기·쓰기 (알려진 필드만 — 점 표기를 setattr 로 풀지 않는다)
# ────────────────────────────────────────────────────────────────────────────

_ROOM_GET: dict[str, Callable[[Room], Any]] = {
    "use": lambda r: r.use,
    "ceiling.has_finish": lambda r: r.ceiling.has_finish,
    "ceiling.finish_height_mm": lambda r: r.ceiling.finish_height_mm,
    "ceiling.slab_height_mm": lambda r: r.ceiling.slab_height_mm,
    "ambient_temp_max_c": lambda r: r.ambient_temp_max_c,
    "special_hazard": lambda r: r.special_hazard,
}


def _set_room(room: Room, field: str, value: Any) -> None:
    if field == "use":
        room.use = None if value is None else str(value)
    elif field == "ceiling.has_finish":
        room.ceiling.has_finish = None if value is None else bool(value)
    elif field == "ceiling.finish_height_mm":
        room.ceiling.finish_height_mm = None if value is None else float(value)
    elif field == "ceiling.slab_height_mm":
        room.ceiling.slab_height_mm = None if value is None else float(value)
    elif field == "ambient_temp_max_c":
        room.ambient_temp_max_c = None if value is None else float(value)
    elif field == "special_hazard":
        room.special_hazard = None if value is None else str(value)
    else:
        raise KeyError(field)


def _set_building(draft: BuildingDraft, field: str, value: Any) -> None:
    if field == "building.floors_total":
        draft.building.floors_total = int(value)
    elif field == "building.structure":
        draft.building.structure = str(value)
    elif field == "building.use":
        draft.building.use = str(value)
    elif field == "obstacles.status":
        if value not in OBSTACLE_STATUSES:
            raise ValueError(f"장애물 상태는 {OBSTACLE_STATUSES} 중 하나여야 합니다")
        draft.obstacles.status = str(value)
    else:
        raise KeyError(field)


def _building_value(draft: BuildingDraft, field: str) -> Any:
    return {
        "building.floors_total": draft.building.floors_total,
        "building.structure": draft.building.structure,
        "building.use": draft.building.use,
        "obstacles.status": draft.obstacles.status,
    }[field]


# ────────────────────────────────────────────────────────────────────────────
# 결손 판정
# ────────────────────────────────────────────────────────────────────────────

def _room_field_applies(room: Room, field: str) -> bool:
    """반자고는 반자가 있다고 확정된 실에만 묻는다."""
    if field == "ceiling.finish_height_mm":
        return room.ceiling.has_finish is True
    return True


def _room_confirmed(room: Room, field: str) -> bool:
    if _ROOM_GET[field](room) is None:
        return False
    return room.provenance.get(field) in _CONFIRMED_PROVENANCE


def unresolved(draft: BuildingDraft) -> list[str]:
    """확정되지 않은 필수 항목 키 목록. 비어야 게이트를 통과할 수 있다."""
    out: list[str] = []
    for room in draft.rooms:
        for spec in ROOM_FIELDS:
            if not spec.required or not _room_field_applies(room, spec.field):
                continue
            if not _room_confirmed(room, spec.field):
                out.append(f"{room.id}.{spec.field}")
    for spec in BUILDING_FIELDS:
        if spec.scope == "core":
            continue
        if _building_value(draft, spec.field) is None:
            out.append(spec.field)
    for core in draft.cores:
        if core.confirmed is None:
            out.append(f"{core.id}.confirmed")
    return out


def _room_suggestion(room: Room, field: str) -> dict | None:
    value = _ROOM_GET[field](room)
    if value is None or room.provenance.get(field) in _CONFIRMED_PROVENANCE:
        return None
    return {"value": value, "confidence": float(room.confidence.get(field, 0.0))}


def _room_confidence(room: Room, field: str) -> float:
    return float(room.confidence.get(field, 0.0))


def gate_items(draft: BuildingDraft) -> dict[str, Any]:
    """§4.3 결손 항목 응답. 검수 우선순위가 곧 정렬 순서라 신뢰도 낮은 순으로 낸다."""
    groups: list[dict[str, Any]] = []
    total = 0

    for spec in ROOM_FIELDS:
        targets = [r for r in draft.rooms
                   if _room_field_applies(r, spec.field) and not _room_confirmed(r, spec.field)]
        if not targets:
            continue
        targets.sort(key=lambda r: (_room_confidence(r, spec.field), r.id))
        suggestion = {r.id: s for r in targets
                      if (s := _room_suggestion(r, spec.field)) is not None}
        groups.append(_group(spec, [r.id for r in targets], suggestion))
        if spec.required:
            total += len(targets)

    for spec in BUILDING_FIELDS:
        if spec.scope == "core":
            targets = [c.id for c in draft.cores if c.confirmed is None]
            suggestion = {c.id: {"value": c.kind, "confidence": c.confidence}
                          for c in draft.cores if c.confirmed is None}
        else:
            if _building_value(draft, spec.field) is not None:
                continue
            targets, suggestion = [spec.field], {}
        if not targets:
            continue
        groups.append(_group(spec, targets, suggestion))
        if spec.required:
            total += len(targets)

    return {"ok": True, "total": total, "groups": groups}


def _group(spec: GateField, targets: list[str], suggestion: dict) -> dict[str, Any]:
    group: dict[str, Any] = {
        "field": spec.field, "label": spec.label, "scope": spec.scope,
        "required": spec.required, "targets": targets, "count": len(targets),
        "bulk_apply": list(spec.bulk_apply), "note": spec.note,
    }
    if spec.options is not None:
        group["options"] = list(spec.options)
    if suggestion:
        group["suggestion"] = suggestion
    return group


# ────────────────────────────────────────────────────────────────────────────
# 반영
# ────────────────────────────────────────────────────────────────────────────

def apply_defaults(draft: BuildingDraft) -> list[dict[str, Any]]:
    """근거가 있는 기본값만 채운다. 반환은 감사 로그에 남길 변경 목록.

    지금은 천장고 하나다 — `source.floors[].height_mm` 가 있을 때만 층고를 쓰고,
    없으면 채우지 않고 결손으로 남긴다. 근거 없는 채움은 G9 위반이다.
    """
    applied: list[dict[str, Any]] = []
    for room in draft.rooms:
        if _room_confirmed(room, "ceiling.slab_height_mm"):
            continue
        height = draft.floor_height_mm(room.floor)
        if height is None:
            continue
        room.ceiling.slab_height_mm = height
        room.provenance["ceiling.slab_height_mm"] = PROV_DEFAULT
        applied.append({"room": room.id, "field": "ceiling.slab_height_mm",
                        "value": height, "basis": "source.floors[].height_mm"})
    return applied


def apply_values(draft: BuildingDraft, values: dict[str, dict]) -> list[dict[str, Any]]:
    """`{대상 id: {필드: 값}}` 을 확정으로 반영. 반환은 감사 로그에 남길 변경 목록.

    대상 id 는 실 id, 코어 id, 또는 건물 사실 필드명(`building.floors_total` 등)이다.
    제안값과 다른 확정(용도 override)은 감리 질의 1순위라 `suggested` 를 함께 남긴다.
    """
    changes: list[dict[str, Any]] = []
    for target_id, fields in values.items():
        if not isinstance(fields, dict):
            raise ValueError(f"{target_id}: 필드 매핑이어야 합니다")
        for field, value in fields.items():
            spec = _ALL_FIELDS.get(field)
            if spec is None:
                raise ValueError(f"알 수 없는 항목: {field}")
            changes.append(_apply_one(draft, target_id, spec, value))
    return changes


def _apply_one(draft: BuildingDraft, target_id: str, spec: GateField, value: Any) -> dict[str, Any]:
    if spec.scope == "room":
        room = draft.room(target_id)
        if room is None:
            raise ValueError(f"없는 실: {target_id}")
        before = _ROOM_GET[spec.field](room)
        suggested = _room_suggestion(room, spec.field)
        _set_room(room, spec.field, value)
        room.provenance[spec.field] = PROV_GATE
        return {"target": room.id, "field": spec.field, "from": before, "to": value,
                "suggested": suggested}

    if spec.scope == "core":
        core = draft.core(target_id)
        if core is None:
            raise ValueError(f"없는 코어: {target_id}")
        before = core.confirmed
        core.confirmed = bool(value)
        return {"target": core.id, "field": "confirmed", "from": before, "to": core.confirmed,
                "suggested": {"value": core.kind, "confidence": core.confidence}}

    before = _building_value(draft, spec.field)
    _set_building(draft, spec.field, value)
    return {"target": spec.field, "field": spec.field, "from": before,
            "to": _building_value(draft, spec.field), "suggested": None}


def require_gate(draft: BuildingDraft) -> None:
    if not draft.gate.passed:
        raise GateNotPassed(draft.gate.unresolved or unresolved(draft))
