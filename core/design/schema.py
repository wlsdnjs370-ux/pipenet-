# -*- coding: utf-8 -*-
"""building.json (schema `fncadnet.building/1`) 의 자료구조 — 지시서 §10.1.

값과 `None` 을 반드시 구분한다. `ceiling.has_finish = False` 는 "반자가 없다"는
주장이고 `None` 은 "아직 모른다"이다. 둘을 같은 칸에 담으면 GATE 가 결손을 찾지
못하고, 찾지 못한 결손은 기본값으로 조용히 채워져 하류 전부를 틀린 기준 위에서
통과시킨다(금지사항 G9).

[문서정합] 지시서 §2 는 이 파일에 Constraints·Design 도 두라고 하지만, Constraints
는 PR-1 에서 `deterministic/constraints.py` 에 있고(NFTC 값의 단일 진실 출처를
세션 스키마와 섞을 수 없다), Design(C3~C5 산출물)은 PR-7~9 에서 만든다.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

BUILDING_SCHEMA = "fncadnet.building/1"


def _opt_float(value: Any) -> float | None:
    return None if value is None else float(value)


def _opt_bool(value: Any) -> bool | None:
    return None if value is None else bool(value)


@dataclass
class Ceiling:
    """반자(마감 천장)와 슬래브. `has_finish` 가 헤드 종류와 신축배관을 가른다."""

    has_finish: bool | None = None
    finish_height_mm: float | None = None
    slab_height_mm: float | None = None

    @classmethod
    def from_dict(cls, d: dict | None) -> "Ceiling":
        d = d or {}
        return cls(
            has_finish=_opt_bool(d.get("has_finish")),
            finish_height_mm=_opt_float(d.get("finish_height_mm")),
            slab_height_mm=_opt_float(d.get("slab_height_mm")),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "has_finish": self.has_finish,
            "finish_height_mm": self.finish_height_mm,
            "slab_height_mm": self.slab_height_mm,
        }


@dataclass
class Room:
    id: str
    floor: str
    polygon: list = field(default_factory=list)
    area_m2: float = 0.0
    perimeter_m: float = 0.0
    virtual_edge_ratio: float = 0.0
    name: str | None = None
    use: str | None = None
    ceiling: Ceiling = field(default_factory=Ceiling)
    ambient_temp_max_c: float | None = None
    special_hazard: str | None = None
    head_exempt: bool = False
    confidence: dict = field(default_factory=dict)
    provenance: dict = field(default_factory=dict)
    flags: list = field(default_factory=list)

    @classmethod
    def from_dict(cls, d: dict) -> "Room":
        return cls(
            id=str(d["id"]),
            floor=str(d.get("floor") or ""),
            polygon=list(d.get("polygon") or []),
            area_m2=float(d.get("area_m2") or 0.0),
            perimeter_m=float(d.get("perimeter_m") or 0.0),
            virtual_edge_ratio=float(d.get("virtual_edge_ratio") or 0.0),
            name=d.get("name"),
            use=d.get("use"),
            ceiling=Ceiling.from_dict(d.get("ceiling")),
            ambient_temp_max_c=_opt_float(d.get("ambient_temp_max_c")),
            special_hazard=d.get("special_hazard"),
            head_exempt=bool(d.get("head_exempt", False)),
            confidence=dict(d.get("confidence") or {}),
            provenance=dict(d.get("provenance") or {}),
            flags=list(d.get("flags") or []),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id, "floor": self.floor, "polygon": self.polygon,
            "area_m2": self.area_m2, "perimeter_m": self.perimeter_m,
            "virtual_edge_ratio": self.virtual_edge_ratio,
            "name": self.name, "use": self.use, "ceiling": self.ceiling.to_dict(),
            "ambient_temp_max_c": self.ambient_temp_max_c,
            "special_hazard": self.special_hazard, "head_exempt": self.head_exempt,
            "confidence": self.confidence, "provenance": self.provenance,
            "flags": self.flags,
        }


@dataclass
class Core:
    """입상관 후보. 단층 도면만으로는 확정할 수 없어 `confirmed` 를 사람이 채운다."""

    id: str
    kind: str = "SHAFT"
    polygon: list = field(default_factory=list)
    area_m2: float = 0.0
    vertical_aligned_floors: list = field(default_factory=list)
    confidence: float = 0.0
    confirmed: bool | None = None

    @classmethod
    def from_dict(cls, d: dict) -> "Core":
        return cls(
            id=str(d["id"]), kind=str(d.get("kind") or "SHAFT"),
            polygon=list(d.get("polygon") or []),
            area_m2=float(d.get("area_m2") or 0.0),
            vertical_aligned_floors=list(d.get("vertical_aligned_floors") or []),
            confidence=float(d.get("confidence") or 0.0),
            confirmed=_opt_bool(d.get("confirmed")),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "id": self.id, "kind": self.kind, "polygon": self.polygon,
            "area_m2": self.area_m2,
            "vertical_aligned_floors": self.vertical_aligned_floors,
            "confidence": self.confidence, "confirmed": self.confirmed,
        }


@dataclass
class Obstacles:
    """`status` 는 `none`/`partial`/`complete`. "미확보" 를 고르는 것도 확정이다."""

    status: str | None = None
    beams: list = field(default_factory=list)
    ducts: list = field(default_factory=list)
    lights: list = field(default_factory=list)

    @classmethod
    def from_dict(cls, d: dict | None) -> "Obstacles":
        d = d or {}
        return cls(
            status=d.get("status"), beams=list(d.get("beams") or []),
            ducts=list(d.get("ducts") or []), lights=list(d.get("lights") or []),
        )

    def to_dict(self) -> dict[str, Any]:
        return {"status": self.status, "beams": self.beams,
                "ducts": self.ducts, "lights": self.lights}


@dataclass
class BuildingFacts:
    """도면 밖에 있는 건물 사실. C2 가 기준개수·수원·R 을 여기서 끌어간다."""

    floors_total: int | None = None
    floors_underground: int | None = None
    structure: str | None = None
    use: str | None = None
    is_underground_arcade: bool = False

    @classmethod
    def from_dict(cls, d: dict | None) -> "BuildingFacts":
        d = d or {}
        total = d.get("floors_total")
        under = d.get("floors_underground")
        return cls(
            floors_total=None if total is None else int(total),
            floors_underground=None if under is None else int(under),
            structure=d.get("structure"), use=d.get("use"),
            is_underground_arcade=bool(d.get("is_underground_arcade", False)),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "floors_total": self.floors_total,
            "floors_underground": self.floors_underground,
            "structure": self.structure, "use": self.use,
            "is_underground_arcade": self.is_underground_arcade,
        }


@dataclass
class Gate:
    """`edits` 는 사람이 인식 결과를 어떻게 고쳤는지 — C1 개선의 학습 데이터다."""

    passed: bool = False
    passed_at: str | None = None
    operator: str | None = None
    unresolved: list = field(default_factory=list)
    edits: list = field(default_factory=list)

    @classmethod
    def from_dict(cls, d: dict | None) -> "Gate":
        d = d or {}
        return cls(
            passed=bool(d.get("passed", False)), passed_at=d.get("passed_at"),
            operator=d.get("operator"), unresolved=list(d.get("unresolved") or []),
            edits=list(d.get("edits") or []),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "passed": self.passed, "passed_at": self.passed_at,
            "operator": self.operator, "unresolved": self.unresolved,
            "edits": self.edits,
        }


@dataclass
class BuildingDraft:
    """C1 산출물이자 GATE 편집 대상. `version` 은 세션 계층이 관리한다."""

    session_id: str = ""
    source: dict = field(default_factory=dict)
    scale: dict = field(default_factory=dict)
    building: BuildingFacts = field(default_factory=BuildingFacts)
    rooms: list = field(default_factory=list)
    virtual_edges: list = field(default_factory=list)
    cores: list = field(default_factory=list)
    obstacles: Obstacles = field(default_factory=Obstacles)
    gate: Gate = field(default_factory=Gate)

    @classmethod
    def from_dict(cls, d: dict) -> "BuildingDraft":
        return cls(
            session_id=str(d.get("session_id") or ""),
            source=dict(d.get("source") or {}), scale=dict(d.get("scale") or {}),
            building=BuildingFacts.from_dict(d.get("building")),
            rooms=[Room.from_dict(r) for r in (d.get("rooms") or [])],
            virtual_edges=list(d.get("virtual_edges") or []),
            cores=[Core.from_dict(c) for c in (d.get("cores") or [])],
            obstacles=Obstacles.from_dict(d.get("obstacles")),
            gate=Gate.from_dict(d.get("gate")),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "schema": BUILDING_SCHEMA,
            "session_id": self.session_id,
            "source": self.source, "scale": self.scale,
            "building": self.building.to_dict(),
            "rooms": [r.to_dict() for r in self.rooms],
            "virtual_edges": self.virtual_edges,
            "cores": [c.to_dict() for c in self.cores],
            "obstacles": self.obstacles.to_dict(),
            "gate": self.gate.to_dict(),
        }

    def room(self, room_id: str) -> Room | None:
        for r in self.rooms:
            if r.id == room_id:
                return r
        return None

    def core(self, core_id: str) -> Core | None:
        for c in self.cores:
            if c.id == core_id:
                return c
        return None

    def floor_height_mm(self, floor_label: str) -> float | None:
        """`source.floors[].height_mm` — 천장고 기본값의 유일한 근거.

        [문서정합] 지시서 §4.2 는 천장고의 기본값을 "층고" 라고 적었지만 §10.1
        building.json 스키마에 층고 필드가 없다. C1 이 층고를 실을 수 있도록
        `source.floors[]` 에 선택 필드로 두고, 없으면 기본값을 만들지 않고
        결손으로 남긴다.
        """
        for entry in (self.source.get("floors") or []):
            if entry.get("label") == floor_label:
                return _opt_float(entry.get("height_mm"))
        return None
