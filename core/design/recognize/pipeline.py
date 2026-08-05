# -*- coding: utf-8 -*-
"""C1 인식 사슬 — 지시서 §3.0 (C130 → C190 → building.json).

[문서정합] §2 파일 레이아웃에 이 모듈이 없다. §3.0 은 C130~C190 을 이어 붙여
`building_draft.json` 을 낸다고만 쓰고 그 자리를 지정하지 않았다. 라우트에 두면
Flask 없이는 사슬을 돌릴 수 없어 부록 B 가 요구하는 도면 벤치마크가 불가능해진다.
인식 계층에 두고, 라우트는 스트리밍과 세션 저장만 맡는다.

[문서정합] §11.1 목록에 없는 `building` 메시지를 마지막에 하나 더 낸다. 생성기가
진행을 흘리면서 산출물도 돌려줘야 하는데, §11.1 의 `result` 는 세션 URL 을 실어야
해서 라우트가 만든다. `building` 은 라우트가 가로채는 내부 전달용이다.

**인식기가 채운 값은 제안이지 확정이 아니다**(§3.6). 그래서 이 모듈이 실에 넣는
`name`·`use` 의 provenance 는 `C180` 이고, GATE 가 확정으로 치는 `GATE`/`default`
가 아니다. 이걸 어기면 결손이 사라져 사람이 볼 기회가 없어진다.
"""
from __future__ import annotations

import time
from dataclasses import dataclass, field

from ..schema import BuildingDraft, Core, Room
from . import arch_category as C
from . import core_detect as K
from . import geom_stats as G
from . import opening_close as O
from . import params as P
from . import room_faces as F
from . import room_label as L
from . import wall_centerline as W

PROV_C180 = "C180"

WALL_SOURCE_C140 = "C140"
WALL_SOURCE_OPERATOR = "operator"

_MM_PER_M = 1000.0
_MM2_PER_M2 = 1_000_000.0


@dataclass
class Stage:
    """한 단계의 소요 시간과 한 줄 요약. 벤치마크 리포트의 행이 된다."""

    name: str
    seconds: float
    summary: str
    provenance: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {"name": self.name, "seconds": round(self.seconds, 3),
                "summary": self.summary, "provenance": list(self.provenance)}


def normalize_bbox(bbox: dict | None) -> dict | None:
    """`x_min` 표기(inspect)와 `minx` 표기(C130)를 하나로 맞춘다.

    두 표기가 실제로 쓰이고 있어 경계에서 한 번 흡수한다. 여기서 안 맞추면
    단위 제안이 조용히 `None` 이 되고 모든 mm 임계값이 함께 틀어진다.
    """
    if not bbox:
        return None
    out = {}
    for key, alts in (("minx", ("minx", "x_min")), ("miny", ("miny", "y_min")),
                      ("maxx", ("maxx", "x_max")), ("maxy", ("maxy", "y_max"))):
        for alt in alts:
            if bbox.get(alt) is not None:
                out[key] = float(bbox[alt])
                break
        else:
            return None
    return out


def recognize(entities, bbox=None, *, floor: str = "1F", session_id: str = "",
              wall_layers=(), source: dict | None = None):
    """C130→C190 을 잇고 §11.1 메시지를 순서대로 낸다. 마지막이 `building`.

    `wall_layers` 는 사람이 WALL 로 확정한 레이어다. C140 이 WALL 을 못 찾는
    도면이 실제로 있어(실측 2/2) 이 통로가 없으면 사슬이 C150 에서 멈춘다.
    """
    entities = list(entities)
    bbox = normalize_bbox(bbox)
    stages: list[Stage] = []
    clock = _Clock()

    unit = G.suggest_unit_to_mm(bbox)
    unit_to_mm = (unit or {}).get("unit_to_mm") or P.UNIT_TO_MM_DEFAULT

    fps = G.fingerprints(entities, unit_to_mm=unit_to_mm, bbox=bbox)
    verdicts = {fp.name: C.evaluate(fp) for fp in fps}
    stages.append(clock.lap("C130/C140", f"레이어 {len(fps)}개 지문·판정"))
    yield {"type": "fingerprint", "unit": unit,
           "layers": [_layer_msg(fp, verdicts[fp.name]) for fp in fps]}

    picked = _wall_layers(fps, verdicts, wall_layers)
    doors = {fp.name for fp in fps if verdicts[fp.name].category == C.DOOR}
    if not picked.names:
        # 조용히 실 0개를 돌려주면 화면은 "인식했는데 방이 없다" 로 읽는다.
        yield {"type": "phase", "phase": "recognize", "blocked": "no_wall_layer",
               "message": "WALL 로 볼 레이어가 없어 중심선화를 시작하지 못했다.",
               "candidates": _wall_candidates(fps, verdicts)}
        yield _building_msg(BuildingDraft(session_id=session_id,
                                          source=dict(source or {}),
                                          scale=_scale(unit)),
                            stages, unit, picked, blocked="no_wall_layer")
        return

    yield {"type": "phase", "phase": "recognize",
           "wall_layers": sorted(picked.names), "wall_source": picked.source,
           "door_layers": sorted(doors)}

    peaks = [p for fp in fps if fp.name in picked.names for p in fp.offset_peaks_mm]
    wall_lines = [e["p"] for e in entities if e["t"] == "L" and e["l"] in picked.names]
    centerlines = W.build_centerlines(wall_lines, offset_peaks_mm=peaks,
                                      unit_to_mm=unit_to_mm)
    stages.append(clock.lap(
        "C150", f"중심선 {len(centerlines.centerlines)}개 "
                f"(짝지음 {centerlines.paired_ratio:.2f}, {centerlines.wall_repr})",
        centerlines.provenance))
    yield {"type": "centerlines", "wall_repr": centerlines.wall_repr,
           "paired_ratio": round(centerlines.paired_ratio, 3),
           "relaxed": centerlines.relaxed,
           "edges": [c.to_dict() for c in centerlines.centerlines],
           "provenance": list(centerlines.provenance)}

    door_ents = [e for e in entities if e["t"] == "A" and e["l"] in doors]
    closure = O.close_openings(centerlines, door_ents, unit_to_mm=unit_to_mm)
    stages.append(clock.lap(
        "C160", f"가상 간선 {len(closure.virtual_edges)}개 "
                f"/ 열린 끝점 {closure.open_endpoints}개", closure.provenance))
    yield {"type": "virtual_edges", "open_endpoints": closure.open_endpoints,
           "edges": [e.to_dict() for e in closure.virtual_edges],
           "provenance": list(closure.provenance)}

    faces = F.build_faces(centerlines, closure, bbox_area_mm2=_bbox_area(bbox, unit_to_mm))
    stages.append(clock.lap("C170", f"실 {len(faces.faces)}개", faces.provenance))

    room_text = {fp.name for fp in fps if verdicts[fp.name].category == C.ROOM_TEXT}
    texts = [e for e in entities if e["t"] == "T" and e["l"] in room_text]
    labels = L.assign_labels(faces.faces, texts,
                             [e for e in entities if e["t"] == "L"],
                             unit_to_mm=unit_to_mm)
    named = sum(1 for lb in labels.labels if lb.name)
    stages.append(clock.lap("C180", f"{len(faces.faces)}개 중 {named}개 명명",
                            labels.provenance))

    cores = K.detect_cores([faces.faces], names=[[lb.name for lb in labels.labels]])
    stages.append(clock.lap("C190", f"코어 후보 {len(cores.candidates)}개",
                            cores.provenance))

    rooms = [_room(i, face, labels.labels[i], floor) for i, face in enumerate(faces.faces)]
    yield {"type": "rooms", "rooms": [r.to_dict() for r in rooms],
           "dropped": dict(faces.dropped), "unassigned_texts": labels.unassigned_texts,
           "provenance": list(faces.provenance) + list(labels.provenance)}

    core_list = [_core(i, cand, faces.faces, floor)
                 for i, cand in enumerate(cores.candidates)]
    yield {"type": "cores", "cores": [c.to_dict() for c in core_list],
           "provenance": list(cores.provenance)}

    draft = BuildingDraft(
        session_id=session_id,
        source={**(source or {}), "floors": [{"label": floor}]},
        scale=_scale(unit), rooms=rooms, cores=core_list,
        virtual_edges=[e.to_dict() for e in closure.virtual_edges])
    yield _building_msg(draft, stages, unit, picked, faces=faces, labels=labels)


# ────────────────────────────────────────────────────────────────────────────
# 산출물 조립
# ────────────────────────────────────────────────────────────────────────────

def _room(index: int, face, label, floor: str) -> Room:
    room = Room(
        id=f"R-{floor}-{index:03d}", floor=floor,
        polygon=[[round(x, 1), round(y, 1)] for x, y in face.polygon],
        area_m2=round(face.area_m2, 2),
        perimeter_m=round(face.perimeter_mm / _MM_PER_M, 2),
        virtual_edge_ratio=round(face.virtual_ratio, 3),
        flags=list(face.flags),
        confidence={"polygon": round(face.confidence, 2)},
        provenance={"polygon": "C170"})
    if label.name:
        room.name = label.name
        room.confidence["name"] = round(label.confidence, 2)
        room.provenance["name"] = PROV_C180
    if label.use_hint:
        # 확정이 아니다. provenance 가 GATE/default 가 아니라서 §4 결손에 남는다.
        room.use = label.use_hint
        room.confidence["use"] = round(label.use_confidence, 2)
        room.provenance["use"] = PROV_C180
    return room


def _core(index: int, cand, faces, floor: str) -> Core:
    face = faces[cand.face_index] if cand.face_index < len(faces) else None
    return Core(
        id=f"CORE-{floor}-{index:03d}", kind=cand.kind.upper(),
        polygon=([[round(x, 1), round(y, 1)] for x, y in face.polygon]
                 if face is not None else []),
        area_m2=round(cand.area_m2, 2), confidence=round(cand.confidence, 2),
        confirmed=None)


def _building_msg(draft, stages, unit, picked, *, faces=None, labels=None,
                  blocked: str | None = None) -> dict:
    counts = {"rooms": len(draft.rooms), "cores": len(draft.cores),
              "virtual_edges": len(draft.virtual_edges)}
    if faces is not None:
        counts["dropped"] = dict(faces.dropped)
    if labels is not None:
        counts["unassigned_texts"] = labels.unassigned_texts
    return {"type": "building", "draft": draft.to_dict(), "counts": counts,
            "unit": unit, "wall_layers": sorted(picked.names),
            "wall_source": picked.source, "blocked": blocked,
            "stages": [s.to_dict() for s in stages],
            "seconds": round(sum(s.seconds for s in stages), 3)}


def _scale(unit) -> dict:
    if not unit:
        return {"unit_to_mm": None, "basis": "도면 범위를 알 수 없어 단위를 제안하지 못했다",
                "confidence": 0.0}
    return dict(unit)


def _bbox_area(bbox, unit_to_mm: float) -> float | None:
    if not bbox:
        return None
    return ((bbox["maxx"] - bbox["minx"]) * (bbox["maxy"] - bbox["miny"])
            * unit_to_mm * unit_to_mm)


# ────────────────────────────────────────────────────────────────────────────
# WALL 확정
# ────────────────────────────────────────────────────────────────────────────

@dataclass
class _Picked:
    names: set
    source: str


def _wall_layers(fps, verdicts, override) -> _Picked:
    """사람이 고른 것이 C140 판정을 이긴다 (§4 — 확정은 사람이 한다)."""
    known = {fp.name for fp in fps}
    chosen = {name for name in override if name in known}
    if chosen:
        return _Picked(chosen, WALL_SOURCE_OPERATOR)
    return _Picked({fp.name for fp in fps if verdicts[fp.name].category == C.WALL},
                   WALL_SOURCE_C140)


def _wall_candidates(fps, verdicts) -> list:
    """WALL 이 없을 때 사람에게 보여줄 차선. 평행쌍 비율이 높은 순이다."""
    ranked = sorted(fps, key=lambda fp: (-fp.parallel_pair_ratio, -fp.n_entities))
    return [{"name": fp.name, "n_entities": fp.n_entities,
             "category": verdicts[fp.name].category,
             "parallel_pair_ratio": round(fp.parallel_pair_ratio, 2),
             "offset_peaks_mm": list(fp.offset_peaks_mm),
             "near_misses": list(verdicts[fp.name].near_misses)}
            for fp in ranked[:P.WALL_CANDIDATE_LIMIT] if fp.offset_peaks_mm]


def _layer_msg(fp, verdict) -> dict:
    return {"name": fp.name, "n_entities": fp.n_entities, "sampled": fp.sampled,
            **{k: v for k, v in verdict.to_dict().items() if k != "layer"}}


class _Clock:
    def __init__(self):
        self._t = time.perf_counter()

    def lap(self, name: str, summary: str, provenance=()) -> Stage:
        now = time.perf_counter()
        stage = Stage(name=name, seconds=now - self._t, summary=summary,
                      provenance=list(provenance))
        self._t = now
        return stage
