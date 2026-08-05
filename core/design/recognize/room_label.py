# -*- coding: utf-8 -*-
"""C180 — 실명 텍스트 귀속 (지시서 §3.6).

폴리곤에 이름을 붙이고, 이름에서 용도를 **추정**한다. §3.6 이 못 박은 대로 용도는
신뢰도가 아무리 높아도 자동 확정하지 않는다 — 용도가 NFTC 기준개수를 바꾸므로
사람이 GATE 에서 확정해야 한다.

[문서정합] §3.6 3항의 "면적 표기(숫자+㎡)를 제외" 는 ㎡ 가 붙은 것만 말하지만,
실 도면에는 단위 없이 숫자만 적힌 치수·실번호도 폴리곤 안에 들어온다. 숫자와
기호로만 이루어진 텍스트도 함께 뺀다 — 실 이름이 "3,600" 인 방은 없다.

[문서정합] 귀속에 붙일 신뢰도가 §3.6 에 없다. 폴리곤 안에서 직접 찾은 것과
지시선을 따라간 것을 같은 무게로 둘 수 없어 둘로 나눴다(params, `# 미검증`).
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

from . import params as P
from .spatial import point_in_polygon

Point = tuple[float, float]

INSIDE = "inside"
LEADER = "leader"

# §3.6 용도 추정 표. 실명 → NFTC 특정소방대상물 구분. 추정일 뿐이다.
USE_HINTS = {
    "업무시설": ["사무실", "사무", "OFFICE", "업무"],
    "공동주택": ["거실", "침실", "안방", "주방", "세대"],
    "판매시설": ["매장", "판매", "SHOP", "STORE", "마트"],
    "숙박시설": ["객실", "ROOM", "숙박"],
    "의료시설": ["병실", "진료", "수술", "처치"],
    "노유자시설": ["보육", "요양", "노인"],
    "주차장": ["주차", "PARKING", "P.LOT"],
    "창고시설": ["창고", "WAREHOUSE", "저장"],
}

_AREA_UNITS = ("㎡", "M2", "M²", "SQM", "평")


@dataclass
class RoomLabel:
    """폴리곤 하나에 붙은 이름과 용도 추정."""

    face_index: int
    name: str | None
    source: str | None
    confidence: float
    use_hint: str | None = None
    use_confidence: float = 0.0
    candidates: list = field(default_factory=list)
    provenance: list = field(default_factory=list)

    @property
    def needs_input(self) -> bool:
        """이름을 못 붙였으면 GATE 가 사람에게 물어야 한다 (§3.6 4항)."""
        return self.name is None

    def to_dict(self) -> dict:
        return {
            "face_index": self.face_index,
            "name": self.name,
            "source": self.source,
            "confidence": round(self.confidence, 2),
            "use_hint": self.use_hint,
            "use_confidence": round(self.use_confidence, 2),
            "needs_input": self.needs_input,
            # §3.6 — 신뢰도 0.95 이상이어도 자동 확정 금지.
            "needs_confirm": True,
            "candidates": list(self.candidates),
            "provenance": list(self.provenance),
        }


@dataclass
class LabelResult:
    labels: list
    unassigned_texts: int
    provenance: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "labels": [lb.to_dict() for lb in self.labels],
            "unassigned_texts": self.unassigned_texts,
            "provenance": list(self.provenance),
        }


def assign_labels(faces, texts, lines=(), *, unit_to_mm: float = 1.0) -> LabelResult:
    """실 폴리곤에 ROOM_TEXT 를 귀속시킨다.

    `faces` 는 C170 의 `RoomFace` 목록, `texts` 와 `lines` 는 캔버스 엔티티다.
    `lines` 는 지시선 추적(§3.6 2항)에만 쓴다.
    """
    picked = [(_text_point(t, unit_to_mm), str(t.get("v") or "").strip())
              for t in texts if t.get("t") == "T"]
    picked = [(pt, v) for pt, v in picked if pt is not None and v]

    index = _FaceIndex(faces)
    per_face: dict[int, list] = {}
    unassigned: list[tuple] = []
    for point, value in picked:
        fid = index.locate(point)
        if fid is None:
            unassigned.append((point, value))
            continue
        per_face.setdefault(fid, []).append((value, INSIDE))

    leads = _leader_hits(unassigned, lines, index, unit_to_mm)
    for fid, value in leads:
        per_face.setdefault(fid, []).append((value, LEADER))

    labels = []
    for fid in range(len(faces)):
        labels.append(_pick(fid, per_face.get(fid, ())))

    named = sum(1 for lb in labels if lb.name)
    provenance = [
        f"실 {len(faces)}개 중 {named}개에 이름을 붙였다 "
        f"(직접 {sum(1 for lb in labels if lb.source == INSIDE)}, "
        f"지시선 {sum(1 for lb in labels if lb.source == LEADER)})"]
    left = len(unassigned) - len(leads)
    if left > 0:
        provenance.append(f"어느 실에도 못 붙인 텍스트 {left}개")
    return LabelResult(labels=labels, unassigned_texts=max(0, left),
                       provenance=provenance)


def _pick(face_index: int, entries) -> RoomLabel:
    """§3.6 3항 — 면적 표기를 뺀 나머지 중 가장 긴 것."""
    entries = list(entries)
    if not entries:
        return RoomLabel(face_index=face_index, name=None, source=None,
                         confidence=0.0,
                         provenance=["폴리곤 안팎에서 실명 텍스트를 못 찾았다"])

    usable = [(v, s) for v, s in entries if not is_area_annotation(v)]
    if not usable:
        return RoomLabel(
            face_index=face_index, name=None, source=None, confidence=0.0,
            candidates=[v for v, _ in entries],
            provenance=[f"텍스트 {len(entries)}개가 전부 면적·치수 표기였다"])

    # 직접 귀속을 지시선보다 앞세운다. 지시선은 옆 실을 가리켰을 수 있다.
    usable.sort(key=lambda e: (e[1] != INSIDE, -len(e[0]), e[0]))
    name, source = usable[0]
    confidence = P.CONF_LABEL_INSIDE if source == INSIDE else P.CONF_LABEL_LEADER
    use_hint = use_of(name)

    provenance = [f"{'폴리곤 안' if source == INSIDE else '지시선 추적'}에서 "
                  f"'{name}'"]
    if len(usable) > 1:
        provenance.append(f"후보 {len(usable)}개 중 가장 긴 것을 골랐다")
    if use_hint:
        provenance.append(f"실명으로 용도 '{use_hint}' 추정 — GATE 확정 필요")

    return RoomLabel(
        face_index=face_index, name=name, source=source, confidence=confidence,
        use_hint=use_hint,
        use_confidence=P.CONF_USE_HINT if use_hint else 0.0,
        candidates=[v for v, _ in usable], provenance=provenance)


def is_area_annotation(value: str) -> bool:
    """면적 표기이거나 숫자·기호뿐인 텍스트인가."""
    upper = value.upper()
    if any(unit in upper for unit in _AREA_UNITS) and any(c.isdigit() for c in value):
        return True
    return all(c in P.NUMERIC_TEXT_CHARS for c in value)


def use_of(name: str) -> str | None:
    """실명 → NFTC 특정소방대상물 구분 **추정**. 확정은 GATE 가 한다."""
    upper = name.upper()
    for use, hints in USE_HINTS.items():
        if any(hint.upper() in upper for hint in hints):
            return use
    return None


def _leader_hits(unassigned, lines, index, unit_to_mm: float) -> list:
    """§3.6 2항 — 텍스트에 한쪽 끝이 붙고 다른 끝이 폴리곤 안인 LINE.

    [문서정합] "반경 2000mm 이내의 LINE" 은 후보를 모으는 반경으로 읽는다. 반대쪽
    끝까지 2000mm 안에 있으라고 읽으면, 큰 실 밖에 이름을 적고 안으로 지시선을
    끈 도면에서 이름이 통째로 날아간다. 반대쪽 끝을 묶는 것은 거리가 아니라
    **폴리곤 안에 있어야 한다**는 조건이다.
    """
    if not unassigned or not lines:
        return []
    segs = []
    for ent in lines:
        if ent.get("t") != "L":
            continue
        x1, y1, x2, y2 = (float(v) * unit_to_mm for v in ent["p"][:4])
        segs.append((x1, y1, x2, y2))
    if not segs:
        return []

    cell = P.LEADER_SEARCH_RADIUS_MM
    grid: dict = {}
    for i, (x1, y1, x2, y2) in enumerate(segs):
        for x, y in ((x1, y1), (x2, y2)):
            grid.setdefault((math.floor(x / cell), math.floor(y / cell)), []).append(i)

    out = []
    for point, value in unassigned:
        cx, cy = math.floor(point[0] / cell), math.floor(point[1] / cell)
        best = None
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for i in grid.get((cx + dx, cy + dy), ()):
                    x1, y1, x2, y2 = segs[i]
                    for near, far in (((x1, y1), (x2, y2)), ((x2, y2), (x1, y1))):
                        if math.dist(near, point) > P.LEADER_TEXT_ATTACH_MM:
                            continue
                        fid = index.locate(far)
                        if fid is None:
                            continue
                        d = math.dist(near, point)
                        if best is None or d < best[0]:
                            best = (d, fid)
        if best is not None:
            out.append((best[1], value))
    return out


class _FaceIndex:
    """폴리곤 bbox 격자. 후보만 좁히고 판정은 실제 내부 판정으로 한다."""

    def __init__(self, faces):
        self.faces = list(faces)
        self.boxes = []
        self.cell = P.FACE_INDEX_CELL_MIN_MM
        self._cells: dict = {}
        for fid, face in enumerate(self.faces):
            xs = [p[0] for p in face.polygon]
            ys = [p[1] for p in face.polygon]
            box = (min(xs), min(ys), max(xs), max(ys))
            self.boxes.append(box)
            for kx in range(math.floor(box[0] / self.cell),
                            math.floor(box[2] / self.cell) + 1):
                for ky in range(math.floor(box[1] / self.cell),
                                math.floor(box[3] / self.cell) + 1):
                    self._cells.setdefault((kx, ky), []).append(fid)

    def locate(self, point) -> int | None:
        """점을 품은 폴리곤. 여럿이면 가장 작은 것 — 실 안의 실을 우선한다."""
        key = (math.floor(point[0] / self.cell), math.floor(point[1] / self.cell))
        best = None
        for fid in self._cells.get(key, ()):
            x0, y0, x1, y1 = self.boxes[fid]
            if not (x0 <= point[0] <= x1 and y0 <= point[1] <= y1):
                continue
            if not point_in_polygon(point, self.faces[fid].polygon):
                continue
            if best is None or self.faces[fid].area_m2 < self.faces[best].area_m2:
                best = fid
        return best


def _text_point(ent, unit_to_mm: float):
    pos = ent.get("p")
    if not pos or len(pos) < 2:
        return None
    return (float(pos[0]) * unit_to_mm, float(pos[1]) * unit_to_mm)
