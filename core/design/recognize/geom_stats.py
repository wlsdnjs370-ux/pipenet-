# -*- coding: utf-8 -*-
"""C130 — 레이어별 기하 지문 수집 (지시서 §3.1).

레이어 **이름 사전에 의존하지 않기 위한** 통계다. 건축도면은 설계사마다 레이어
규약이 완전히 달라서 (`A-WALL` / `벽체` / `WALL-1` / `건축-벽` / 심지어 `0`)
이름으로 판정하면 처음 보는 도면에서 조용히 전부 틀린다.

입력은 C110/C120 이 만든 **캔버스 엔티티**다. 새로 파싱하지 않는다 — 화면에
보이는 것과 인식이 보는 것이 달라지면 검수자가 무엇을 확인한 것인지 알 수 없다.

  {"t":"L","l":레이어,"p":[x1,y1,x2,y2]}      LINE
  {"t":"A","l":...,"c":[cx,cy],"r":r,"a":[시작각,끝각]}   ARC
  {"t":"C","l":...,"c":[cx,cy],"r":r}         CIRCLE
  {"t":"PL","l":...,"p":[[x,y],...]}          폴리라인 계열
  {"t":"T","l":...,"p":[x,y],"v":문자열}       TEXT 계열
  {"t":"I","l":...,"p":[x,y],"n":블록명}       INSERT
  {"t":"H"/"S","l":...,"p":[[x,y],...]}       HATCH / SOLID

[문서정합] §3.1 의 `LayerFingerprint` 는 필드 10개만 적었지만 §3.2 판정표는 거기
없는 값(ARC 반경 분포, 폐합 면적, 계단 다발, 초장선 비율)을 요구한다. 표를 따르되
필드를 늘렸다. 명세를 줄이지 않고 구현을 맞췄다.

[문서정합] 캔버스 엔티티에는 LWPOLYLINE 의 `closed` 플래그도 `DIMENSION` 타입도
없다. 폐합은 첫점-끝점 간극으로 판정하고(§`CLOSED_GAP_TOL_MM`), DIM 판정은
`text_numeric_ratio` 만으로 한다. 원본 라우트의 payload 를 늘리지 않기 위한
절충이며, 놓치는 경우(닫힘 플래그만 있고 첫점≠끝점인 사각형)는 4-LINE 사각형
탐색이 일부 메운다.
"""
from __future__ import annotations

import math
import random
import statistics
from dataclasses import dataclass, field, asdict

from . import params as P
from .spatial import (
    NodeIndex, SegmentGrid, angle_deg, angle_diff, centroid, length,
    overlap_ratio, perp_offset, polygon_area,
)

_MM2_PER_M2 = 1_000_000.0


@dataclass
class LayerFingerprint:
    """한 레이어의 기하 통계. 전부 mm 기준이다."""

    name: str
    n_entities: int = 0
    sampled: bool = False
    type_hist: dict = field(default_factory=dict)
    len_median_mm: float = 0.0
    len_p90_mm: float = 0.0
    parallel_pair_ratio: float = 0.0
    offset_peaks_mm: list = field(default_factory=list)
    closed_shape_count: int = 0
    closed_repeat_score: float = 0.0
    arc_attach_ratio: float = 0.0
    text_numeric_ratio: float = 0.0
    text_len_median: float = 0.0
    grid_alignment_score: float = 0.0
    door_radius_ratio: float = 0.0
    closed_area_median_m2: float = 0.0
    small_closed_count: int = 0
    stair_bundle_max: int = 0
    stair_gap_cv: float = 0.0
    long_line_ratio: float = 0.0

    def to_dict(self) -> dict:
        return asdict(self)


# ────────────────────────────────────────────────────────────────────────────
# 단위
# ────────────────────────────────────────────────────────────────────────────

def suggest_unit_to_mm(bbox: dict | None) -> dict | None:
    """도면 대각 길이로 단위를 **제안**한다. 적용은 하지 않는다.

    모든 임계값이 mm 기준이라 단위가 틀리면 판정 전체가 함께 틀린다. 그래서
    자동으로 갈아끼우지 않고 후보와 근거만 내놓는다 — 확정은 GATE 가 한다.
    """
    if not bbox:
        return None
    try:
        diag = math.hypot(float(bbox["maxx"]) - float(bbox["minx"]),
                          float(bbox["maxy"]) - float(bbox["miny"]))
    except (KeyError, TypeError, ValueError):
        return None
    if diag <= 0:
        return None
    for unit_to_mm, label in ((1.0, "mm"), (10.0, "cm"), (1000.0, "m")):
        if P.PLAUSIBLE_DIAG_MIN_MM <= diag * unit_to_mm <= P.PLAUSIBLE_DIAG_MAX_MM:
            return {"unit_to_mm": unit_to_mm, "basis": f"도면 대각 {diag:.1f} → {label} 가정",
                    "confidence": 0.6 if unit_to_mm == 1.0 else 0.4}
    return {"unit_to_mm": None, "basis": f"도면 대각 {diag:.1f} 가 어떤 단위로도 층 평면 범위를 벗어남",
            "confidence": 0.0}


# ────────────────────────────────────────────────────────────────────────────
# 진입점
# ────────────────────────────────────────────────────────────────────────────

def fingerprints(entities: list, *, unit_to_mm: float = P.UNIT_TO_MM_DEFAULT,
                 bbox: dict | None = None) -> list[LayerFingerprint]:
    """레이어별 지문. 입력 순서와 무관하게 레이어 이름 순으로 돌려준다."""
    by_layer: dict[str, list] = {}
    for ent in entities:
        by_layer.setdefault(str(ent.get("l") or ""), []).append(ent)

    diag_mm = _bbox_diag_mm(bbox, unit_to_mm)
    return [_one_layer(name, items, unit_to_mm, diag_mm)
            for name, items in sorted(by_layer.items())]


def _bbox_diag_mm(bbox: dict | None, unit_to_mm: float) -> float:
    if not bbox:
        return 0.0
    try:
        return math.hypot(float(bbox["maxx"]) - float(bbox["minx"]),
                          float(bbox["maxy"]) - float(bbox["miny"])) * unit_to_mm
    except (KeyError, TypeError, ValueError):
        return 0.0


def _one_layer(name: str, items: list, unit_to_mm: float,
               diag_mm: float) -> LayerFingerprint:
    total = len(items)
    sampled = total > P.LAYER_SAMPLE_LIMIT
    if sampled:
        # 부록 C.1 — 표본은 고정 시드로 뽑는다. 매번 달라지면 같은 도면에서
        # 지문이 흔들려 무엇 때문에 판정이 바뀌었는지 알 수 없다.
        items = random.Random(P.SAMPLE_SEED).sample(items, P.LAYER_SAMPLE_LIMIT)

    fp = LayerFingerprint(name=name, n_entities=total, sampled=sampled)
    for ent in items:
        key = str(ent.get("t") or "?")
        fp.type_hist[key] = fp.type_hist.get(key, 0) + 1

    lines = _lines_mm(items, unit_to_mm)
    arcs = [e for e in items if e.get("t") == "A"]
    texts = [str(e.get("v") or "") for e in items if e.get("t") == "T"]

    _fill_line_stats(fp, lines, diag_mm)
    _fill_parallel_stats(fp, lines)
    _fill_stair_stats(fp, lines)
    _fill_arc_stats(fp, arcs, lines, unit_to_mm)
    _fill_closed_stats(fp, items, lines, unit_to_mm)
    fp.text_numeric_ratio = _numeric_ratio(texts)
    stripped = [t.strip() for t in texts if t.strip()]
    if stripped:
        fp.text_len_median = statistics.median(len(t) for t in stripped)
    return fp


# ────────────────────────────────────────────────────────────────────────────
# 개별 통계
# ────────────────────────────────────────────────────────────────────────────

def _lines_mm(items: list, unit_to_mm: float) -> list[tuple]:
    out = []
    for ent in items:
        if ent.get("t") != "L":
            continue
        p = ent.get("p") or []
        if len(p) < 4:
            continue
        seg = (float(p[0]) * unit_to_mm, float(p[1]) * unit_to_mm,
               float(p[2]) * unit_to_mm, float(p[3]) * unit_to_mm)
        if length(seg) > 0.0:
            out.append(seg)
    return out


def _fill_line_stats(fp: LayerFingerprint, lines: list, diag_mm: float) -> None:
    if not lines:
        return
    lens = sorted(length(s) for s in lines)
    fp.len_median_mm = statistics.median(lens)
    fp.len_p90_mm = lens[min(len(lens) - 1, int(len(lens) * 0.9))]
    if diag_mm > 0.0:
        cut = diag_mm * P.LONG_LINE_DIAG_RATIO
        fp.long_line_ratio = sum(1 for v in lens if v >= cut) / len(lens)


def _fill_parallel_stats(fp: LayerFingerprint, lines: list) -> None:
    """평행쌍 비율과 오프셋 peak — WALL 판정의 1차 근거.

    격자를 각도 버킷별로 따로 둔다. 하나로 합치면 가구 레이어처럼 짧은 선이
    수만 개 몰린 셀에서 후보가 그대로 O(n^2) 로 불어나 실 도면에서만 멈춘다.
    각도로 먼저 가르면 대부분의 후보가 조회 전에 사라진다.
    """
    if len(lines) < 2:
        return
    angles = [angle_deg(s) for s in lines]
    n_buckets = int(180.0 / P.ANGLE_BUCKET_DEG)

    grids: dict[int, SegmentGrid] = {}
    cells: list[tuple] = []
    buckets: list[int] = []
    for i, seg in enumerate(lines):
        b = int(angles[i] / P.ANGLE_BUCKET_DEG) % n_buckets
        grid = grids.get(b)
        if grid is None:
            grid = grids[b] = SegmentGrid(P.PARALLEL_OFFSET_MAX_MM)
        walked = grid.walk(seg)
        grid.add(i, walked)
        cells.append(walked)
        buckets.append(b)

    paired: set[int] = set()
    offsets: list[float] = []
    for i, seg in enumerate(lines):
        # 이웃 버킷까지 봐야 179.9° 와 0.1° 처럼 버킷 경계를 사이에 둔 짝을 놓치지 않는다.
        candidates: set[int] = set()
        for d in (-1, 0, 1):
            grid = grids.get((buckets[i] + d) % n_buckets)
            if grid is not None:
                candidates |= grid.lookup(cells[i])
        for j in candidates:
            if j <= i:
                continue
            if angle_diff(angles[i], angles[j]) > P.PARALLEL_ANGLE_TOL_DEG:
                continue
            other = lines[j]
            offset = perp_offset(seg, (other[0] + other[2]) * 0.5,
                                 (other[1] + other[3]) * 0.5)
            if not (P.PARALLEL_OFFSET_MIN_MM <= offset <= P.PARALLEL_OFFSET_MAX_MM):
                continue
            if overlap_ratio(seg, other) < P.PARALLEL_OVERLAP_MIN_RATIO:
                continue
            paired.add(i)
            paired.add(j)
            offsets.append(offset)

    fp.parallel_pair_ratio = len(paired) / len(lines)
    fp.offset_peaks_mm = _offset_peaks(offsets)


def _offset_peaks(offsets: list) -> list:
    """10mm bin 히스토그램에서 전체의 8% 이상을 차지하는 상위 bin 들.

    peak 을 찾는 데만 bin 을 쓰고 값은 그 bin 안 오프셋의 중앙값으로 돌려준다.
    bin 중심을 그대로 쓰면 150mm 벽이 155mm 로 보고되고, C150 이 그 값을
    ±15mm 로 다시 맞춰 볼 때 오차가 그대로 얹힌다.
    """
    if not offsets:
        return []
    hist: dict[int, list[float]] = {}
    for value in offsets:
        hist.setdefault(int(value // P.OFFSET_HIST_BIN_MM), []).append(value)
    floor = len(offsets) * P.OFFSET_PEAK_MIN_SHARE
    top = sorted(((b, v) for b, v in hist.items() if len(v) >= floor),
                 key=lambda kv: (-len(kv[1]), kv[0]))[:P.OFFSET_PEAK_MAX]
    return [round(statistics.median(v), 1) for _, v in top]


def _fill_stair_stats(fp: LayerFingerprint, lines: list) -> None:
    """등간격 평행 단선 다발 — 계단 디딤판.

    각도 버킷 안에서 수직 위치로 정렬한 뒤 간격이 고른 최장 구간을 찾는다.
    길이가 고른지도 함께 본다(미검증) — 벽 두 겹도 등간격이지만 계단 디딤판은
    길이까지 나란하다.
    """
    if len(lines) < P.STAIR_BUNDLE_MIN:
        return
    buckets: dict[int, list[tuple[float, float]]] = {}
    for seg in lines:
        ang = angle_deg(seg)
        rad = math.radians(ang)
        mx, my = (seg[0] + seg[2]) * 0.5, (seg[1] + seg[3]) * 0.5
        offset = -mx * math.sin(rad) + my * math.cos(rad)
        buckets.setdefault(int(ang // P.ANGLE_BUCKET_DEG), []).append((offset, length(seg)))

    best_run, best_cv = 0, 0.0
    for group in buckets.values():
        if len(group) < P.STAIR_BUNDLE_MIN:
            continue
        group.sort()
        gaps = [group[i + 1][0] - group[i][0] for i in range(len(group) - 1)]
        start = 0
        for i in range(len(gaps) + 1):
            same = (i < len(gaps) and gaps[start] > 0.0
                    and abs(gaps[i] - gaps[start]) <= gaps[start] * P.STAIR_GAP_TOL_RATIO)
            if same:
                continue
            run_gaps = gaps[start:i]
            run_lens = [ln for _, ln in group[start:i + 1]]
            if len(run_gaps) + 1 > best_run and _cv(run_lens) <= P.STAIR_LEN_CV_MAX:
                best_run, best_cv = len(run_gaps) + 1, _cv(run_gaps)
            start = i
    fp.stair_bundle_max = best_run
    fp.stair_gap_cv = best_cv


def _fill_arc_stats(fp: LayerFingerprint, arcs: list, lines: list,
                    unit_to_mm: float) -> None:
    """ARC 가 LINE 끝점에 붙은 비율 — 문 손잡이쪽 회전 궤적의 지문."""
    if not arcs:
        return
    ends = NodeIndex(P.ARC_ATTACH_TOL_MM)
    for x1, y1, x2, y2 in lines:
        ends.add(x1, y1)
        ends.add(x2, y2)

    attached = 0
    in_range = 0
    for arc in arcs:
        c = arc.get("c") or [0.0, 0.0]
        r = float(arc.get("r") or 0.0) * unit_to_mm
        a = arc.get("a") or [0.0, 0.0]
        cx, cy = float(c[0]) * unit_to_mm, float(c[1]) * unit_to_mm
        if P.DOOR_ARC_RADIUS_MIN_MM <= r <= P.DOOR_ARC_RADIUS_MAX_MM:
            in_range += 1
        for deg in (float(a[0]), float(a[1])):
            rad = math.radians(deg)
            if ends.find(cx + r * math.cos(rad), cy + r * math.sin(rad)) is not None:
                attached += 1
                break
    fp.arc_attach_ratio = attached / len(arcs)
    fp.door_radius_ratio = in_range / len(arcs)


def _fill_closed_stats(fp: LayerFingerprint, items: list, lines: list,
                       unit_to_mm: float) -> None:
    """폐합 도형의 개수·면적·반복성·격자 정렬 — COLUMN/SHAFT 판정의 근거."""
    polys: list[list] = []
    for ent in items:
        if ent.get("t") not in ("PL", "H", "S"):
            continue
        pts = [(float(p[0]) * unit_to_mm, float(p[1]) * unit_to_mm)
               for p in (ent.get("p") or []) if len(p) >= 2]
        if len(pts) < 3:
            continue
        if ent.get("t") == "PL" and math.hypot(pts[0][0] - pts[-1][0],
                                               pts[0][1] - pts[-1][1]) > P.CLOSED_GAP_TOL_MM:
            continue
        polys.append(pts)
    polys.extend(_line_quads(lines))

    fp.closed_shape_count = len(polys)
    if not polys:
        return
    areas = [polygon_area(p) / _MM2_PER_M2 for p in polys]
    areas = [a for a in areas if a > 0.0]
    if not areas:
        return
    fp.closed_area_median_m2 = statistics.median(areas)
    fp.small_closed_count = sum(1 for a in areas if a <= P.SMALL_CLOSED_AREA_MAX_M2)
    # 반복성 — 크기가 고를수록 1 에 가깝다. 기둥은 같은 단면이 반복된다.
    fp.closed_repeat_score = 1.0 / (1.0 + _cv(areas))
    fp.grid_alignment_score = _grid_alignment([centroid(p) for p in polys])


def _line_quads(lines: list) -> list[list]:
    """4개의 LINE 이 이루는 사각형. 폐합 플래그를 못 읽는 만큼을 여기서 메운다.

    각 노드에서 이웃 쌍을 기록해 두고, 같은 이웃 쌍을 공유하는 노드가 둘이면
    4-사이클이다. 차수가 과도한 노드는 건너뛴다 — 도면 잡음에서 조합이 폭발한다.
    """
    if len(lines) < 4:
        return []
    index = NodeIndex(P.NODE_SNAP_TOL_MM)
    adj: dict[int, set[int]] = {}
    for x1, y1, x2, y2 in lines:
        a, b = index.add(x1, y1), index.add(x2, y2)
        if a == b:
            continue
        adj.setdefault(a, set()).add(b)
        adj.setdefault(b, set()).add(a)

    seen_pairs: dict[tuple[int, int], list[int]] = {}
    for node, neighbours in adj.items():
        if len(neighbours) > P.QUAD_MAX_DEGREE:
            continue
        ordered = sorted(neighbours)
        for i, u in enumerate(ordered):
            for v in ordered[i + 1:]:
                seen_pairs.setdefault((u, v), []).append(node)

    quads: list[list] = []
    emitted: set[frozenset] = set()
    for (u, v), mids in seen_pairs.items():
        for i, m1 in enumerate(mids):
            for m2 in mids[i + 1:]:
                key = frozenset((u, v, m1, m2))
                if len(key) != 4 or key in emitted:
                    continue
                emitted.add(key)
                quads.append([index.points[m1], index.points[u],
                              index.points[m2], index.points[v]])
    return quads


def _grid_alignment(centers: list) -> float:
    """중심 좌표가 같은 통심에 늘어선 정도. 0~1."""
    if len(centers) < 2:
        return 0.0
    shares = []
    for axis in (0, 1):
        values = sorted(c[axis] for c in centers)
        clustered = 0
        start = 0
        for i in range(len(values) + 1):
            if i < len(values) and values[i] - values[start] <= P.GRID_ALIGN_TOL_MM:
                continue
            if i - start >= 2:
                clustered += i - start
            start = i
        shares.append(clustered / len(values))
    return sum(shares) / len(shares)


def _numeric_ratio(texts: list) -> float:
    """숫자·치수 기호만으로 이루어진 TEXT 의 비율."""
    values = [t.strip() for t in texts if t.strip()]
    if not values:
        return 0.0
    allowed = set(P.NUMERIC_TEXT_CHARS)
    numeric = sum(1 for t in values
                  if any(ch.isdigit() for ch in t) and set(t) <= allowed)
    return numeric / len(values)


def _cv(values: list) -> float:
    """변동계수. 값이 하나뿐이거나 평균이 0 이면 0(=고르다)."""
    if len(values) < 2:
        return 0.0
    mean = statistics.fmean(values)
    if mean == 0.0:
        return 0.0
    return statistics.pstdev(values) / abs(mean)
