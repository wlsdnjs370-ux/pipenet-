# -*- coding: utf-8 -*-
"""C140 — 레이어 카테고리 판정 (지시서 §3.2).

지문이 1순위, 이름은 +0.15 가산점이다. 이름만으로는 어떤 카테고리도 나오지
않는다 — 레이어 이름은 설계사마다 제각각이라 이름을 단독 근거로 삼는 순간
처음 보는 도면에서 조용히 전부 틀린다.

[문서정합] §3.2 의 시그니처는 `(category, confidence)` 두 값이다. 그대로 두되
`evaluate()` 를 함께 둔다. 작업 규칙 9 가 인식 결과에 `provenance` 를 요구하는데
두 값으로는 실을 자리가 없고, 무엇보다 **왜 그 카테고리가 아닌지**를 못 싣는다.
실 도면에서 걸린 게 그 경우다 — `ARCHI` 레이어는 평행쌍 0.81 에 벽 두께 peak 도
잡히는데 `len_median` 636mm 하나 때문에 WALL 이 되지 못한다. 결과만 보면 그냥
DOOR 로 보이고, 검수자는 한 조건이 스친 것을 알 방법이 없다.

[문서정합] 표의 12종 중 둘은 C140 에서 판정할 수 없다. 규칙을 지어내지 않고
`undecidable` 로 보고한다.
  - `WINDOW` — 조건이 "평행 3~4선 반복 **+ 벽 중심선 위에 위치**" 인데 중심선은
    C150 의 산출물이다. C140 은 C150 보다 먼저 돈다.
  - `BEAM` — 12종에 이름은 있으나 표에 지문 조건 행이 없다. 이름 힌트만 있고,
    이름 단독 판정은 §3.2 자신이 금지한다.

[문서정합] `STAIR` 행은 "등간격 평행 단선 ≥ 6개 다발" 이라는 절대 개수뿐이다.
선이 20,000개인 레이어에서는 등간격 6선 다발이 우연히 거의 항상 나온다 — 실
도면의 가장 큰 레이어 셋이 전부 STAIR 0.70 으로 잡혔다. 다발이 레이어에서
차지하는 몫(`STAIR_BUNDLE_MIN_SHARE`, 미검증)을 조건에 더했다.

[문서정합] `SHAFT` 행에는 면적 상한(6㎡)만 있고 하한이 없다. 실 도면에서 가구
조각 같은 0.001㎡ 폐합 수천 개가 통째로 SHAFT 로 잡혀서, `COLUMN` 행의 하한과
같은 값을 `SHAFT_AREA_MIN_M2` 로 두었다(미검증).

[문서정합] `DIM` 의 "DIMENSION 엔티티 존재" 는 쓸 수 없다. 캔버스 엔티티에
DIMENSION 타입이 아예 없어(C110/C120 이 렌더하지 않는다) 치수 문자열조차
도달하지 않는다. 실 도면 `A-DIM` 레이어가 `text_numeric_ratio 0.00` 으로 나온
이유다. `text_numeric_ratio ≥ 0.8` 만으로 판정하며, 이 절반짜리 조건으로는 실
도면의 치수 레이어를 대부분 놓친다. 캔버스 payload 확장은 별건이다.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from . import params as P
from .geom_stats import LayerFingerprint

WALL = "WALL"
DOOR = "DOOR"
WINDOW = "WINDOW"
COLUMN = "COLUMN"
STAIR = "STAIR"
SHAFT = "SHAFT"
ROOM_TEXT = "ROOM_TEXT"
DIM = "DIM"
FURNITURE = "FURNITURE"
GRID = "GRID"
BEAM = "BEAM"
OTHER = "OTHER"

CATEGORIES = (WALL, DOOR, WINDOW, COLUMN, STAIR, SHAFT,
              ROOM_TEXT, DIM, FURNITURE, GRID, BEAM, OTHER)

# §3.2 이름 힌트 사전 — 가산점 전용. `sprinkler_remote30_extractor.py` 의
# `DEFAULT_*_LAYER_KEYWORDS` 는 모듈 A 가 의존하므로 건드리지 않고 여기에 따로 둔다.
NAME_HINTS = {
    WALL:      ["WALL", "벽", "벽체", "A-WALL", "AR-WALL", "W-"],
    DOOR:      ["DOOR", "문", "출입", "A-DOOR", "DR-"],
    WINDOW:    ["WINDOW", "창", "창호", "A-GLAZ", "WIN"],
    COLUMN:    ["COL", "COLUMN", "기둥", "S-COL", "PILLAR"],
    STAIR:     ["STAIR", "계단", "STR", "A-STRS"],
    SHAFT:     ["PD", "P.D", "AD", "A.D", "SHAFT", "덕트", "PS", "샤프트", "EPS", "TPS"],
    ROOM_TEXT: ["ROOM", "실명", "NAME", "TEXT", "A-ANNO"],
    DIM:       ["DIM", "치수", "A-DIMS"],
    FURNITURE: ["FURN", "가구", "A-FURN", "집기"],
    GRID:      ["GRID", "통심", "A-GRID", "AXIS"],
    BEAM:      ["BEAM", "보", "GIRDER", "S-BEAM"],
}


@dataclass
class Verdict:
    """한 레이어의 판정 결과. 근거와 **탈락 사유**를 함께 싣는다."""

    layer: str
    category: str
    confidence: float
    provenance: list = field(default_factory=list)
    name_bonus: bool = False
    alternatives: list = field(default_factory=list)
    near_misses: list = field(default_factory=list)
    undecidable: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {"layer": self.layer, "category": self.category,
                "confidence": round(self.confidence, 3),
                "provenance": list(self.provenance), "name_bonus": self.name_bonus,
                "alternatives": [[c, round(v, 3)] for c, v in self.alternatives],
                "near_misses": list(self.near_misses),
                "undecidable": list(self.undecidable)}


def arch_category(fp: LayerFingerprint, name_hint: str | None) -> tuple[str, float]:
    """§3.2 의 시그니처. 근거가 필요하면 `evaluate()` 를 쓴다."""
    verdict = evaluate(fp, name_hint)
    return verdict.category, verdict.confidence


def evaluate(fp: LayerFingerprint, name_hint: str | None = None, *,
             multi_floor_repeat: bool | None = None) -> Verdict:
    """지문으로 카테고리를 고르고, 왜 그렇게 골랐는지를 함께 돌려준다.

    `multi_floor_repeat` 는 SHAFT 의 "다층 도면에서 같은 좌표 반복" 이다. 한 층만
    본 시점에는 알 수 없으므로 기본이 `None` 이고, 그때 SHAFT 신뢰도는 0.40 에서
    멈춘다. 모르는 것을 아는 것처럼 올리지 않는다.
    """
    hint_name = (name_hint if name_hint is not None else fp.name) or ""
    rules = _rules(fp, multi_floor_repeat)

    scored: list[tuple[float, str, list, bool]] = []
    near: list[str] = []
    for category, base, checks, notes in rules:
        failed = [text for text, ok in checks if not ok]
        if not failed:
            bonus = _name_matches(hint_name, category)
            scored.append((base + (P.NAME_HINT_BONUS if bonus else 0.0),
                           category, [t for t, _ in checks] + notes, bonus))
        elif len(failed) == 1 and len(checks) - 1 >= 2:
            # 두 조건짜리 규칙에서 하나 어긋난 것은 "스쳤다" 가 아니다. 그걸 다
            # 실으면 모든 레이어에 "STAIR: 다발 0개 ≥ 6" 이 붙어 근거가 묻힌다.
            near.append(f"{category}: {failed[0]}")

    if not scored:
        scored = _furniture_fallback(fp, hint_name)

    undecidable = [
        f"{WINDOW}: 조건의 '벽 중심선 위에 위치' 는 C150 산출물이라 C140 에서 판정 불가",
        f"{BEAM}: §3.2 표에 지문 조건 행이 없다 — 이름 단독 판정은 금지",
    ]
    if not scored:
        return Verdict(layer=fp.name, category=OTHER, confidence=P.OTHER_CONFIDENCE,
                       provenance=["어떤 지문 조건도 성립하지 않았다"],
                       near_misses=near, undecidable=undecidable)

    scored.sort(key=lambda s: (-s[0], s[1]))
    score, category, provenance, bonus = scored[0]
    return Verdict(layer=fp.name, category=category, confidence=min(1.0, score),
                   provenance=provenance, name_bonus=bonus,
                   alternatives=[(c, v) for v, c, _, _ in scored[1:]],
                   near_misses=near, undecidable=undecidable)


def _rules(fp: LayerFingerprint, multi_floor_repeat: bool | None) -> list:
    """§3.2 판정표를 그대로 옮긴 것.

    각 규칙은 `(카테고리, 기본신뢰도, [(설명, 통과여부)], [단서])` 다. 설명을
    문자열로 들고 다니는 이유는 provenance 와 near_miss 가 같은 문장을 써야 하기
    때문이다. 통과 사유와 탈락 사유가 다른 말로 적히면 검수자는 둘이 같은
    조건인지 알 수 없다. 단서는 판정을 가르지 않고 신뢰도의 근거만 적는다.
    """
    n_typed = sum(fp.type_hist.values())
    n_lines = fp.type_hist.get("L", 0)
    text_share = fp.type_hist.get("T", 0) / n_typed if n_typed else 0.0
    stair_share = fp.stair_bundle_max / n_lines if n_lines else 0.0
    shaft_multi = bool(multi_floor_repeat)
    return [
        (WALL, P.CONF_WALL, [
            (f"평행쌍 비율 {fp.parallel_pair_ratio:.2f} ≥ {P.WALL_PARALLEL_MIN_RATIO}",
             fp.parallel_pair_ratio >= P.WALL_PARALLEL_MIN_RATIO),
            (f"벽 두께 peak {len(fp.offset_peaks_mm)}개 ≥ 1",
             len(fp.offset_peaks_mm) >= 1),
            (f"선 길이 중앙값 {fp.len_median_mm:.0f}mm ≥ {P.WALL_LEN_MEDIAN_MIN_MM:.0f}mm",
             fp.len_median_mm >= P.WALL_LEN_MEDIAN_MIN_MM),
        ], []),
        (DOOR, P.CONF_DOOR, [
            (f"ARC 가 선 끝점에 붙은 비율 {fp.arc_attach_ratio:.2f} ≥ {P.DOOR_ARC_ATTACH_MIN_RATIO}",
             fp.arc_attach_ratio >= P.DOOR_ARC_ATTACH_MIN_RATIO),
            (f"문 반경 범위 비율 {fp.door_radius_ratio:.2f} ≥ {P.DOOR_RADIUS_MIN_RATIO}",
             fp.door_radius_ratio >= P.DOOR_RADIUS_MIN_RATIO),
        ], []),
        (COLUMN, P.CONF_COLUMN, [
            (f"폐합 반복성 {fp.closed_repeat_score:.2f} ≥ {P.COLUMN_REPEAT_MIN}",
             fp.closed_repeat_score >= P.COLUMN_REPEAT_MIN),
            (f"격자 정렬 {fp.grid_alignment_score:.2f} ≥ {P.COLUMN_GRID_ALIGN_MIN}",
             fp.grid_alignment_score >= P.COLUMN_GRID_ALIGN_MIN),
            (f"폐합 면적 중앙값 {fp.closed_area_median_m2:.2f}㎡ 가 "
             f"{P.COLUMN_AREA_MIN_M2}~{P.COLUMN_AREA_MAX_M2}㎡",
             P.COLUMN_AREA_MIN_M2 <= fp.closed_area_median_m2 <= P.COLUMN_AREA_MAX_M2),
        ], []),
        (STAIR, P.CONF_STAIR, [
            (f"등간격 평행 단선 다발 {fp.stair_bundle_max}개 ≥ {P.STAIR_BUNDLE_MIN}",
             fp.stair_bundle_max >= P.STAIR_BUNDLE_MIN),
            (f"간격 변동계수 {fp.stair_gap_cv:.2f} ≤ {P.STAIR_GAP_CV_MAX}",
             fp.stair_gap_cv <= P.STAIR_GAP_CV_MAX),
            (f"다발이 레이어 선의 {stair_share:.3f} ≥ {P.STAIR_BUNDLE_MIN_SHARE}",
             stair_share >= P.STAIR_BUNDLE_MIN_SHARE),
        ], []),
        (SHAFT, P.CONF_SHAFT_MULTIFLOOR if shaft_multi else P.CONF_SHAFT_SINGLE, [
            (f"소형 폐합 {fp.small_closed_count}개 ≥ 1", fp.small_closed_count >= 1),
            (f"폐합 면적 중앙값 {fp.closed_area_median_m2:.2f}㎡ 가 "
             f"{P.SHAFT_AREA_MIN_M2}~{P.SMALL_CLOSED_AREA_MAX_M2}㎡",
             P.SHAFT_AREA_MIN_M2 <= fp.closed_area_median_m2 <= P.SMALL_CLOSED_AREA_MAX_M2),
        ], ["다층 도면에서 같은 좌표 반복 확인됨" if shaft_multi else
            f"다층 반복을 확인하지 못해 신뢰도를 {P.CONF_SHAFT_SINGLE} 로 제한"]),
        (DIM, P.CONF_DIM, [
            (f"숫자 전용 TEXT 비율 {fp.text_numeric_ratio:.2f} ≥ {P.DIM_NUMERIC_MIN_RATIO}",
             fp.text_numeric_ratio >= P.DIM_NUMERIC_MIN_RATIO),
        ], ["DIMENSION 엔티티는 캔버스에 없어 조건의 나머지 절반은 보지 못했다"]),
        (ROOM_TEXT, P.CONF_ROOM_TEXT, [
            (f"TEXT 비중 {text_share:.2f} ≥ {P.ROOM_TEXT_SHARE_MIN}",
             text_share >= P.ROOM_TEXT_SHARE_MIN),
            (f"숫자 전용 TEXT 비율 {fp.text_numeric_ratio:.2f} ≤ {P.ROOM_TEXT_NUMERIC_MAX_RATIO}",
             fp.text_numeric_ratio <= P.ROOM_TEXT_NUMERIC_MAX_RATIO),
            (f"글자수 중앙값 {fp.text_len_median:.1f} 가 "
             f"{P.ROOM_TEXT_LEN_MIN}~{P.ROOM_TEXT_LEN_MAX}",
             P.ROOM_TEXT_LEN_MIN <= fp.text_len_median <= P.ROOM_TEXT_LEN_MAX),
        ], []),
        (GRID, P.CONF_GRID, [
            (f"초장선 비율 {fp.long_line_ratio:.2f} ≥ {P.GRID_LONG_LINE_MIN_RATIO}",
             fp.long_line_ratio >= P.GRID_LONG_LINE_MIN_RATIO),
            (f"원형 심볼 {fp.type_hist.get('C', 0)}개 ≥ 1", fp.type_hist.get("C", 0) >= 1),
        ], []),
    ]


def _furniture_fallback(fp: LayerFingerprint, hint_name: str) -> list:
    """"위 어디에도 안 걸리는 INSERT 밀집" — 다른 규칙이 전부 진 뒤에만 본다."""
    n_typed = sum(fp.type_hist.values())
    if not n_typed:
        return []
    share = fp.type_hist.get("I", 0) / n_typed
    if share < P.FURNITURE_INSERT_SHARE_MIN:
        return []
    bonus = _name_matches(hint_name, FURNITURE)
    return [(P.CONF_FURNITURE + (P.NAME_HINT_BONUS if bonus else 0.0), FURNITURE,
             [f"다른 조건이 모두 어긋났고 INSERT 비중 {share:.2f} ≥ "
              f"{P.FURNITURE_INSERT_SHARE_MIN}"], bonus)]


def _name_matches(name: str, category: str) -> bool:
    upper = name.upper()
    return any(hint.upper() in upper for hint in NAME_HINTS.get(category, ()))
