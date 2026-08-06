# -*- coding: utf-8 -*-
"""NFTC 표 데이터 — 모듈 C 설계 자동화의 유일한 수치 출처.

여기와 `constraints.py` 밖에서 NFTC 수치를 쓰지 마라. 필요하면 import 한다.
숫자가 여러 곳에 흩어지면 개정 때 어디를 고쳐야 하는지 알 수 없게 되고,
그 상태의 오류는 "표를 잘못 봤다"가 아니라 "한 군데만 고쳤다"로 나타난다.

조문 추적은 4-tuple(`Rule`)로 남긴다. 조번호만 저장하면 개정 때마다 전수
재검토가 필요하지만, 조문 텍스트의 해시를 함께 박아두면 텍스트가 바뀐 순간
테스트가 깨지면서 **영향받는 필드 목록이 실패 메시지로 출력**된다.

[문서정합] `Rule.article_text` 는 지시서 §5.8 이 요구하는 "조문 원문 전체"가
아니라 **본 구현이 실제로 의존하는 규범 내용의 재기술(요지)** 이다. 저장소
안에 NFTC 조문 원문 전문이 없어 현재로서는 원문을 채울 근거가 없다. 해시 기반
개정 감지는 그대로 동작하지만(텍스트를 고치면 깨진다), **원문이 개정되어도
요지가 그대로면 감지되지 않는다**. `ARTICLE_TEXT_SOURCE` 를 "원문" 으로 바꾸는
작업이 지시서 §14.3 V2 의 실제 내용이다.
"""
from __future__ import annotations

import hashlib
from dataclasses import dataclass

# 조문 텍스트의 성격. "요지" 인 동안에는 개정 감지가 반쪽이다(모듈 docstring 참조).
ARTICLE_TEXT_SOURCE = "요지"

# 확인된 최신 시행본. 개정 시 이 값과 함께 아래 텍스트·해시를 갱신한다.
NFTC_EFFECTIVE_DATE = "2026-03-01"


@dataclass(frozen=True)
class Rule:
    """조문 4-tuple. `code` 는 하류 trace 가 참조하는 안정 식별자다."""

    code: str
    article: str
    article_text: str
    effective_date: str
    text_hash: str
    note: str = ""


def _sha256(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()


def _rule(code: str, article: str, article_text: str, note: str = "") -> Rule:
    return Rule(
        code=code,
        article=article,
        article_text=article_text,
        effective_date=NFTC_EFFECTIVE_DATE,
        text_hash=_sha256(article_text),
        note=note,
    )


# ────────────────────────────────────────────────────────────────────────────
# 1. 기준개수 (NFTC 103 표 2.1.1.1)
# ────────────────────────────────────────────────────────────────────────────
#
# ★ 30 이 되는 것은 판매시설과 "판매시설이 설치되는" 복합건축물뿐이다.
#   근린생활시설 단독·운수시설 단독·판매시설 없는 복합건축물은 20 이다.
#   이걸 뒤집으면 기준개수가 1.5배가 되고 수원·펌프가 전부 과대해진다.

RULE_APT = _rule(
    "RULE-NFTC-211-APT", "2.1.1.1",
    "아파트등의 세대 내에 설치된 스프링클러헤드의 기준개수는 10개로 한다.",
    note="공동주택 세대 내. 층수 판정보다 앞선다.",
)
RULE_PARK = _rule(
    "RULE-NFTC-211-PARK", "2.1.1.1",
    "각 동이 주차장으로 연결된 구조인 경우 그 주차장 부분의 기준개수는 30개로 한다.",
)
RULE_HIGH = _rule(
    "RULE-NFTC-211-H", "2.1.1.1",
    "층수가 11층 이상인 특정소방대상물, 지하가 또는 지하역사의 기준개수는 30개로 한다.",
)
RULE_FACTORY_SPECIAL = _rule(
    "RULE-NFTC-211-A", "2.1.1.1",
    "10층 이하 특정소방대상물 중 공장 또는 창고로서 특수가연물을 저장·취급하는"
    " 것의 기준개수는 30개로 한다.",
)
RULE_FACTORY = _rule(
    "RULE-NFTC-211-B", "2.1.1.1",
    "10층 이하 특정소방대상물 중 공장 또는 창고로서 그 밖의 것의 기준개수는"
    " 20개로 한다.",
)
RULE_RETAIL = _rule(
    "RULE-NFTC-211-C", "2.1.1.1",
    "10층 이하 특정소방대상물 중 판매시설 또는 판매시설이 설치되는 복합건축물의"
    " 기준개수는 30개로 한다.",
)
RULE_NEIGHBORHOOD = _rule(
    "RULE-NFTC-211-D", "2.1.1.1",
    "10층 이하 특정소방대상물 중 근린생활시설·운수시설 및 판매시설이 설치되지"
    " 않는 복합건축물의 기준개수는 20개로 한다.",
    note="★ 30 아님. 지시서 G6 가 금지하는 대표 오류.",
)
RULE_OTHER_HIGH_MOUNT = _rule(
    "RULE-NFTC-211-E", "2.1.1.1",
    "10층 이하 특정소방대상물 중 그 밖의 것으로서 헤드의 부착높이가 8m 이상인"
    " 것의 기준개수는 20개로 한다.",
)
RULE_OTHER_LOW_MOUNT = _rule(
    "RULE-NFTC-211-F", "2.1.1.1",
    "10층 이하 특정소방대상물 중 그 밖의 것으로서 헤드의 부착높이가 8m 미만인"
    " 것의 기준개수는 10개로 한다.",
)

# 30 이 되는 판매 관련 용도. `복합건축물` 은 판매시설 포함 여부로 갈리므로 제외.
RETAIL_USES = ("판매시설",)
FACTORY_USES = ("공장", "창고", "랙크식창고")
TWENTY_OR_THIRTY_USES = ("근린생활시설", "판매시설", "운수시설", "복합건축물")

# GATE 용도 선택지의 어휘. 위 기준개수 행들이 **이름으로** 분기하므로, 여기 없는
# 이름을 넣으면 조용히 '그 밖의 것' 으로 떨어진다. 뒤쪽 항목들은 표에 전용 행이
# 없어 실제로 '그 밖의 것' 으로 가지만, 사람이 고를 수 있어야 실명 귀속이 된다.
OCCUPANCY_USES = (
    "근린생활시설", "판매시설", "운수시설", "복합건축물",
    "공장", "창고", "랙크식창고",
    "공동주택", "업무시설", "숙박시설", "의료시설", "노유자시설",
    "교육연구시설", "문화및집회시설", "운동시설", "주차장", "지하가",
)


# ────────────────────────────────────────────────────────────────────────────
# 2. 수원·방사시간·비상전원 (NFTC 103 2.1.1 / 2.9.3)
# ────────────────────────────────────────────────────────────────────────────
#
# 층수 분기를 빠뜨리면 초고층에서 수원이 최대 1/3 로 과소 산정된다.
# 비상전원 시간도 같은 분기를 쓴다.

RULE_WATER_STD = _rule(
    "RULE-NFTC-212-STD", "2.1.1",
    "기준개수에 1.6㎥를 곱한 양 이상으로 하고, 방사시간은 20분으로 한다.",
)
RULE_WATER_30F = _rule(
    "RULE-NFTC-212-H30", "2.1.1",
    "층수가 30층 이상 49층 이하인 특정소방대상물은 기준개수에 3.2㎥를 곱한 양"
    " 이상으로 하고, 방사시간은 40분으로 한다.",
)
RULE_WATER_50F = _rule(
    "RULE-NFTC-212-H50", "2.1.1",
    "층수가 50층 이상인 특정소방대상물은 기준개수에 4.8㎥를 곱한 양 이상으로"
    " 하고, 방사시간은 60분으로 한다.",
)

# (층수 하한, 기준개수 1개당 수원 ㎥, 방사·비상전원 분, Rule)
WATER_SUPPLY_TIERS: tuple[tuple[int, float, int, Rule], ...] = (
    (50, 4.8, 60, RULE_WATER_50F),
    (30, 3.2, 40, RULE_WATER_30F),
    (0, 1.6, 20, RULE_WATER_STD),
)


# ────────────────────────────────────────────────────────────────────────────
# 3. 수평거리 R (NFTC 103 2.7.3)
# ────────────────────────────────────────────────────────────────────────────

RULE_R_STAGE = _rule(
    "RULE-NFTC-273-A", "2.7.3.1",
    "무대부 또는 특수가연물을 저장·취급하는 장소는 수평거리 1.7m 이하로 한다.",
)
RULE_R_RACK_SPECIAL = _rule(
    "RULE-NFTC-273-B", "2.7.3.2",
    "랙크식 창고 중 특수가연물을 저장·취급하는 것은 수평거리 1.7m 이하로 한다.",
)
RULE_R_RACK = _rule(
    "RULE-NFTC-273-C", "2.7.3.2",
    "랙크식 창고로서 그 밖의 것은 수평거리 2.5m 이하로 한다.",
)
RULE_R_APT = _rule(
    "RULE-NFTC-273-D", "2.7.3.3",
    "공동주택(아파트) 세대 내의 거실은 수평거리 3.2m 이하로 한다.",
)
RULE_R_NON_FIRE_RESISTANT = _rule(
    "RULE-NFTC-273-E", "2.7.3.4",
    "그 밖의 특정소방대상물로서 비내화구조는 수평거리 2.1m 이하로 한다.",
)
RULE_R_FIRE_RESISTANT = _rule(
    "RULE-NFTC-273-F", "2.7.3.4",
    "그 밖의 특정소방대상물로서 내화구조는 수평거리 2.3m 이하로 한다.",
)

HORIZONTAL_DISTANCE_M: dict[str, float] = {
    RULE_R_STAGE.code: 1.7,
    RULE_R_RACK_SPECIAL.code: 1.7,
    RULE_R_RACK.code: 2.5,
    RULE_R_APT.code: 3.2,
    RULE_R_NON_FIRE_RESISTANT.code: 2.1,
    RULE_R_FIRE_RESISTANT.code: 2.3,
}


# ────────────────────────────────────────────────────────────────────────────
# 4. 표시온도 (NFTC 103 표 2.7.6)
# ────────────────────────────────────────────────────────────────────────────

RULE_TEMP = _rule(
    "RULE-NFTC-276", "2.7.6",
    "최고 주위온도 39℃ 미만은 표시온도 79℃ 미만, 39℃ 이상 64℃ 미만은 79℃ 이상"
    " 121℃ 미만, 64℃ 이상 106℃ 미만은 121℃ 이상 162℃ 미만, 106℃ 이상은 162℃"
    " 이상의 것으로 한다.",
)

# (주위온도 하한 포함, 주위온도 상한 미만, 표시온도 하한 포함, 표시온도 상한 미만)
TEMPERATURE_BANDS: tuple[tuple[float, float, float, float], ...] = (
    (float("-inf"), 39.0, float("-inf"), 79.0),
    (39.0, 64.0, 79.0, 121.0),
    (64.0, 106.0, 121.0, 162.0),
    (106.0, float("inf"), 162.0, float("inf")),
)

# 각 구간에서 실제로 조달 가능한 표준 표시온도(℃).
# 이 값은 NFTC 가 아니라 **시중 제품 규격**이다. 구간은 법정, 대표값은 시장.
# 다른 제품군을 쓰면 여기만 바꾼다.
STANDARD_MARKING_C: tuple[int, ...] = (72, 93, 141, 182)


# ────────────────────────────────────────────────────────────────────────────
# 5. 별표1 급수관 구경 (NFTC 103 표 2.5.3.3)
# ────────────────────────────────────────────────────────────────────────────
#
# 값은 "해당 호칭경이 담당할 수 있는 헤드 수의 상한"이다.
#   가 — 폐쇄형 헤드 (아래 '나'·'다' 이외)
#   나 — 폐쇄형 헤드 중 무대부·특수가연물 저장취급 장소
#   다 — 개방형 헤드
#
# 관경을 별표 값보다 한 단계 자동으로 올리지 마라. 현장 관행이지 규정이 아니고,
# 전 구간에 걸리면 자재비가 크게 늘어난다.

_UNBOUNDED = 10 ** 9

PIPE_SIZE_TABLE: dict[str, dict[int, int]] = {
    "가": {25: 2, 32: 3, 40: 5, 50: 10, 65: 30,
           80: 60, 90: 80, 100: 100, 125: 160, 150: _UNBOUNDED},
    "나": {25: 2, 32: 3, 40: 5, 50: 10, 65: 30,
           80: 60, 90: 65, 100: 100, 125: 160, 150: _UNBOUNDED},
    "다": {25: 1, 32: 2, 40: 5, 50: 8, 65: 15,
           80: 27, 90: 40, 100: 55, 125: 90, 150: _UNBOUNDED},
}

RULE_PIPE_SIZE = _rule(
    "RULE-NFTC-2533", "2.5.3.3",
    "배관의 구경은 수리계산에 의하거나 표 2.5.3.3의 기준에 따라 설치할 것."
    " 다만 수리계산에 따르는 경우 가지배관의 유속은 6m/s, 그 밖의 배관의 유속은"
    " 10m/s를 초과할 수 없다.",
)

# 표 값의 검증 상태. 근거 없이 "확인함"이라고 쓰지 않는다.
TABLE_VERIFICATION: dict[str, str] = {
    "PIPE_SIZE_TABLE.가":
        "대조 완료 — 모듈 A `remote30_prototype.py` `_nfpc_min_bore_mm` 의 10행과"
        " 1:1 일치. 저장소 안에 독립 구현이 이미 있어 교차 확인이 가능했다.",
    "PIPE_SIZE_TABLE.나":
        "미검증 — 저장소에 대조 대상이 없다. '가'와의 유일한 차이인 90mm 행"
        "(80 → 65)을 포함해 NFTC 원본 별표1 대조가 필요하다(지시서 §14.3 V1).",
    "PIPE_SIZE_TABLE.다":
        "미검증 — 저장소에 대조 대상이 없다. 개방형은 현재 설계 경로에서"
        " 사용되지 않으므로 사용 시점에 반드시 먼저 대조한다.",
}


def min_dn(head_count: int, column: str = "가") -> int:
    """담당 헤드 수 → 법정 최소 호칭경(mm). 별표1 조회 그 자체."""
    table = PIPE_SIZE_TABLE[column]
    for dn in sorted(table):
        if head_count <= table[dn]:
            return dn
    return 150


def next_dn(dn: int, column: str = "가") -> int | None:
    """별표1 사다리에서 한 치수 위. 최대치면 `None`.

    사다리를 따로 두지 않고 표의 행을 그대로 쓴다. 따로 두면 표에는 있는 90A 가
    사다리에는 없는 식으로 어긋나고, 어긋난 쪽으로 상향하면 별표1 조회가 안 되는
    호칭경이 나온다.
    """
    for candidate in sorted(PIPE_SIZE_TABLE[column]):
        if candidate > dn:
            return candidate
    return None


# ────────────────────────────────────────────────────────────────────────────
# 6. 보 이격 (NFTC 103 2.7.7.7)
# ────────────────────────────────────────────────────────────────────────────
#
# 이 표는 헤드 주위 60cm 살수공간 규정(2.7.7.1)과 **별개**다. 보를 60cm 검사에
# 집어넣으면 정상 설계가 대량으로 오탐된다.

RULE_BEAM = _rule(
    "RULE-NFTC-2777", "2.7.7.7",
    "스프링클러헤드의 반사판과 보의 수평거리에 따라, 0.75m 미만은 반사판을 보의"
    " 하단보다 낮게 하고, 0.75m 이상 1m 미만은 수직거리 0.1m 미만, 1m 이상"
    " 1.5m 미만은 0.15m 미만, 1.5m 이상은 0.3m 미만으로 한다.",
)

# (수평거리 상한(미만), 수직거리 조건). "below_beam_bottom" 은 수치가 아니라
# 반사판이 보 하단보다 낮아야 한다는 위치 조건이다.
BEAM_CLEARANCE: tuple[tuple[float, object], ...] = (
    (0.75, "below_beam_bottom"),
    (1.00, 0.10),
    (1.50, 0.15),
    (float("inf"), 0.30),
)


def beam_clearance(horizontal_m: float) -> object:
    """수평거리(m) → 요구 수직거리(m) 또는 `"below_beam_bottom"`."""
    for limit, requirement in BEAM_CLEARANCE:
        if horizontal_m < limit:
            return requirement
    return BEAM_CLEARANCE[-1][1]


# ────────────────────────────────────────────────────────────────────────────
# 7. 헤드 사양·이격·구역
# ────────────────────────────────────────────────────────────────────────────

RULE_PRESSURE_FLOW = _rule(
    "RULE-NFTC-2210", "2.2.1.10",
    "가압송수장치의 정격토출압력은 하나의 헤드선단에 0.1MPa 이상 1.2MPa 이하의"
    " 방수압력이 될 수 있게 하고, 방수량은 80L/min 이상이어야 한다.",
)
RULE_HEAD_CLEARANCE = _rule(
    "RULE-NFTC-2771", "2.7.7.1",
    "살수가 방해되지 않도록 스프링클러헤드로부터 반경 60cm 이상의 공간을 보유할"
    " 것. 다만 벽과 스프링클러헤드 간의 공간은 10cm 이상으로 한다.",
)
RULE_HEAD_TO_CEILING = _rule(
    "RULE-NFTC-2772", "2.7.7.2",
    "스프링클러헤드와 그 부착면과의 거리는 30cm 이하로 할 것.",
)
RULE_ZONE_AREA = _rule(
    "RULE-NFTC-231", "2.3.1",
    "하나의 방호구역의 바닥면적은 3,000㎡를 초과하지 않을 것. 하나의 방호구역은"
    " 2개 층에 미치지 않도록 하되, 1개 층에 설치되는 헤드 수가 10개 이하인 경우와"
    " 복층형 구조의 공동주택은 3개 층 이내로 할 수 있다.",
)
RULE_SPRAY_ZONE = _rule(
    "RULE-NFTC-2412", "2.4.1.2",
    "하나의 방수구역에 설치하는 개방형 스프링클러헤드는 50개 이하로 할 것."
    " 다만 2개 이상의 방수구역으로 나눌 경우에는 하나의 방수구역에 설치하는"
    " 헤드의 수는 25개 이상으로 한다.",
)
RULE_BRANCH_HEADS = _rule(
    "RULE-NFTC-2533-BR", "2.5.3.3",
    "가지배관에 설치되는 스프링클러헤드의 개수는 8개 이하로 할 것."
    " 이는 교차배관에서 분기되는 지점을 기준으로 한쪽 방향의 개수이다.",
    note="★ 전체가 아니라 분기점 기준 한쪽. 양쪽 합계 16개는 적법하다.",
)
RULE_TOURNAMENT = _rule(
    "RULE-NFTC-2532", "2.5.3.2",
    "가지배관의 배열은 토너먼트 방식이 아닐 것.",
)
RULE_CROSS_MAIN = _rule(
    "RULE-NFTC-2534", "2.5.3.4",
    "교차배관의 위치·청소구 및 가지배관의 헤드설치에서 교차배관의 최소구경은"
    " 40mm 이상이 되도록 할 것.",
)

PRESSURE_MPA_MIN = 0.1
PRESSURE_MPA_MAX = 1.2
FLOW_LPM_MIN = 80.0
# 표준형 헤드 K(SI, L/min·bar^-0.5). 제품 사양이며 NFTC 수치가 아니다.
K_FACTOR_STANDARD = 80.0

HEAD_CLEARANCE_RADIUS_M = 0.6
HEAD_TO_WALL_CLEARANCE_M = 0.1
HEAD_TO_CEILING_MAX_M = 0.3

ZONE_AREA_MAX_M2 = 3000.0
ZONE_FLOORS_MAX = 1
SPRAY_ZONE_HEADS_MAX = 50
SPRAY_ZONE_HEADS_MIN_WHEN_SPLIT = 25
BRANCH_HEADS_PER_SIDE_MAX = 8
CROSS_MAIN_MIN_DN = 40

# 2.5.3.3 단서의 법정 상한이다(`RULE_PIPE_SIZE` 본문에 그대로 적혀 있다). 나누는
# 기준은 호칭경이 아니라 배관의 역할 — 가지배관이냐, 그 밖의 배관이냐다.
VELOCITY_LIMIT_MPS: dict[str, float] = {"branch": 6.0, "other": 10.0}


# ────────────────────────────────────────────────────────────────────────────
# 8. 조기반응형 의무 장소 (NFTC 103 2.7.5.5) / 헤드 설치제외 (2.12)
# ────────────────────────────────────────────────────────────────────────────

RULE_QUICK_RESPONSE = _rule(
    "RULE-NFTC-2755", "2.7.5.5",
    "공동주택·노유자시설의 거실, 오피스텔·숙박시설의 침실, 병원의 입원실에는"
    " 조기반응형 스프링클러헤드를 설치할 것.",
)

QUICK_RESPONSE_ROOM_USES: tuple[str, ...] = (
    "공동주택거실", "노유자시설거실", "오피스텔침실", "숙박시설침실", "병원입원실",
)

RULE_HEAD_EXEMPT = _rule(
    "RULE-NFTC-212-EXEMPT", "2.12.1",
    "계단실·경사로·승강기의 승강로·파이프덕트·목욕실·수영장(관람석 제외)·화장실·"
    "직접 외기에 개방되어 있는 복도 등 규정된 장소에는 스프링클러헤드를 설치하지"
    " 않을 수 있다.",
)

HEAD_EXEMPT_PLACES: tuple[str, ...] = (
    "계단실", "경사로", "승강로", "파이프덕트", "목욕실", "수영장", "화장실",
    "개방복도", "발전실", "변전실", "전기실", "통신기기실", "냉동실",
)


# ────────────────────────────────────────────────────────────────────────────
# 9. 개정 감지용 해시 대장
# ────────────────────────────────────────────────────────────────────────────
#
# `Rule.text_hash` 는 생성 시점의 텍스트로 계산되므로 그 자체로는 개정을
# 못 잡는다. 아래 **문자열 리터럴**과 대조해야 비로소 "텍스트가 바뀌었다"가
# 시끄럽게 드러난다. 조문을 고쳤으면 이 대장도 함께 고치면서 영향 필드를
# 재검토하라 — 그 재검토를 강제하는 것이 이 대장의 목적이다.

ALL_RULES: tuple[Rule, ...] = (
    RULE_APT, RULE_PARK, RULE_HIGH,
    RULE_FACTORY_SPECIAL, RULE_FACTORY, RULE_RETAIL, RULE_NEIGHBORHOOD,
    RULE_OTHER_HIGH_MOUNT, RULE_OTHER_LOW_MOUNT,
    RULE_WATER_STD, RULE_WATER_30F, RULE_WATER_50F,
    RULE_R_STAGE, RULE_R_RACK_SPECIAL, RULE_R_RACK, RULE_R_APT,
    RULE_R_NON_FIRE_RESISTANT, RULE_R_FIRE_RESISTANT,
    RULE_TEMP, RULE_PIPE_SIZE, RULE_BEAM,
    RULE_PRESSURE_FLOW, RULE_HEAD_CLEARANCE, RULE_HEAD_TO_CEILING,
    RULE_ZONE_AREA, RULE_SPRAY_ZONE, RULE_BRANCH_HEADS,
    RULE_TOURNAMENT, RULE_CROSS_MAIN,
    RULE_QUICK_RESPONSE, RULE_HEAD_EXEMPT,
)

RULES: dict[str, Rule] = {r.code: r for r in ALL_RULES}

RULE_TEXT_HASHES: dict[str, str] = {
    "RULE-NFTC-211-APT":         "1c3d6b1bb282967d22480f390da84d4c33435701df79b12e3e16e7578614b7e7",
    "RULE-NFTC-211-PARK":        "6d434721d983cc0638edfb05fbe3e021b9cd681568ea609b1e0efc82d8d41b32",
    "RULE-NFTC-211-H":           "e8c2224d76b2012a68a262298eb6b5c9d9840b638ed35117c1c5233bb29ca7fb",
    "RULE-NFTC-211-A":           "f8a537f46cfd638b2002bd301527f0b84ad52b2f085f584588b7b37fd3f51d8d",
    "RULE-NFTC-211-B":           "9e435ea3cb970b1f06772f3a7c35e939f29421055119220b0ddd88e28ee219f8",
    "RULE-NFTC-211-C":           "427bd01d22c065e26bd008260f8802caedfc74050f407b172c67992b426627d8",
    "RULE-NFTC-211-D":           "16cf4a3c6fce8dcea0475f224703b2b894b70354395831c6e38a54542ba5a5c8",
    "RULE-NFTC-211-E":           "ae2ca9bb55120ac3af4f6b1247e28c9f625357b69d4d37c4a80b8dc1a363afd6",
    "RULE-NFTC-211-F":           "b89597cc2d515d59fc05a7d2a7147cf7c85772e7688313ab7b54b81c05394855",
    "RULE-NFTC-212-STD":         "9d0120f9f3cb79f3d328f4fa8ab87c90515ca721fa16b6c7771daa313918e9e7",
    "RULE-NFTC-212-H30":         "345b49974a06ba3088152304a73639e9adab411b5ff929e0dc513554832605f5",
    "RULE-NFTC-212-H50":         "7fcc560e368a55078a641297b989569fa4550f26053c4aea93f75760a8d1c071",
    "RULE-NFTC-273-A":           "d8a8ba2a729dcc8e97067f0e8c9328e90b2a51d17cab646c60c19f35cfb340b7",
    "RULE-NFTC-273-B":           "70a3f7ae6ebe5d4cee0ea2aa4f9b161cc5429a4162f74776f21959756d4d356d",
    "RULE-NFTC-273-C":           "3c4b89a124bc30cc18712f7d5be1a4e7997662e0d2fe1c859209bd01400be764",
    "RULE-NFTC-273-D":           "e4eb4de5ffff41d051aa3668fa51fa24b915f45e6042234d9679cb452e125b77",
    "RULE-NFTC-273-E":           "4b7cb7002e4ef689d825a3766d90bc505052ab12760af92076721279075de898",
    "RULE-NFTC-273-F":           "027f1a79904a61251c304b9d573f6b040f30d7be832c032c9cd5b085f5401d26",
    "RULE-NFTC-276":             "af282a55be6f119754449e20f6242ae9fbcd299186377956e18f00133fdac41f",
    "RULE-NFTC-2533":            "3e4a5c84684af97a9de84f11d4f5bc83c8f225c42a6e611312c9c1d1032d546a",
    "RULE-NFTC-2777":            "f2718611c4b227f81098ff43a0015c791bc56df17c116e5402b5381335cda1a8",
    "RULE-NFTC-2210":            "56140358e554795e8b360053bddcd721ab1aef2504ec7daa2e782a603d6add13",
    "RULE-NFTC-2771":            "658dc256712ba2ecabd38060bf32be996052dd4a6825dbef44453a04de8d3034",
    "RULE-NFTC-2772":            "51d949676265ef726a560752391de9b95059b0fcbe1d4bc26393868a2d6c544d",
    "RULE-NFTC-231":             "077d4046ff4f1fffb678dc0b63f98f292a1c60bcf669db5cf03b673e3505382d",
    "RULE-NFTC-2412":            "543483d1a39cc52977a40c296d7b01adaff6d06da406288444f381143faf7e44",
    "RULE-NFTC-2533-BR":         "f06ad027762c8fb8963fe62c64d1ba5c1692afae2594e17dc88595ed68482704",
    "RULE-NFTC-2532":            "a603ec49c9c9617f07512914b68cee63a96b62d2223a2e838208f041759bd907",
    "RULE-NFTC-2534":            "bbc43c7096ba62ad3b46a0b4ef59d9080553b2d2c1d473ffdeaa241f2f4eee6d",
    "RULE-NFTC-2755":            "41e8c2dee2f4e4c3b44f10405ad67dbb85f4f9411f70661ef8414e591a0d8151",
    "RULE-NFTC-212-EXEMPT":      "77779690b897dda6e683849b5a9f84421201d1d62fd187289ffd47ddb3c3f5a7",
}
