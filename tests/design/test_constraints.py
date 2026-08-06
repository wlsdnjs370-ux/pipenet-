# -*- coding: utf-8 -*-
"""지시서 §13.2 — 결정론 코어 단위 테스트.

이 파일이 통과한다는 것은 "도면과 무관하게 NFTC 표를 옳게 읽는다"는 뜻이고,
도면이 바뀌어도 깨지지 않아야 한다. 여기가 깨지면 인식 셸이 아무리 정확해도
전부 틀린 기준 위에서 통과한 값이다.
"""
from __future__ import annotations

import math
import sys
from dataclasses import FrozenInstanceError
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
# 기존 엔진(`nftc_rules`)은 배포판마다 `core/` 아래에도, 루트 평면에도 놓인다.
# 둘 다 경로에 올리고 평면 이름으로 부르면 어느 배치에서도 같은 모듈을 읽는다.
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from core.design.deterministic import constraints as C  # noqa: E402
from core.design.deterministic import nftc_tables as T  # noqa: E402


# ────────────────────────────────────────────────────────────────────────────
# T1~T11 — 기준개수 (표 2.1.1.1)
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("tid,kwargs,expected", [
    ("T1", dict(use="근린생활시설", floors_total=8, head_mount_height_m=4.0), 20),
    ("T2", dict(use="운수시설", floors_total=5, head_mount_height_m=4.0), 20),
    ("T3", dict(use="판매시설", floors_total=6, head_mount_height_m=4.0), 30),
    ("T4", dict(use="복합건축물", floors_total=9, head_mount_height_m=4.0,
                has_retail_occupancy=True), 30),
    ("T5", dict(use="복합건축물", floors_total=9, head_mount_height_m=4.0,
                has_retail_occupancy=False), 20),
    ("T6", dict(use="업무시설", floors_total=7, head_mount_height_m=7.5), 10),
    ("T7", dict(use="업무시설", floors_total=7, head_mount_height_m=8.0), 20),
    ("T8", dict(use="업무시설", floors_total=11, head_mount_height_m=4.0), 30),
    ("T9", dict(use="공동주택", floors_total=15, is_apartment_unit=True), 10),
    ("T10", dict(use="공장", floors_total=5, has_special_combustible=True), 30),
    ("T11", dict(use="공장", floors_total=5, has_special_combustible=False), 20),
])
def test_scenario_head_count(tid, kwargs, expected):
    n, code = C.scenario_head_count(**kwargs)
    assert n == expected, f"{tid}: {kwargs} → {n} (기대 {expected}, 근거 {code})"
    assert code in T.RULES


def test_t9_apartment_horizontal_distance():
    """T9 는 기준개수 10 과 R=3.2 를 함께 요구한다."""
    r_m, code = C.horizontal_distance(is_apartment_unit=True)
    assert r_m == 3.2
    assert code == T.RULE_R_APT.code


def test_apartment_beats_floor_count():
    """행 순서 회귀 — 아파트를 11층 행 뒤에 두면 10 이 30 으로 뒤집힌다."""
    n, _ = C.scenario_head_count(use="공동주택", floors_total=25, is_apartment_unit=True)
    assert n == 10


def test_connected_parking_beats_apartment_only_when_flagged():
    n, code = C.scenario_head_count(
        use="공동주택", floors_total=25, is_connected_parking=True)
    assert (n, code) == (30, T.RULE_PARK.code)


# ────────────────────────────────────────────────────────────────────────────
# W1~W5 — 수원·방사시간 경계
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("wid,n,floors,volume,minutes", [
    ("W1", 20, 29, 32.0, 20),
    ("W2", 20, 30, 64.0, 40),
    ("W3", 20, 49, 64.0, 40),
    ("W4", 20, 50, 96.0, 60),
    ("W5", 30, 55, 144.0, 60),
])
def test_water_supply(wid, n, floors, volume, minutes):
    got_volume, got_minutes = C.water_supply(n, floors)
    assert got_volume == pytest.approx(volume), wid
    assert got_minutes == minutes, wid


@pytest.mark.parametrize("floors", [1, 29, 30, 49, 50, 80])
def test_emergency_power_follows_discharge_branch(floors):
    """비상전원 시간은 방사시간과 같은 층수 분기를 따라야 한다(G1)."""
    assert C.emergency_power_minutes(floors) == C.water_supply(1, floors)[1]


# ────────────────────────────────────────────────────────────────────────────
# 관경 경계 (별표1 '가')
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("heads,dn", [
    (2, 25), (3, 32), (4, 40), (5, 40), (6, 50),
    (10, 50), (11, 65), (30, 65), (31, 80), (161, 150),
])
def test_min_dn_boundaries(heads, dn):
    assert T.min_dn(heads) == dn


def test_min_dn_column_b_differs_at_90():
    """'나' 칸 90mm 행만 '가' 칸과 다르다(65 vs 80). 표를 복사하다 흔히 뭉갠다."""
    assert T.PIPE_SIZE_TABLE["가"][90] == 80
    assert T.PIPE_SIZE_TABLE["나"][90] == 65
    assert T.min_dn(70, "가") == 90
    assert T.min_dn(70, "나") == 100


def test_pipe_size_table_is_monotonic():
    for column, rows in T.PIPE_SIZE_TABLE.items():
        caps = [rows[dn] for dn in sorted(rows)]
        assert caps == sorted(caps), f"'{column}' 칸 담당 헤드 수가 단조증가가 아니다"


def test_min_dn_is_not_upsized():
    """G5 — 별표 최소를 그대로 쓴다. 한 단계 올려 '안전하게' 만들지 마라."""
    assert T.min_dn(4) == 40
    assert T.min_dn(30) == 65


# ────────────────────────────────────────────────────────────────────────────
# 보 이격 경계
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("horizontal,expected", [
    (0.74, "below_beam_bottom"),
    (0.75, 0.10),
    (0.99, 0.10),
    (1.00, 0.15),
    (1.49, 0.15),
    (1.50, 0.30),
])
def test_beam_clearance(horizontal, expected):
    assert T.beam_clearance(horizontal) == expected


# ────────────────────────────────────────────────────────────────────────────
# 표시온도·수평거리
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("ambient,rating", [
    (0.0, 72), (38.0, 72), (39.0, 93), (63.0, 93),
    (64.0, 141), (105.0, 141), (106.0, 182),
])
def test_temp_rating(ambient, rating):
    assert C.temp_rating(ambient)[0] == rating


@pytest.mark.parametrize("kwargs,r_m", [
    (dict(), 2.1),
    (dict(structure="내화구조"), 2.3),
    (dict(is_stage=True), 1.7),
    (dict(has_special_combustible=True), 1.7),
    (dict(is_rack_storage=True), 2.5),
    (dict(is_rack_storage=True, has_special_combustible=True), 1.7),
    (dict(is_apartment_unit=True), 3.2),
])
def test_horizontal_distance(kwargs, r_m):
    assert C.horizontal_distance(**kwargs)[0] == r_m


def test_head_spacing_square():
    assert C.head_spacing_square(2.3) == pytest.approx(2.3 * math.sqrt(2.0))


# ────────────────────────────────────────────────────────────────────────────
# 개정 감지 (지시서 §5.8)
# ────────────────────────────────────────────────────────────────────────────

def test_rule_registry_is_consistent():
    assert len(T.ALL_RULES) == len(T.RULES), "rule code 가 중복됐다"
    assert set(T.RULES) == set(T.RULE_TEXT_HASHES)


def test_rule_text_hashes_unchanged():
    """조문 요지가 바뀌면 여기서 시끄럽게 터진다. 해시만 고쳐 통과시키지 마라 —
    바뀐 조문에 딸린 수치를 전부 재확인한 뒤 갱신하는 것이 이 테스트의 목적이다."""
    for code, rule in T.RULES.items():
        assert rule.text_hash == T.RULE_TEXT_HASHES[code], f"{code} 조문 요지 변경됨"


def test_every_rule_carries_article_and_date():
    for rule in T.ALL_RULES:
        assert rule.article, f"{rule.code} 조번호 누락"
        assert rule.effective_date == T.NFTC_EFFECTIVE_DATE


# ────────────────────────────────────────────────────────────────────────────
# build_constraints
# ────────────────────────────────────────────────────────────────────────────

def _building(**overrides) -> dict:
    b = {
        "floors_total": 8,
        "structure": "내화구조",
        "use": "업무시설",
    }
    b.update(overrides)
    return {
        "building": b,
        "rooms": [
            {"use": "사무실", "ambient_temp_max_c": 30.0,
             "ceiling": {"has_finish": True, "finish_height_mm": 2700,
                         "slab_height_mm": 3200}},
        ],
    }


def test_build_constraints_roundtrip():
    c = C.build_constraints(_building())
    assert c.scenario_head_count == 10
    assert c.horizontal_distance_m == 2.3
    assert c.head_spacing_square_m == pytest.approx(2.3 * math.sqrt(2.0))
    assert c.wall_clearance_max_m == pytest.approx(c.head_spacing_square_m / 2.0)
    assert c.discharge_minutes == c.emergency_power_minutes == 20
    assert c.temp_rating_c == 72
    d = c.to_dict()
    assert d["schema"] == "fncadnet.constraints/1"
    assert len(d["trace"]) == len(c.trace)
    assert all(t["code"] in T.RULES for t in d["trace"])


def test_constraints_is_frozen():
    c = C.build_constraints(_building())
    with pytest.raises(FrozenInstanceError):
        c.horizontal_distance_m = 3.2  # type: ignore[misc]


def test_zone_grid_relief_is_undecided():
    """D1 — 격자형 완화는 결정 대기. 값이 들어오면 라우팅 모드부터 확인하라."""
    assert C.build_constraints(_building()).zone_grid_relief_m2 is None


@pytest.mark.parametrize("missing", ["floors_total", "structure", "use"])
def test_missing_fact_raises_not_defaults(missing):
    payload = _building()
    payload["building"].pop(missing)
    with pytest.raises(C.MissingBuildingFact) as e:
        C.build_constraints(payload)
    assert e.value.field.endswith(missing)


def test_missing_ceiling_raises():
    payload = _building()
    payload["rooms"][0]["ceiling"] = {}
    with pytest.raises(C.MissingBuildingFact) as e:
        C.build_constraints(payload)
    assert e.value.field == "rooms[].ceiling"


def test_missing_ambient_temp_raises():
    payload = _building()
    payload["rooms"][0].pop("ambient_temp_max_c")
    with pytest.raises(C.MissingBuildingFact) as e:
        C.build_constraints(payload)
    assert e.value.field == "rooms[].ambient_temp_max_c"


def test_head_mount_height_uses_slab_when_no_finish():
    """반자가 없으면 슬래브 높이가 부착높이다. 8m 경계를 이 선택이 가른다."""
    payload = _building(use="창고", floors_total=3)
    payload["rooms"][0]["ceiling"] = {"has_finish": False, "slab_height_mm": 9000,
                                      "finish_height_mm": 2700}
    assert C.build_constraints(payload).scenario_head_count == 20


def test_branch_heads_is_per_side():
    """G3 — 8개는 분기점 기준 한쪽. 양쪽 합계 16개는 적법하다."""
    c = C.build_constraints(_building())
    assert c.branch_heads_per_side_max == 8
    assert not hasattr(c, "branch_heads_max")


# ────────────────────────────────────────────────────────────────────────────
# 기존 엔진(core/nftc_rules.py) 교차검증
# ────────────────────────────────────────────────────────────────────────────

_USE_TO_LEGACY = {
    "근린생활시설": "neighborhood",
    "운수시설": "transit",
    "판매시설": "retail",
    "공장": "factory",
    "창고": "warehouse",
    "업무시설": "other_low",
}


@pytest.mark.parametrize("use,floors,mount,special", [
    ("근린생활시설", 8, 4.0, False),
    ("운수시설", 5, 4.0, False),
    ("판매시설", 6, 4.0, False),
    ("공장", 5, 4.0, True),
    ("공장", 5, 4.0, False),
    ("창고", 5, 4.0, False),
    ("업무시설", 11, 4.0, False),
])
def test_agrees_with_legacy_nftc_rules(use, floors, mount, special):
    """모듈 C 결정론 코어와 기존 `core/nftc_rules.py` 가 같은 답을 내야 한다.

    두 개의 진실 출처가 조용히 갈라지는 것이 가장 비싼 실패다. 아래
    `test_known_divergence_*` 가 고정한 1건 외에 새 불일치가 생기면 실패한다.
    """
    import nftc_rules as legacy

    mine, _ = C.scenario_head_count(
        use=use, floors_total=floors, head_mount_height_m=mount,
        has_special_combustible=special)
    theirs = legacy.decide_reference_count({
        "use": _USE_TO_LEGACY[use],
        "floors_total": floors,
        "head_attach_h_m": mount,
        "has_special_combustible": special,
    }).value
    assert mine == theirs, f"{use}/{floors}층 — 모듈C {mine} vs nftc_rules {theirs}"


def test_known_divergence_complex_without_retail():
    """알려진 불일치 1건 — 판매시설이 없는 복합건축물.

    `core/nftc_rules.py` 는 use=="complex" 만 보고 30 을 준다. 표 2.1.1.1 의
    해당 행은 "판매시설 **또는 판매시설이 설치되는** 복합건축물"이므로 판매시설이
    없으면 20 이 맞다(지시서 T5). 기존 엔진 쪽이 과대설계(비용) 방향이라 안전
    문제는 아니지만, 고쳐질 때까지 이 차이를 여기에 고정해 둔다. 기존 엔진이
    수정되면 이 테스트가 실패하고 — 그때 이 테스트를 지우면 된다.
    """
    import nftc_rules as legacy

    mine, _ = C.scenario_head_count(
        use="복합건축물", floors_total=9, head_mount_height_m=4.0,
        has_retail_occupancy=False)
    theirs = legacy.decide_reference_count({
        "use": "complex", "floors_total": 9,
        "head_attach_h_m": 4.0, "has_special_combustible": False,
    }).value
    assert mine == 20
    assert theirs == 30


def test_agrees_with_legacy_temperature_rating():
    import nftc_rules as legacy

    for ambient in (0.0, 38.0, 39.0, 63.0, 64.0, 105.0, 106.0):
        mine = C.temp_rating(ambient)[0]
        band = legacy.decide_temperature_rating(ambient_temp_c=ambient).value
        lo = band["min_c"]
        hi = band["max_c"] if band["max_c"] is not None else float("inf")
        assert lo <= mine <= hi, (
            f"{ambient}℃ — 모듈C 대표값 {mine} 이 기존 구간 {lo}~{hi} 밖")
