# -*- coding: utf-8 -*-
"""기준개수 표 — NFTC 103 표 2.1.1.1 이 화면까지 «한 출처» 로 간다.

법정 수치를 두 곳에 두면 개정이 왔을 때 한쪽만 고쳐지고, 그 어긋남은 기준개수가
30 이어야 할 자리에 20 이 들어간 산출로만 드러난다. 그래서 화면은 표를 옮겨
적지 않고 서버에서 받는다 — 그 계약을 여기서 못박는다.
"""
from __future__ import annotations

import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from nftc_rules import (  # noqa: E402
    _REFERENCE_COUNT_TABLE, decide_reference_count, reference_count_options)


def test_목록이_표와_같은_길이다():
    assert len(reference_count_options()) == len(_REFERENCE_COUNT_TABLE)


def test_목록_순서가_표_순서다():
    """순서가 곧 규칙이다 — 먼저 맞는 행이 이긴다. 화면도 같은 순서여야 한다."""
    got = [r["rule_id"] for r in reference_count_options()]
    want = [r["rule_id"] for r in _REFERENCE_COUNT_TABLE]
    assert got == want


def test_각_행이_규칙번호_이름_개수를_갖는다():
    for row in reference_count_options():
        assert set(row) == {"rule_id", "label", "count"}
        assert isinstance(row["count"], int) and row["count"] > 0
        assert row["label"] and row["rule_id"]


def test_개수가_표의_값_그대로다():
    for got, want in zip(reference_count_options(), _REFERENCE_COUNT_TABLE):
        assert got["count"] == want["count"], got["rule_id"]
        assert got["label"] == want["label"]


def test_표에_10_20_30이_모두_있다():
    """세 값이 다 나오지 않으면 표를 잘못 읽은 것이다."""
    counts = {r["count"] for r in reference_count_options()}
    assert {10, 20, 30} <= counts, counts


def test_아파트는_10개다():
    """순서 회귀 방벽 — 11층 이상 행이 앞에 오면 국내 아파트가 30 으로 뒤집힌다."""
    d = decide_reference_count({"use": "apartment", "floors_total": 25})
    assert d.value == 10


def test_아파트라도_지하주차장_연결이면_30개다():
    d = decide_reference_count({"use": "apartment", "floors_total": 25,
                                "connected_to_basement_parking": True})
    assert d.value == 30


def test_목록이_결정함수와_어긋나지_않는다():
    """같은 rule_id 면 같은 개수여야 한다 — 두 경로가 한 표를 본다는 뜻."""
    by_id = {r["rule_id"]: r["count"] for r in reference_count_options()}
    d = decide_reference_count({"use": "retail", "floors_total": 5})
    assert by_id[d.rule_id] == d.value


def test_목록은_원본을_건드리지_않는다():
    """돌려준 dict 를 고쳐도 표는 그대로여야 한다."""
    rows = reference_count_options()
    rows[0]["count"] = 999
    assert reference_count_options()[0]["count"] != 999
