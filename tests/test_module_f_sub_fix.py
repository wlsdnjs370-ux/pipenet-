# -*- coding: utf-8 -*-
"""[§27 후속] 뽑힌 계통도·기계실 배관을 사람이 고칠 수 있다.

■ 왜 «고치는 자리» 인가

  대명동 계통도를 뽑으면 배관 53개가 나오는데 **관경이 53개 모두 «추측
  150A»** 다(도면 치수 텍스트 매치 0건 · `scripts/_probe_bridge_share.py`
  옆에서 같이 쟀다). 그 값이 그대로 최종 SDF 의 입상관이 되는데 볼 자리도
  고칠 자리도 없었다.

  그래프도 조각나 있어 200mm~10m «다리» 를 61~66개 놓아 잇는다. 다리를
  자동으로 판정해 표시하는 길도 있었지만, 사용자 결정은 «다리는 고려하지
  말고 사람이 쉽게 고치게 하라» 였다. §27 에서 두 번 확인한 방향과도 같다 —
  추측 규칙을 얹으면 «틀린 확신» 만 는다.

■ 이 시험이 지키는 넷

  ⑴ 자리는 **좌표**로 가리킨다 — 라벨(`r1`)은 경로 순서라 다시 뽑으면
     같은 이름이 다른 배관을 가리킨다(D-F11-4 가 겪은 사고).
  ⑵ **원래 값이 살아 있다** — 덮기를 지우면 정확히 되돌아온다.
  ⑶ **다시 뽑아도 손질이 안 사라진다** — 좌표가 같으면 되붙는다.
  ⑷ **못 붙인 것을 조용히 버리지 않는다**.
"""
from __future__ import annotations

import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from routes.module_f.sub_fix import (        # noqa: E402
    apply_overrides, parse_rows, pipe_keys, rows_for_view)


def _got():
    """절점 3개가 한 줄로 — 뽑힌 결과의 최소 모양."""
    return {
        "nodes": [{"label": "n1", "x": 0, "y": 0},
                  {"label": "n2", "x": 1000, "y": 0},
                  {"label": "10", "x": 1000, "y": 2000}],
        "pipes": [
            {"label": "r1", "in": "n1", "out": "n2", "dia": 150,
             "length": 1.0, "dia_source": "default"},
            {"label": "r2", "in": "n2", "out": "10", "dia": 150,
             "length": 2.0, "dia_source": "default"},
        ],
    }


def test_자리는_라벨이_아니라_좌표로_가리킨다():
    """★라벨은 경로 순서로 매겨진다 — 다시 뽑으면 같은 이름이 남을 가리킨다."""
    keys = pipe_keys(_got())
    assert keys["r1"] == ((0, 0), (1000, 0))
    # 방향이 뒤집혀도 같은 자리다 — 뽑는 방향은 사람이 어디를 먼저 찍느냐다.
    flipped = _got()
    flipped["pipes"][0]["in"], flipped["pipes"][0]["out"] = "n2", "n1"
    assert pipe_keys(flipped)["r1"] == keys["r1"]


def test_덮으면_원래_값이_같이_남는다():
    got = _got()
    rows = parse_rows([{"a": [0, 0], "b": [1000, 0], "dia": 100,
                        "note": "도면 치수 확인"}])
    stat = apply_overrides(got, rows)
    assert stat == {"applied": 1, "given": 1, "unmatched": 0}
    p = got["pipes"][0]
    assert p["dia"] == 100 and p["orig_dia"] == 150
    assert p["dia_source"] == "user_fix" and p["orig_dia_source"] == "default"
    assert p["fix_note"] == "도면 치수 확인"
    assert got["pipes"][1]["dia"] == 150, "안 고른 배관까지 건드렸다"


def test_덮기를_지우면_정확히_되돌아온다():
    """★«덮은 위에 덮기» 로 원래 값이 씻겨 나가면 되돌릴 길이 없어진다."""
    got = _got()
    before = [dict(p) for p in got["pipes"]]
    apply_overrides(got, parse_rows([{"a": [0, 0], "b": [1000, 0], "dia": 100}]))
    apply_overrides(got, parse_rows([{"a": [0, 0], "b": [1000, 0], "dia": 80}]))
    assert got["pipes"][0]["orig_dia"] == 150, "원래 값이 80 으로 씻겼다"
    apply_overrides(got, [])
    assert got["pipes"] == before


def test_길이를_고치면_연장도_같이_바뀐다():
    """요약과 표가 다른 말을 하면 사람은 요약을 믿는다."""
    got = _got()
    apply_overrides(got, parse_rows([{"a": [1000, 0], "b": [1000, 2000],
                                      "length": 5.5}]))
    assert got["pipes"][1]["length"] == 5.5
    assert got["pipes"][1]["orig_length"] == 2.0
    assert got["total_length_m"] == 6.5


def test_다시_뽑아도_손질이_되붙는다():
    """★안 되붙이면 추출을 한 번 더 누른 순간 손질이 통째로 사라진다."""
    rows = parse_rows([{"a": [0, 0], "b": [1000, 0], "dia": 65}])
    fresh = _got()                       # 다시 뽑은 «새» 결과
    stat = apply_overrides(fresh, rows)
    assert stat["applied"] == 1
    assert fresh["pipes"][0]["dia"] == 65


def test_사라진_구간은_조용히_안_버린다():
    got = _got()
    rows = parse_rows([{"a": [9, 9], "b": [8, 8], "dia": 65}])
    stat = apply_overrides(got, rows)
    assert stat == {"applied": 0, "given": 1, "unmatched": 1}


@pytest.mark.parametrize("row, word", [
    ({"a": [0, 0], "b": [1, 1]}, "관경이나 길이"),
    ({"a": [0, 0], "b": [1, 1], "dia": "굵게"}, "숫자가 아닙니다"),
    ({"a": [0, 0], "b": [1, 1], "dia": 0}, "범위"),
    ({"a": [0, 0], "b": [1, 1], "length": 0}, "범위"),
    ({"b": [1, 1], "dia": 65}, "두 끝 좌표"),
    ({"a": [0, 0], "b": [1, 1], "dia": 65, "note": "가" * 201}, "너무 깁니다"),
])
def test_잘못된_입력은_사유를_들고_막힌다(row, word):
    """★길이 0 을 열어 두면 PIPENET 이 그 배관을 못 푼다 — 문제를 뒤로 미룰 뿐."""
    with pytest.raises(ValueError) as e:
        parse_rows([row])
    assert word in str(e.value), str(e.value)


def test_같은_배관을_두_번_덮으면_막힌다():
    with pytest.raises(ValueError):
        parse_rows([{"a": [0, 0], "b": [1, 1], "dia": 65},
                    {"a": [1, 1], "b": [0, 0], "dia": 80}])


def test_표에는_고친_자리가_보인다():
    got = _got()
    apply_overrides(got, parse_rows([{"a": [0, 0], "b": [1000, 0], "dia": 100,
                                      "note": "협의"}]))
    rows = rows_for_view(got)
    assert rows[0]["orig_dia"] == 150 and rows[0]["fixed"] is True
    assert rows[0]["fix_note"] == "협의"
    assert rows[0]["a"] == [0, 0] and rows[0]["b"] == [1000, 0]
    assert "orig_dia" not in rows[1], "안 고친 행에 덮기 표시가 붙었다"
