# -*- coding: utf-8 -*-
"""[H-0] 도면 슬롯 — 특허 S650 의 계약.

확인하는 것은 셋이다:
  1. 슬롯 셋이 서로를 덮지 않는다 (도면별 상태 독립)
  2. 기존 세션 규약이 그대로다 (평면 dict · 기본 활성 = 평면도)
  3. 열거하지 않은 도면별 키도 슬롯에 갇힌다 (SESSION_KEYS 여집합 규칙)
"""
from __future__ import annotations

import os
import sys

import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from routes.module_f.slots import (  # noqa: E402
    SESSION_KEYS, SLOT_KINDS, _check_slot_kind, _slot_active, _slot_blank,
    _slot_capture, _slot_init, _slot_progress, _slot_state, _slot_switch)


def _sess() -> dict:
    """`_new_session` 이 만드는 것과 같은 모양 — 잡·로그는 세션 전역."""
    s = {"id": "t", "created": 0.0, "touched": 0.0, "job": None, "log": []}
    s.update(_slot_blank())
    _slot_init(s, "plan")
    return s


# ─────────────────────────────────────────────── 1. 기본 규약
def test_기본_활성은_평면도():
    s = _sess()
    assert _slot_active(s) == "plan"
    assert set(s["slots"]) == {"system", "machineroom"}, "활성은 저장소에 없다"


def test_슬롯_종류_검사():
    for k in SLOT_KINDS:
        assert _check_slot_kind(k) == k
    for bad in ("", None, "plan2", "PLAN", "../plan"):
        with pytest.raises(ValueError):
            _check_slot_kind(bad)


def test_슬롯_모르는_옛세션도_평면도로_본다():
    assert _slot_active({"id": "x"}) == "plan"


# ─────────────────────────────────────────────── 2. 독립성
def test_슬롯끼리_덮지_않는다():
    s = _sess()
    s["key"] = "평면.dxf"
    s["pick"] = "PICK-A"

    _slot_switch(s, "system")
    assert s["key"] is None, "계통도 슬롯은 비어 있어야 한다"
    assert s["pick"] is None
    s["key"] = "계통.dxf"
    s["pick"] = "PICK-B"

    _slot_switch(s, "plan")
    assert s["key"] == "평면.dxf" and s["pick"] == "PICK-A"

    _slot_switch(s, "system")
    assert s["key"] == "계통.dxf" and s["pick"] == "PICK-B"


def test_열거하지_않은_키도_슬롯에_갇힌다():
    """`design_sdf_path` 는 _slot_blank 에 없다 — 그래도 새면 안 된다."""
    s = _sess()
    s["design_sdf_path"] = "/out/plan.sdf"
    s["worst_kfp_path"] = "/out/plan.kfp"

    _slot_switch(s, "machineroom")
    assert "design_sdf_path" not in s, "남의 산출물이 따라왔다"
    assert "worst_kfp_path" not in s

    _slot_switch(s, "plan")
    assert s["design_sdf_path"] == "/out/plan.sdf"
    assert s["worst_kfp_path"] == "/out/plan.kfp"


def test_세션_전역은_슬롯을_따라가지_않는다():
    s = _sess()
    s["job"] = {"state": "done"}
    s["log"] = ["한 줄"]
    sid = s["id"]

    _slot_switch(s, "system")
    assert s["id"] == sid
    assert s["job"] == {"state": "done"}, "잡은 세션 전역이다"
    assert s["log"] == ["한 줄"]


def test_같은_슬롯으로_바꾸면_그대로():
    s = _sess()
    s["key"] = "평면.dxf"
    assert _slot_switch(s, "plan") == "plan"
    assert s["key"] == "평면.dxf"


def test_전환이_겹쳐도_상태가_안_섞인다():
    """3차 검토 — 같은 sid 의 전환 «둘» 이 겹치면 순회 중 변경으로 죽거나
    두 슬롯이 반쯤 섞였다. 전환은 _SWITCH_LOCK 으로 직렬화된다."""
    import threading

    s = _sess()
    s["key"] = "평면.dxf"
    errs: list[BaseException] = []

    def flip(n):
        try:
            for _ in range(n):
                _slot_switch(s, "system")
                _slot_switch(s, "plan")
        except BaseException as exc:  # noqa: BLE001
            errs.append(exc)

    ts = [threading.Thread(target=flip, args=(200,)) for _ in range(4)]
    for t in ts:
        t.start()
    for t in ts:
        t.join()
    assert not errs, f"전환 경쟁으로 죽었다: {errs[0]!r}"
    # 어느 쪽이 활성이든 상태는 온전해야 한다 — 평면도의 key 가 살아 있다.
    _slot_switch(s, "plan")
    assert s["key"] == "평면.dxf"
    assert set(s["slots"]) == {"system", "machineroom"}


def test_capture_는_세션전역을_빼고_걷는다():
    s = _sess()
    s["key"] = "k"
    got = _slot_capture(s)
    assert "key" in got
    assert not (SESSION_KEYS & set(got)), "세션 전역이 섞였다"


# ─────────────────────────────────────────────── 3. 진행 보고
def test_진행_단계_판정():
    assert _slot_progress({})["stage"] == ""
    assert _slot_progress({"pick": 1})["stage"] == "pick"
    # edit 이 있으면 pick 보다 앞선다 — /api/module-f/job 의 stage 규약과 같다
    assert _slot_progress({"pick": 1, "edit": 1})["stage"] == "edit"


def test_진행_열림_판정():
    assert _slot_progress({})["opened"] is False
    assert _slot_progress({"dxf": "/x.dxf"})["opened"] is True
    assert _slot_progress({"key": "x.dxf"})["opened"] is True
    assert _slot_progress({"design_sdf_path": "/o.sdf"})["designed"] is True


def test_상태_한장은_세_슬롯을_모두_보고한다():
    s = _sess()
    s["key"] = "평면.dxf"
    s["pick"] = 1
    _slot_switch(s, "system")
    s["key"] = "계통.dxf"

    out = _slot_state(s)
    assert out["active"] == "system"
    assert [x["kind"] for x in out["slots"]] == list(SLOT_KINDS)

    by = {x["kind"]: x for x in out["slots"]}
    assert by["plan"]["opened"] is True and by["plan"]["stage"] == "pick"
    assert by["system"]["active"] is True and by["system"]["key"] == "계통.dxf"
    assert by["machineroom"]["opened"] is False
    assert by["plan"]["label"] == "평면도"


def test_상태_한장은_활성을_평면dict에서_읽는다():
    """저장소가 아니라 지금 펼쳐진 값을 봐야 한다 — 안 그러면 한 박자 늦는다."""
    s = _sess()
    s["key"] = "방금바꾼.dxf"
    by = {x["kind"]: x for x in _slot_state(s)["slots"]}
    assert by["plan"]["key"] == "방금바꾼.dxf"
