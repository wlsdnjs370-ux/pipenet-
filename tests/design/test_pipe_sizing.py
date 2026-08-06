# -*- coding: utf-8 -*-
"""지시서 §9.2 C550·§9.5 C570·§13.3 R5~R6 — 헤드 부착과 관경.

지키는 것은 셋이다. **별표1 값을 그대로 쓴다**는 것(한 단계 올려 잡는 관행은
규정이 아니고 전 구간에 걸리면 자재비가 통째로 늘어난다), **유속 상한은 호칭경이
아니라 배관 역할로 갈린다**는 것(2.5.3.3 단서), 그리고 **닿지 않는 헤드를 0 으로
세지 않는다**는 것(0 으로 세면 그 헤드 몫의 유량이 관경에서 조용히 빠진다).
"""
from __future__ import annotations

import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
sys.path.insert(0, str(_ROOT))
# 평면 import 로 짜인 모듈 A 수리해석기와 유속식을 맞대어 보기 위해서만 쓴다.
sys.path.insert(0, str(_ROOT / "core"))

import hydraulic_solver as HS  # noqa: E402
from core.design.deterministic import nftc_tables as T  # noqa: E402
from core.design.deterministic import pipe_sizing as PS  # noqa: E402
from core.design.deterministic.constraints import build_constraints  # noqa: E402
from core.design.schema import Room  # noqa: E402


@pytest.fixture(scope="module")
def c():
    return build_constraints({
        "building": {"floors_total": 8, "structure": "내화구조", "use": "업무시설"},
        "rooms": [{"use": "사무실", "ambient_temp_max_c": 30.0,
                   "ceiling": {"has_finish": True, "finish_height_mm": 2700,
                               "slab_height_mm": 3200}}],
    }).to_dict()


def _bore(dn: int) -> float:
    return HS.inner_diameter_mm(dn)


def _room(rid: str, has_finish) -> Room:
    ceiling = {"has_finish": has_finish}
    if has_finish:
        ceiling["finish_height_mm"] = 2700
    if has_finish is not None:
        ceiling["slab_height_mm"] = 3200
    return Room.from_dict({"id": rid, "floor": "1F", "use": "사무실",
                           "polygon": [(0, 0), (6000, 0), (6000, 6000), (0, 6000)],
                           "ceiling": ceiling})


def _head(hid: str, room_id: str) -> dict:
    return {"id": hid, "room_id": room_id, "x": 0.0, "y": 0.0}


def _size(dn: int, load: int, role: str, limit: float) -> PS.PipeSize:
    flow = load * 80.0
    return PS.PipeSize(dn=dn, load=load, flow_lpm=flow,
                       velocity_mps=PS.velocity_mps(flow, _bore(dn)),
                       limit_mps=limit, role=role)


# ── C550 신축배관 부착 (§9.2) ────────────────────────────────────────────

def test_반자가_있으면_신축배관으로_매단다(c):
    rooms = [_room("R-A", True)]
    conns, flags = PS.head_connections([_head("H-1", "R-A")], rooms, c)
    assert flags == []
    assert conns[0]["kind"] == "flex"
    assert conns[0]["orientation"] == "pendent"
    assert conns[0]["flex"]["equivalent_length_m"] == 22.4
    assert conns[0]["flex"]["dn"] == 25
    assert conns[0]["flex"]["inner_dia_mm"] == 28.0
    assert conns[0]["flex"]["c_factor"] == 120.0
    assert conns[0]["flex"]["physical_length_m"] == 0.7


def test_반자가_없으면_상향식_직결이고_신축배관이_없다(c):
    conns, flags = PS.head_connections([_head("H-1", "R-B")],
                                       [_room("R-B", False)], c)
    assert flags == []
    assert conns[0]["kind"] == "direct"
    assert conns[0]["orientation"] == "upright"
    assert conns[0]["flex"] is None


def test_반자_유무를_모르면_채우지_않고_보고한다(c):
    """등가길이 22.4 m 는 붙고 안 붙고가 손실을 통째로 가른다 — 짐작으로 정할 수 없다."""
    conns, flags = PS.head_connections(
        [_head("H-1", "R-C"), _head("H-2", "R-C")], [_room("R-C", None)], c)
    assert conns == []
    assert [f["code"] for f in flags] == ["MISSING_BUILDING_FACT"]
    assert flags[0]["room_id"] == "R-C"


def test_실을_못_찾은_헤드는_빠지지_않고_보고된다(c):
    conns, flags = PS.head_connections([_head("H-9", "R-없음")],
                                       [_room("R-A", True)], c)
    assert conns == []
    assert flags[0]["code"] == "HEAD_ROOM_MISSING"
    assert flags[0]["head_id"] == "H-9"


# ── C570 담당 헤드 수 (§9.5) ─────────────────────────────────────────────

_TREE = {"AV": ["T1"], "T1": ["AV", "T2", "H1"],
         "T2": ["T1", "H2", "H3"], "H1": ["T1"], "H2": ["T2"], "H3": ["T2"]}


def test_담당_헤드_수는_하류를_거슬러_누적된다():
    loads, unreached = PS.edge_load(_TREE, "AV", ["H1", "H2", "H3"])
    assert unreached == []
    assert loads == {("AV", "T1"): 3, ("T1", "T2"): 2, ("T1", "H1"): 1,
                     ("T2", "H2"): 1, ("T2", "H3"): 1}


def test_닿지_않는_헤드는_0으로_세지_않는다():
    loads, unreached = PS.edge_load(_TREE, "AV", ["H1", "H2", "H3", "H9"])
    assert unreached == ["H9"]
    assert loads[("AV", "T1")] == 3


# ── C570 관경 (§9.5 / §13.3 R5~R6) ──────────────────────────────────────

def test_R6_담당_4개는_별표_최소_그대로_40A(c):
    """40A 는 별표1 조회 결과다. 여기에 한 단계를 더하면 근거 없는 굵기가 된다."""
    sizes, flags = PS.size_pipes({("T", "B"): 4}, {("T", "B"): "branch"},
                                 c, _bore)
    assert flags == []
    assert sizes[("T", "B")].dn == 40
    assert sizes[("T", "B")].basis == "table"
    assert sizes[("T", "B")].flow_lpm == 320.0


def test_유속이_넘칠_때만_한_단계_올린다(c):
    """담당 30개는 별표1 로 65A 지만 그 관경에서 10.7 m/s 다 — 2.5.3.3 단서에 걸린다."""
    edge = ("MN", "CM")
    sizes, flags = PS.size_pipes({edge: 30}, {edge: "cross"}, c, _bore)
    assert flags == []
    assert T.min_dn(30) == 65                    # 별표1 자체는 65A
    assert sizes[edge].dn == 80                  # 유속 때문에 한 단계
    assert sizes[edge].basis == "velocity"
    assert sizes[edge].velocity_mps <= 10.0


def test_상한은_호칭경이_아니라_배관_역할로_갈린다(c):
    """같은 굵기·같은 유량이라도 가지배관은 6 m/s, 그 밖의 배관은 10 m/s 다."""
    loads = {("A", "B"): 30, ("C", "D"): 30}
    sizes, _f = PS.size_pipes(loads, {("A", "B"): "branch", ("C", "D"): "main"},
                              c, _bore)
    assert sizes[("A", "B")].limit_mps == 6.0
    assert sizes[("C", "D")].limit_mps == 10.0
    assert sizes[("A", "B")].dn > sizes[("C", "D")].dn


def test_교차배관은_별표가_더_작아도_40A_아래로_내려가지_않는다(c):
    edge = ("MN", "CM")
    sizes, _f = PS.size_pipes({edge: 1}, {edge: "cross"}, c, _bore)
    assert T.min_dn(1) == 25
    assert sizes[edge].dn == c["cross_main_min_dn"] == 40
    assert sizes[edge].basis == "cross_main_min"


def test_150A에서도_넘치면_지어내지_않고_보고한다(c):
    edge = ("MN", "RS")
    sizes, flags = PS.size_pipes({edge: 200}, {edge: "main"}, c, _bore)
    assert sizes[edge].dn == 150                 # 별표1 사다리의 끝
    assert [f["code"] for f in flags] == ["VELOCITY_EXCEEDED"]
    assert flags[0]["velocity_mps"] > 10.0


def test_별표_사다리_밖의_호칭경을_만들지_않는다(c):
    rungs = set(T.PIPE_SIZE_TABLE["가"])
    loads = {("N", f"E{n}"): n for n in range(1, 60)}
    roles = {edge: "other" for edge in loads}
    sizes, _f = PS.size_pipes(loads, roles, c, _bore)
    assert {s.dn for s in sizes.values()} <= rungs


# ── R5 단조성 (§13.3) ───────────────────────────────────────────────────

def test_R5_상류보다_굵은_하류는_상류에_맞춰_내린다():
    sizes = {("A", "B"): _size(65, 20, "other", 10.0),
             ("B", "C"): _size(80, 20, "other", 10.0)}
    adjusted = PS.enforce_monotonic(sizes, "A", _bore)
    assert sizes[("B", "C")].dn == 65
    assert sizes[("A", "B")].dn == 65             # 상류는 건드리지 않는다
    assert adjusted[0]["code"] == "MONOTONIC_CLAMP"
    assert (adjusted[0]["from_dn"], adjusted[0]["to_dn"]) == (80, 65)


def test_상류보다_가는_하류는_그대로_둔다():
    sizes = {("A", "B"): _size(80, 30, "other", 10.0),
             ("B", "C"): _size(50, 10, "branch", 6.0)}
    assert PS.enforce_monotonic(sizes, "A", _bore) == []
    assert sizes[("B", "C")].dn == 50


def test_내린_관경이_법정_유속을_다시_넘으면_조용히_넘어가지_않는다():
    """가지배관 상한(6)이 그 밖의 배관(10)보다 좁아 하류가 더 굵어질 수 있다."""
    sizes = {("A", "B"): _size(50, 11, "other", 10.0),
             ("B", "C"): _size(65, 11, "branch", 6.0)}
    adjusted = PS.enforce_monotonic(sizes, "A", _bore)
    assert sizes[("B", "C")].dn == 50             # 명세(R5)대로 내린다
    assert adjusted[0]["code"] == "MONOTONIC_VELOCITY_CONFLICT"
    assert adjusted[0]["velocity_mps"] > 6.0


# ── 유속식 교차검증 ──────────────────────────────────────────────────────

@pytest.mark.parametrize("flow_lpm,dn", [(80.0, 25), (320.0, 40),
                                         (800.0, 50), (2400.0, 80)])
def test_유속식은_모듈_A_수리해석기와_같은_값을_낸다(flow_lpm, dn):
    """식이 두 벌이면 관경은 통과하고 수리계산은 떨어지는 설계가 나온다."""
    bore = _bore(dn)
    assert PS.velocity_mps(flow_lpm, bore) == pytest.approx(
        HS.velocity_mps(flow_lpm, bore))
