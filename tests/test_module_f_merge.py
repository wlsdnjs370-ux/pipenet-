# -*- coding: utf-8 -*-
"""[H-1] S700 접합 — G 의 설계 표를 A 의 헤드망 규약으로.

특허 S550 은 «기준점 번호 = 10», S740 은 «기준점 번호 10 을 공통 절점으로
결합» 이라고 못박는다. G 는 BFS 로 1 부터 매기므로 +9 오프셋이 필요하다.

여기서 못박는 계약:
  ① 기준점이 10 이 되고 급수원(Input)이다
  ② 라벨이 박힌 자리를 **빠짐없이** 옮긴다 (배관·노즐·부속·기기의 in/out)
  ③ 배관 이름(label · pipe)은 노드 라벨이 아니므로 **옮기지 않는다**
  ④ 성립하지 않으면 임의로 메우지 않고 MergeError 를 올린다
"""
from __future__ import annotations

import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from routes.module_f.merge import (  # noqa: E402
    ANCHOR_LABEL, LABEL_OFFSET, SUPPLY_MODES, HeadTables, MergeError,
    _shift, check_supply_mode, to_head_tables, zone_type_of)


class _G:
    """G 의 `PipeTablesG` 를 흉내낸 최소 표 — 필드 이름이 계약이다."""

    def __init__(self, **kw):
        self.nodes = kw.get("nodes", [])
        self.pipes = kw.get("pipes", [])
        self.nozzles = kw.get("nozzles", [])
        self.fittings = kw.get("fittings", [])
        self.equipment = kw.get("equipment", [])
        self.meta = kw.get("meta", [])


def _sample() -> _G:
    """급수원 1 · 중간 2 · 헤드 3 — G 가 내놓는 모양 그대로."""
    return _G(
        nodes=[
            {"label": "1", "x": 0, "y": 0, "elevation": 0.0, "io_node": "Input"},
            {"label": "2", "x": 100, "y": 0, "elevation": 0.0, "io_node": "No"},
            {"label": "3", "x": 200, "y": 0, "elevation": 0.0, "io_node": "No"},
        ],
        pipes=[
            {"label": "P1", "in": "1", "out": "2", "dia": 65, "length": 1.0},
            {"label": "P2", "in": "2", "out": "3", "dia": 25, "length": 1.0},
        ],
        nozzles=[{"label": "1", "in": "3", "out": "@/1", "flow_lmin": 80}],
        fittings=[{"pipe": "P2", "in": "2", "out": "3", "type": "Elbow 90",
                   "count": 1}],
        equipment=[{"pipe": "P1", "in": "1", "out": "2", "label": "AV",
                    "desc": "A/V", "eq_len": 0.0, "rel_pos": 0.5}],
        meta=[("제목", "시험")],
    )


# ─────────────────────────────────────────────── ① 기준점
def test_기준점이_10이_된다():
    ht = to_head_tables(_sample())
    assert isinstance(ht, HeadTables)
    assert [n["label"] for n in ht.nodes] == ["10", "11", "12"]
    anchor = next(n for n in ht.nodes if n["label"] == ANCHOR_LABEL)
    assert anchor["io_node"] == "Input"


def test_오프셋은_9다():
    assert LABEL_OFFSET == 9
    assert _shift("1") == "10"
    assert _shift("21") == "30"


def test_숫자가_아니면_그대로():
    """`?` · `@/3` 같은 것은 노드 라벨이 아니다."""
    assert _shift("?") == "?"
    assert _shift("@/3") == "@/3"
    assert _shift(None) == "None"


# ─────────────────────────────────────────────── ② 빠짐없이
def test_배관_in_out이_같이_옮겨진다():
    ht = to_head_tables(_sample())
    assert [(p["in"], p["out"]) for p in ht.pipes] == [("10", "11"), ("11", "12")]


def test_노즐_부속_기기도_옮겨진다():
    ht = to_head_tables(_sample())
    assert ht.nozzles[0]["in"] == "12"
    assert (ht.fittings[0]["in"], ht.fittings[0]["out"]) == ("11", "12")
    assert (ht.equipment[0]["in"], ht.equipment[0]["out"]) == ("10", "11")


# ─────────────────────────────────────────────── ③ 옮기면 안 되는 것
def test_배관_이름은_옮기지_않는다():
    """`label`·`pipe` 는 배관 이름이지 노드 라벨이 아니다."""
    ht = to_head_tables(_sample())
    assert [p["label"] for p in ht.pipes] == ["P1", "P2"]
    assert ht.fittings[0]["pipe"] == "P2"
    assert ht.equipment[0]["pipe"] == "P1"


def test_노즐_out은_노즐참조라_불변():
    ht = to_head_tables(_sample())
    assert ht.nozzles[0]["out"] == "@/1"


def test_원본을_건드리지_않는다():
    g = _sample()
    to_head_tables(g)
    assert g.nodes[0]["label"] == "1", "제자리에서 고치면 같은 세션의 표가 흔들린다"
    assert g.pipes[0]["in"] == "1"


def test_그밖의_칸은_보존된다():
    ht = to_head_tables(_sample())
    assert ht.pipes[0]["dia"] == 65
    assert ht.nodes[1]["x"] == 100
    assert ht.meta == [("제목", "시험")]


# ─────────────────────────────────────────────── ④ 성립하지 않으면 올린다
def test_표가_없으면_올린다():
    with pytest.raises(MergeError):
        to_head_tables(None)


def test_급수원이_뿌리가_아니면_올린다():
    g = _sample()
    g.nodes[0]["io_node"] = "No"
    with pytest.raises(MergeError, match="급수원"):
        to_head_tables(g)


def test_고아_참조를_잡는다():
    """옮기다 한 자리를 빠뜨리면 PIPENET 이 조용히 Unset 으로 읽는다."""
    g = _sample()
    g.pipes.append({"label": "P9", "in": "2", "out": "99"})
    with pytest.raises(MergeError, match="없는 절점"):
        to_head_tables(g)


def test_기준점이_없으면_올린다():
    g = _sample()
    # 라벨을 죄다 옮겨 10 이 안 생기게 만든다
    for n in g.nodes:
        n["label"] = str(int(n["label"]) + 100)
    for p in g.pipes:
        p["in"] = str(int(p["in"]) + 100)
        p["out"] = str(int(p["out"]) + 100)
    for r in g.nozzles:
        r["in"] = str(int(r["in"]) + 100)
    for r in g.fittings + g.equipment:
        r["in"] = str(int(r["in"]) + 100)
        r["out"] = str(int(r["out"]) + 100)
    with pytest.raises(MergeError, match="기준점"):
        to_head_tables(g)


# ─────────────────────────────────────────────── S710 급수방식
def test_급수방식_4종():
    assert set(SUPPLY_MODES) == {"hsp_pump", "lsp_gravity", "lsp_1stage",
                                 "llsp_2stage"}


def test_급수방식은_자동추정하지_않는다():
    for bad in (None, "", "  ", "펌프", "pump"):
        with pytest.raises(MergeError, match="급수방식"):
            check_supply_mode(bad)


def test_급수방식이_엔진_ZoneType과_1대1():
    """이름이 둘이면 어느 쪽이 권위인지를 매번 정해야 한다 — 하나로 둔다."""
    from remote30_full_network import ZoneType
    assert {z.value for z in ZoneType} == set(SUPPLY_MODES)
    for m in SUPPLY_MODES:
        assert zone_type_of(m) == ZoneType(m)


def test_라이저_빌더가_넷다_기준점_10을_세운다():
    """S740 이 성립하려면 라이저 쪽 AV 도 10 이어야 한다."""
    import remote30_full_network as fn
    for zt, builder in fn._RISER_BUILDERS.items():
        src = __import__("inspect").getsource(builder)
        assert 'av_node_label="10"' in src, f"{zt} 의 AV 라벨이 10 이 아니다"
