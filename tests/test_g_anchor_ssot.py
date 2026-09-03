# -*- coding: utf-8 -*-
"""접속점(앵커) 하나 · 이름 하나 · 등가길이 한 자리.

이 파일이 지키는 넷은 한 뿌리에서 나왔다 — 「평면도의 급수원과 알람밸브는 같은
장치인가」. 같다. 통합(S740)에서 평면도 쪽 접속점 노드는 라이저의 알람밸브
노드에 자리를 내주고 사라진다(`remote30_full_network`: 「AV 는 라이저 쪽에서
이미 포함 — 헤드망 쪽 사본 skip」). `_COORDS_AV` 의 주석도 그 노드를 「헤드망
source ★」라고 적어 두었다.

  ① 평면도에서 찍는 특수 점은 하나 — 알람밸브가 접속점을 겸한다
  ② 뿌리는 접속점이고, 없으면 «아무 노드나» 고르지 않고 던진다
  ③ 등가길이는 한 함수가 정한다 — 기기표의 알람밸브도 부속표와 같은 자리
  ④ 「앵커」는 접속점만 뜻한다 — 반대쪽 끝은 «기준 헤드»
"""
from __future__ import annotations

import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
if _G not in sys.path:
    sys.path.insert(0, _G)

from services.cad_import.design.anchor import (            # noqa: E402
    AnchorMissing, find_anchor, require_anchor)
from services.cad_import.design.tables import build_design_tables  # noqa: E402
from services.cad_import.design.worst import worst_k_heads          # noqa: E402


def _net(*, with_anchor=True):
    nodes = {
        "N1": {"coords": [0.0, 0.0, 0.0],
               "type_id": "pump" if with_anchor else "base"},
        "N2": {"coords": [2.0, 0.0, 0.0], "type_id": "base"},
    }
    return {"pipe_data": {"P1": {"start": "N1", "end": "N2", "length_m": 2.0}},
            "nodes_meta_runtime": nodes}


# ═════════════════════════════ ② 뿌리 = 접속점, 없으면 던진다
def test_접속점을_못_찾으면_아무_노드나_고르지_않는다():
    """★종전엔 `next(iter(nodes))` 로 눕었다.

    dict 에 먼저 들어온 노드가 Input 경계가 되고, 물 흐르는 방향이 통째로
    거기서 유도된다 — 사람이 찍은 자리와 아무 상관이 없는 지점에서 계산이
    시작되는데도 표는 «틀렸다» 고 말해 주지 않는다.
    """
    net = _net(with_anchor=False)
    assert find_anchor(net["nodes_meta_runtime"]) is None
    with pytest.raises(AnchorMissing) as e:
        build_design_tables(net, {"heads": [], "loads": {}}, {"P1": (0, 1)}, [])
    msg = str(e.value)
    assert "알람밸브" in msg and "찍으세요" in msg, msg


def test_접속점이_있으면_그것이_Input_이다():
    net = _net()
    tbl = build_design_tables(net, {"heads": [], "loads": {}},
                              {"P1": (0, 1)}, [])
    ins = [n for n in tbl.nodes if n.get("io_node") == "Input"]
    assert len(ins) == 1, ins
    assert ins[0].get("pressure_pa") == 101325.0


def test_require_anchor_는_사람이_읽을_문장으로_던진다():
    with pytest.raises(AnchorMissing) as e:
        require_anchor({}, what="관경 산정")
    assert "관경 산정" in str(e.value)


# ═════════════════════════════ ③ 등가길이는 한 자리에서 정한다
def test_알람밸브_등가길이가_라이브러리에서_온다():
    """★종전엔 기기표가 `"eq_len": 0.0` 으로 박혀 있었다.

    라이브러리에 값이 버젓이 있는데도(실측: 알람밸브 100A = 9.5m) SDF 에는 0 이
    실렸다 — 0 은 「손실이 없다」는 **주장**이라 그만큼 계산이 낙관적으로 틀어진다.
    """
    tbl = build_design_tables(_net(), {"heads": [], "loads": {}},
                              {"P1": (0, 1)}, [],
                              bores={"P1": (100, "시험")}, valve_nodes=["N2"])
    av = [e for e in tbl.equipment if e["desc"] == "A/V"]
    assert len(av) == 1, tbl.equipment
    assert av[0]["eq_len"] > 0, "알람밸브 등가길이가 0 이다"
    assert av[0]["eq_len_src"] == "라이브러리"


def test_등가길이를_못_구하면_0_으로_때우지_않고_신고한다():
    """라이브러리에 그 호칭경이 없으면(알람밸브 15A) «미해결» 로 남긴다.

    값은 0 으로 두되 **어디가 빈지** 목록에 넣는다 — 0 을 조용히 쓰면 사람이
    채울 자리를 영영 모른다.
    """
    tbl = build_design_tables(_net(), {"heads": [], "loads": {}},
                              {"P1": (0, 1)}, [],
                              bores={"P1": (15, "시험")}, valve_nodes=["N2"])
    assert dict(tbl.meta)["등가길이 미해결"] == "1"
    items = tbl.unresolved["length_items"]
    assert any(i.get("kind") == "alarm_valve" for i in items), items
    av = [e for e in tbl.equipment if e["desc"] == "A/V"][0]
    assert "eq_len_src" not in av, "못 구했는데 근거가 붙었다"


def test_사람이_채운_값은_기기표에도_먹는다():
    """부속표와 기기표가 **같은 칸**을 본다 — 채울 자리가 하나여야 한다."""
    tbl = build_design_tables(
        _net(), {"heads": [], "loads": {}}, {"P1": (0, 1)}, [],
        bores={"P1": (15, "시험")}, valve_nodes=["N2"],
        fitting_overrides={"eq_len": [{"kind": "alarm_valve", "dia": 15,
                                       "m": 6.5, "note": "KFI"}]})
    av = [e for e in tbl.equipment if e["desc"] == "A/V"][0]
    assert av["eq_len"] == 6.5 and av["eq_len_src"] == "KFI"
    assert dict(tbl.meta)["등가길이 미해결"] == "0"


# ═════════════════════════════ ① 픽은 하나 · ④ 이름은 갈렸다
def test_평면도_픽은_하나이고_접속점을_겸한다():
    """알람밸브를 찍으면 접속점이 따라온다 — 어긋날 수가 없다."""
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import (
        EditSession, MODE_SOURCE, MODE_VALVE)

    pts = [(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)]
    b = EditBoard("시험_접속점", pts, [(0, 1), (1, 2)], [])
    es = EditSession(b, key="시험_접속점")
    es.set_mode(MODE_VALVE)
    rep = es.click(1500.0, 0.0, 500.0)
    assert rep and rep["동작"] == "알람밸브", rep
    assert len(b.valves) == 1 and list(b.sources) == list(b.valves)

    # 은퇴한 모드로 들어와도 **같은 동작** 이다 — 옛 화면이 어긋남을 못 만든다.
    es.set_mode(MODE_SOURCE)
    es.click(1500.0, 0.0, 500.0)                 # 토글 → 꺼짐
    assert list(b.valves) == [] and list(b.sources) == []


def test_한_번_찍은_것은_한_번에_되돌아온다():
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.edit.session import EditSession, MODE_VALVE

    b = EditBoard("시험_되돌리기", [(0.0, 0.0), (3000.0, 0.0)], [(0, 1)], [])
    es = EditSession(b, key="시험_되돌리기")
    es.set_mode(MODE_VALVE)
    es.click(1500.0, 0.0, 500.0)
    assert (len(b.valves), len(b.sources)) == (1, 1)
    assert b.undo() is True
    assert (len(b.valves), len(b.sources)) == (0, 0), \
        "한 번 찍은 것이 한 번에 안 풀렸다 — 중간 상태가 남는다"


def test_앵커라는_낱말이_한_뜻만_갖는다():
    """★정반대 두 끝을 같은 이름으로 부르지 않는다.

    앵커 = 접속점(라이저가 붙는 자리, 물이 들어오는 쪽).
    반대쪽 끝(급수에서 가장 먼 헤드)은 «기준 헤드» 다.
    """
    # 소스에서 낱말을 찾지 않는다(주석을 코드로 오독한다) — 산출물을 본다.
    pts = [(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)]
    got = worst_k_heads(pts, [(0, 1), (1, 2)], [{2}], [0], k=1)
    for key in ("worst_head", "worst_path", "worst_path_m"):
        assert key in got, f"{key} 가 없다 — {sorted(got)}"
    for gone in ("anchor", "anchor_path", "anchor_path_m"):
        assert gone not in got, f"최불리 쪽이 아직 «{gone}» 를 쓴다"
