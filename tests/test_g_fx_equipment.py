# -*- coding: utf-8 -*-
"""[§29] 신축배관(FX) — 사람이 켰을 때만, 모듈 A 와 같은 규칙으로.

■ 무엇이 비어 있었나

  권위 레퍼런스는 헤드 접속관마다 **FX 등가길이 15.6m** 를 싣는다. 우리 망에
  그 규약을 대면 헤드 32개 × 15.6 = 499.2m 로 **배관 총연장의 372%** 인데,
  모듈 F 산출물에는 0개였다. 등가길이는 마찰손실에 배관 길이와 같은 자격으로
  들어가므로 그만큼 계산이 낙관적이 된다.

  모듈 A 는 이 규칙을 갖고 있었고(`FX_SPEC_PROFILES`), G/F 의 표 조립에는
  자리조차 없었다.

■ 이 시험이 지키는 셋

  ⑴ **기본은 «안 함»** — 신축배관을 쓰는 현장인지는 설계 결정이다. 켜지 않으면
     기존 산출물이 한 바이트도 안 바뀐다(D-F11-1).
  ⑵ **규칙은 한 벌** — 수치를 F 가 다시 적지 않고 A 와 같은 공용 상수를 쓴다.
  ⑶ **하향식 계열에만** — 신축배관은 세대 안 하향식 헤드에 붙는 물건이고,
     상향식은 촛대로 직결되므로 안 단다(A 의 판단 그대로).
"""
from __future__ import annotations

import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "cad_project_editor_g"),
           os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.design.tables import (          # noqa: E402
    _fx_spec, build_design_tables)


def _net(kinds):
    """접속점 — 본관 — 헤드 n개. 헤드마다 종류를 지정한다."""
    nodes = {"N1": {"coords": [1.0, 1.0, 0.0], "type_id": "pump"},
             "N2": {"coords": [5.0, 1.0, 0.0], "type_id": "base"}}
    pipes = {"P1": {"start": "N1", "end": "N2", "length_m": 4.0}}
    nhk = {}
    for i, k in enumerate(kinds, start=1):
        nid = f"H{i}"
        nodes[nid] = {"coords": [5.0 + i, 1.0, 0.6], "type_id": "head"}
        pipes[f"PH{i}"] = {"start": "N2", "end": nid, "length_m": 0.7}
        nhk[nid] = k
    return {"pipe_data": pipes, "nodes_meta_runtime": nodes,
            "node_head_kinds": nhk}


def _build(kinds, fx=None):
    net = _net(kinds)
    return build_design_tables(
        net, {"heads": [], "loads": {}}, {}, [],
        bores={k: (25, "시험") for k in net["pipe_data"]},
        tree_loads={"P1": len(kinds)}, fx_profile=fx,
        # ★종류표는 전개 결과(`got`)에서 온다 — `net`(kfp) 안에는 없다.
        #   처음에 net 에서 읽어 늘 비었고 FX 가 한 개도 안 붙었다.
        node_head_kinds=net["node_head_kinds"])


def test_기본은_안_단다():
    """★켜지 않으면 산출이 한 줄도 안 바뀐다 — 골든이 흔들리지 않는 근거."""
    t = _build(["하향식", "하향식"])
    assert [e for e in t.equipment if e["desc"] == "FX"] == []
    assert dict(t.meta)["신축배관(FX)"] == "안 함"


def test_켜면_하향식_계열에만_붙는다():
    """상향식은 촛대 직결이라 신축배관이 없다 — A 의 판단 그대로."""
    t = _build(["하향식", "상향식", "상하향식", "미지정"], fx="평균")
    fx = [e for e in t.equipment if e["desc"] == "FX"]
    assert len(fx) == 2, [e["pipe"] for e in fx]
    assert {e["pipe"] for e in fx} == {"PH1", "PH3"}


def test_수치는_모듈_A_와_같은_표에서_온다():
    """★F 가 숫자를 다시 적지 않는다 — 두 벌이면 한쪽만 고쳐지는 날이 온다."""
    from remote30_constants import FX_DEFAULT_PROFILE, FX_SPEC_PROFILES

    assert _fx_spec("평균") is FX_SPEC_PROFILES["평균"]
    t = _build(["하향식"], fx="평균")
    fx = [e for e in t.equipment if e["desc"] == "FX"][0]
    assert fx["eq_len"] == FX_SPEC_PROFILES["평균"]["eq_len_m"]
    assert FX_DEFAULT_PROFILE in FX_SPEC_PROFILES


def test_모르는_규격은_조용히_기본값으로_안_떨어진다():
    """오타가 기본값이 되면 사람은 제가 고른 줄 안다 — 사유를 들고 던진다."""
    with pytest.raises(ValueError) as e:
        _build(["하향식"], fx="없는규격")
    msg = str(e.value)
    assert "없는규격" in msg and "평균" in msg


def test_무엇을_달았는지_산출물이_말한다():
    """자동이 낸 값과 사람이 켠 값을 같은 얼굴로 두지 않는다(§18 원칙)."""
    t = _build(["하향식", "하향식"], fx="한백표준")
    note = dict(t.meta)["신축배관(FX)"]
    assert "한백표준" in note and "2개" in note, note
    fx = [e for e in t.equipment if e["desc"] == "FX"][0]
    assert fx["eq_len_src"] == "신축배관 한백표준"
    assert fx["spec_ref"] == "한백표준"


def test_고르는_자리도_같은_표를_따른다():
    """★화면 드롭다운은 «세 번째 벌» 이 되기 쉽다.

    엔진과 상수는 한 벌로 묶었는데 고르는 자리에 규격 이름을 손으로 적어 두면,
    표에 규격이 하나 늘어도 화면에는 안 뜨고 이름이 바뀌면 «모르는 규격» 으로
    던진다. 값(=서버로 가는 것)만 본다 — 설명 글귀는 사람 몫이다.
    """
    import re

    from remote30_constants import FX_SPEC_PROFILES

    html = open(os.path.join(_ROOT, "templates", "module_f.html"),
                encoding="utf-8").read()
    m = re.search(r'<select id="dg-fx">(.*?)</select>', html, re.S)
    assert m, "FX 를 고르는 자리를 못 찾았다 — id 가 바뀌었나"
    vals = re.findall(r'<option value="([^"]*)"', m.group(1))
    assert vals[0] == "", "기본값 «안 함» 이 첫 자리가 아니다"
    assert set(vals[1:]) == set(FX_SPEC_PROFILES), (vals, list(FX_SPEC_PROFILES))


def test_붙인_것을_화면에서_볼_자리가_있다():
    """★meta 한 줄로만 「켜졌다」고 말하면 사람은 확인할 길이 없다.

    기기표는 이미 화면으로 보내고 있었는데(`as_dict()["equipment"]`) 고르는
    자리에 없어서 **볼 수가 없었다**. 표 이름은 서버 키와 한 글자도 달라선
    안 된다 — 다르면 그 표는 늘 빈 채로 뜬다(`tables[which] || []`).
    """
    import re

    t = _build(["하향식"], fx="평균")
    keys = set(t.as_dict())
    html = open(os.path.join(_ROOT, "templates", "module_f.html"),
                encoding="utf-8").read()
    m = re.search(r'<select id="dg-table">(.*?)</select>', html, re.S)
    assert m, "표를 고르는 자리를 못 찾았다 — id 가 바뀌었나"
    vals = re.findall(r'<option value="([^"]+)"', m.group(1))
    assert "equipment" in vals, f"기기표를 볼 자리가 없다: {vals}"
    assert not (set(vals) - keys), f"서버에 없는 표를 고르게 뒀다: {set(vals) - keys}"


def test_기기는_물이_지나는_관에_붙는다():
    """알람밸브와 같은 규약 — 담당 헤드 수로 고른다."""
    t = _build(["하향식"], fx="평균")
    fx = [e for e in t.equipment if e["desc"] == "FX"][0]
    # 헤드 접속관(PH1)에 붙어야 한다 — 본관(P1)이 아니다.
    assert fx["pipe"] == "PH1", fx["pipe"]
