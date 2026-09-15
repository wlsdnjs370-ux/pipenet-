# -*- coding: utf-8 -*-
"""[가지치기·부속판정·두사전 §4-3] 지워진 갈래의 티는 건물에 그대로 있다.

지시서 `ModuleF_가지치기_부속판정_두사전_지시서.md` · 그림 16·17.

■ 한 줄

  헤드로 물을 나르지 않는 배관은 회랑에서 지우되, **부속 «종류» 는 지우기 전의
  배관망(손질 정본 G)의 차수**로 정한다. 판정은 한 벌, 표기는 두 벌.

■ 여기서 못박는 것

  F2  판정이 **가지치기의 결과에 안 기댄다** — 같은 회랑이라도 G 에 갈래가
      있었으면 티, 없었으면 없음/엘보다. 지금은 가지치기가 차수를 바꿔 그
      줄을 다시 정하고 있었다(그림 10 의 다섯 번째 ✕).
  F3  표기 두 벌 — `.sdf` 는 PIPENET 어휘, `.kfp` 는 K-Solver 어휘. 어느 쪽에도
      사전 밖 값이 없다. 직류티는 `.sdf` 에 **안 실린다**.
  F4  숫자 하나 — `.kfp` 의 `equivalent_length` = 표의 `eq_len`.
  F6  관경 — `.kfp` 의 `nominal_mm` = 배관표의 `dia`.
  그리고 `phys` 를 **안 주면 종전과 한 글자도 같게** 동작한다(전체망 보호).
"""
from __future__ import annotations

import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core"),
           os.path.join(_ROOT, "cad_project_editor_g"),
           os.path.join(_ROOT, "scripts"), os.path.join(_ROOT, "tests")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import fitting_rules as fr                                    # noqa: E402
from services.cad_import.design.fitting import (              # noqa: E402
    build_fittings, fitting_label, load_fitting_dict, resolve_eq_len)


# ═══════════════════════════════════════════ ① 어휘와 판정 함수
def test_편향각을_0_45_90_중_가까운_쪽으로_스냅한다():
    """오너 2026-09-15 — 「0, 45, 90 도 중 더 가까운 임계로 직진 처리」.

    45° 한 자리에서 자르면 46° 와 89° 가 한 통에 든다. 고를 수 있는 형상이
    사실상 셋뿐이므로 경계도 그 셋의 **중점**(22.5 · 67.5)에 둔다 — 엘보
    규칙이 이미 쓰는 그 경계다.
    """
    assert fr.snap_turn_deg(0) == 0.0
    assert fr.snap_turn_deg(22.5) == 0.0
    assert fr.snap_turn_deg(23) == 45.0
    assert fr.snap_turn_deg(67.5) == 45.0
    assert fr.snap_turn_deg(68) == 90.0
    assert fr.snap_turn_deg(180) == 90.0

    import math
    node, up = (0.0, 0.0), (-1.0, 0.0)

    def side(deg):
        r = math.radians(deg)
        turn, run, _bad = fr.tee_split(
            node, up, [("a", (math.cos(r), math.sin(r)))])
        return "직진" if run else "꺾임"

    # ★46°·60° 는 종전에 «꺾임» 이었다 — 45° 에 더 가까우므로 이제 직진이다.
    assert [side(d) for d in (0, 20, 30, 46, 60)] == ["직진"] * 5
    assert [side(d) for d in (70, 90, 135)] == ["꺾임"] * 3


def test_tee_split_의_꺾인_갈래는_tee_fittings_와_같다():
    """새 함수가 옛 규칙을 **바꾸지 않았다** — 직진 갈래를 더 돌려줄 뿐이다.

    ★단 하나 일부러 다르다: 45° < 편향 ≤ 67.5° 는 `tee_fittings` 가 «꺾임»,
      `tee_split` 은 0/45/90 스냅 때문에 «직진» 이다(오너 2026-09-15).
      아래 여섯 벌은 그 구간에 걸치지 않아 둘이 같아야 한다 — 같지 않으면
      스냅 말고 **다른 것**이 바뀐 것이다.
    """
    cases = [
        ((0, 0), (-1, 0), [("a", (1, 0)), ("b", (0, 1))]),
        ((0, 0), (-1, 0), [("a", (0, 1)), ("b", (0, -1))]),
        ((0, 0), None, [("a", (1, 0)), ("b", (0, 1))]),
        ((0, 0), (-1, 0), [("a", (1, 0)), ("b", (1, 1))]),
        ((0, 0), (-1, 0), [("a", (1, 0)), ("b", (0, 1)), ("c", (0, -1))]),
        ((0, 0), (0, -1), [("a", (0, 1)), ("b", (1, 0))]),
    ]
    for node, up, downs in cases:
        turn_old, bad_old = fr.tee_fittings(node, up, downs)
        turn_new, _run, bad_new = fr.tee_split(node, up, downs)
        assert turn_new == turn_old, f"{node} {up} {downs}"
        assert bad_new == bad_old


def test_하류가_하나여도_판정한다():
    """갈래가 지워진 분기점이 바로 하류 1개짜리 분기점이다.

    `tee_fittings` 는 그것을 「분기가 아니다」로 보고 빈 값을 돌려준다 —
    차수는 부르는 쪽이 `phys` 로 이미 가렸으므로 여기서 또 가리면 안 된다.
    """
    node, up = (0.0, 0.0), (-1.0, 0.0)
    assert fr.tee_fittings(node, up, [("a", (0.0, 1.0))]) == ([], 0)
    turn, run, bad = fr.tee_split(node, up, [("a", (0.0, 1.0))])
    assert turn == ["a"] and run == [] and bad == 0      # 꺾임 → 분류티
    turn, run, bad = fr.tee_split(node, up, [("a", (1.0, 0.0))])
    assert turn == [] and run == ["a"] and bad == 0      # 직진 → 직류티


def test_기존_어휘와_상수는_그대로다():
    """§5 금지 — `fitting_rules` 의 기존 함수·상수 diff 0."""
    assert fr.TEE == "tee" and fr.ELBOW_90 == "elbow" and fr.ELBOW_45 == "elbow-45"
    assert fr.TRUNK_TURN_TOL_DEG == 45.0
    assert fr.ELBOW_STRAIGHT_MAX_DEG == 22.5
    assert fr.ELBOW_45_MAX_DEG == 67.5
    assert fr.ELBOW_MAX_DEG == 95.0
    # 새 어휘
    assert (fr.TEE_RUN, fr.CROSS, fr.CROSS_RUN) == ("tee-run", "cross", "cross-run")


# ═══════════════════════════════════════════ ② 합성 망 — 그림 16
def _net(pipes):
    """{pid: (start, end)} → 최소 kfp. 좌표는 노드 이름이 정한다."""
    nodes = {}
    for a, b in pipes.values():
        for n in (a, b):
            nodes.setdefault(n, {"coords": list(_XY[n]) + [0.0],
                                 "type_id": "base"})
    return {"nodes_meta_runtime": nodes,
            "pipe_data": {pid: {"start": a, "end": b, "length_m": 1.0}
                          for pid, (a, b) in pipes.items()}}


#   W ── E ── H1 ── C ── X        (동쪽으로 직진)
#             │           ╲
#             (아래 Hd)     N (C 에서 북쪽)
_XY = {
    "W": (0.0, 0.0), "E": (1.0, 0.0), "H1": (2.0, 0.0),
    "C": (3.0, 0.0), "X": (4.0, 0.0), "N": (3.0, 1.0),
    "S": (3.0, -1.0), "Hd": (1.0, -1.0),
}
_LINE = {"P1": ("W", "E"), "P2": ("E", "H1"), "P3": ("H1", "C"),
         "P4": ("C", "N")}
_PARENT = {"E": "W", "H1": "E", "C": "H1", "N": "C"}
_XYD = {k: (v[0], v[1]) for k, v in _XY.items()}
_BORES = {p: (50, "시험") for p in ("P1", "P2", "P3", "P4", "P5")}


def _kinds(res, pid):
    return sorted(res["per_pipe"][pid]["fittings"])


def test_갈래가_지워진_분기점이_직진이면_직류티다():
    """E 는 G 에서 차수 3(아래로 Hd)이었고, 회랑에서 흐름은 직진 — 직류티."""
    net = _net(_LINE)
    res = build_fittings(net, _XYD, _BORES, parents=_PARENT,
                         phys={"E": 3})
    assert _kinds(res, "P2") == ["tee-run"], res["per_pipe"]["P2"]


def test_갈래가_지워진_분기점이_꺾이면_분류티다():
    """C 는 G 에서 차수 3 이었고 흐름이 북으로 꺾인다 — 분류티(지금은 엘보 ✗)."""
    net = _net(_LINE)
    res = build_fittings(net, _XYD, _BORES, parents=_PARENT,
                         phys={"C": 3})
    assert _kinds(res, "P4") == ["tee"], res["per_pipe"]["P4"]


def test_갈래가_없이_꺾이면_엘보다():
    """F2 의 짝 — 같은 자리라도 G 에 갈래가 없었으면 엘보다."""
    net = _net(_LINE)
    res = build_fittings(net, _XYD, _BORES, parents=_PARENT,
                         phys={"C": 2})
    assert _kinds(res, "P4") == ["elbow"], res["per_pipe"]["P4"]


def test_F2_판정이_가지치기의_결과에_안_기댄다():
    """★이 저장소가 고치려던 그 결함.

    **같은 회랑**(같은 kfp · 같은 흐름)인데 G 의 차수만 다르면 부속이 달라야
    한다. 지금 차수(`len(links)`)로 정하면 둘이 **같아진다** — 그것이 그림 10
    의 다섯 번째 ✕ 다.
    """
    net = _net(_LINE)
    with_branch = build_fittings(net, _XYD, _BORES, parents=_PARENT,
                                 phys={"E": 3, "C": 3})
    without = build_fittings(net, _XYD, _BORES, parents=_PARENT,
                             phys={"E": 2, "C": 2})
    assert _kinds(with_branch, "P2") == ["tee-run"]
    assert _kinds(without, "P2") == []
    assert _kinds(with_branch, "P4") == ["tee"]
    assert _kinds(without, "P4") == ["elbow"]
    assert with_branch["counts"] != without["counts"]


def test_십자는_이름만_다르고_규칙은_티와_같다():
    """phys ≥ 4 — 직진 갈래는 `cross-run`, 꺾인 갈래는 `cross`."""
    pipes = dict(_LINE)
    pipes["P5"] = ("C", "S")                  # C 에서 남쪽으로 하나 더
    pipes["P6"] = ("C", "X")                  # C 에서 동쪽(직진)
    net = _net(pipes)
    parents = dict(_PARENT)
    parents["S"] = "C"
    parents["X"] = "C"
    res = build_fittings(net, _XYD, dict(_BORES, P6=(50, "시험")),
                         parents=parents, phys={"C": 4})
    assert _kinds(res, "P6") == ["cross-run"], res["per_pipe"]["P6"]
    assert _kinds(res, "P4") == ["cross"]
    assert _kinds(res, "P5") == ["cross"]
    # 등가길이는 티와 같다(D1) — `cross` 는 TEE_BRANCH 를 본다.
    assert resolve_eq_len("cross", 50)[0] == resolve_eq_len("tee", 50)[0]


def test_D6_덮인_분기점은_덮는_배관에_직류티로_남는다():
    """노드정리가 지운 분기점 — 그 티도 건물에는 그대로 있다."""
    net = _net(_LINE)
    res = build_fittings(net, _XYD, _BORES, parents=_PARENT,
                         phys={"E": 2}, interior_junctions={"P3": 2})
    assert _kinds(res, "P3") == ["tee-run", "tee-run"]
    # 등가길이 0 — 그래서 수치는 안 움직인다.
    assert res["per_pipe"]["P3"]["equivalent_length"] == 0.0


# ═══════════════════════════════════════════ ③ 전체망 보호
def test_phys_를_안_주면_종전과_같다():
    """★새 규칙은 **`phys` 를 받았을 때만** 켜진다.

    한 번 그러지 못했다: `phys` 없이 불렀는데 `tee-run` 32건이 생기고
    `tee` 가 34 → 36 이 됐다(실측 대명동 K=30). 문이 하나여야 한다.
    """
    pipes = dict(_LINE)
    pipes["P5"] = ("C", "S")
    pipes["P6"] = ("C", "X")
    net = _net(pipes)
    parents = dict(_PARENT, S="C", X="C")
    bores = dict(_BORES, P6=(50, "시험"))
    old = build_fittings(net, _XYD, bores, parents=parents)
    # 지금 차수로 보면 C 는 4 지만, phys 가 없으므로 «티» 여야 한다(크로스 ✗)
    got = set(old["counts"])
    assert "cross" not in got and "cross-run" not in got and "tee-run" not in got
    # 직진 갈래는 종전대로 **버려진다**
    assert _kinds(old, "P6") == []


# ═══════════════════════════════════════════ ④ 두 사전
def test_사전은_라벨만_갖는다():
    """숫자·라이브러리 id 를 사전에 적으면 주인이 둘이 된다(§5 금지)."""
    for fmt in ("sdf", "kfp"):
        d = load_fitting_dict(fmt)
        assert set(d["map"]) >= {"elbow", "elbow-45", "tee", "cross",
                                 "tee-run", "cross-run", "none"}
        blob = repr(d)
        assert "TEE_BRANCH" not in blob and "ELBOW_90_STD" not in blob


def test_직류티는_sdf_에_안_실리고_kfp_에는_실린다():
    """PIPENET 은 직류티를 «말할 수 없고», K-Solver 는 «말할 수 있다»(그림 17)."""
    assert fitting_label("tee-run", "sdf") is None
    assert fitting_label("cross-run", "sdf") is None
    assert fitting_label("tee-run", "kfp") == "Tee(Run)"
    assert fitting_label("cross-run", "kfp") == "Tee(Run)"
    # 십자는 PIPENET 에서 티로 적는다(어휘가 없다)
    assert fitting_label("cross", "sdf") == "tee"
    assert fitting_label("cross", "kfp") == "Tee(Branch)"


def test_사전에_없는_종류는_던진다():
    """조용히 빠뜨리면 그 부속이 산출에서 소리 없이 사라진다."""
    with pytest.raises(KeyError):
        fitting_label("없는부속", "sdf")
    with pytest.raises(KeyError):
        fitting_label("없는부속", "kfp")


def test_직류티의_0_은_규칙이_정한_0_이지_미해결이_아니다():
    """미해결은 «값을 모른다» 이고 이것은 «값이 0 이라고 규칙이 정했다» 다."""
    m, why = resolve_eq_len("tee-run", 50)
    assert m == 0.0 and why == "규칙"
    m, why = resolve_eq_len("cross-run", None)     # 호칭경을 몰라도 0 이다
    assert m == 0.0 and why == "규칙"
    # 어휘에 **없는** 종류는 종전대로 미해결
    assert resolve_eq_len("없는부속", 50) == (None, None)


def test_규칙이_정한_0_은_직접입력으로_세지_않는다():
    """「직접 입력 n건」이 거짓이 되면 안 된다(§3-4 6)."""
    net = _net(_LINE)
    res = build_fittings(net, _XYD, _BORES, parents=_PARENT, phys={"E": 3})
    assert _kinds(res, "P2") == ["tee-run"]
    assert res["unresolved_length"] == 0
    for a in (res.get("applied_overrides") or ()):
        assert a.get("note") != "규칙"


# ═══════════════════════════════════════════ ⑤ 실도면 — 대명동 K=30
_DXF = os.path.join(_ROOT, "routes", "제출용[최종]",
                    "1. 입력도면 대명동 단위세대 평면도.dxf")


@pytest.mark.skipif(not os.path.exists(_DXF), reason="대명동 도면이 없다")
def test_실도면에서_F3_F4_F6_가_선다():
    """끝에서 끝까지 — 표 → `.sdf` · `.kfp` 가 각자의 어휘로 같은 부속을 적는다."""
    import argparse
    import io
    import contextlib
    import json
    import xml.etree.ElementTree as ET

    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _probe_candidate_drop import setup, wait
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    buf = io.StringIO()
    with isolated_workdir(prefix="pf_"), srv.app.test_client() as c:
        with contextlib.redirect_stdout(buf):
            with c.session_transaction() as s:
                s["authed"] = True
            sid = setup(c, argparse.Namespace(key=""))
            st = c.get(f"/api/module-f/edit/state?sid={sid}").get_json()["state"]
            if not (st.get("sources") or ()):
                seg = sorted((g.get("segs") or [] for g in st["body_groups"]),
                             key=len, reverse=True)[0]
                c.post("/api/module-f/edit/mode",
                       json={"sid": sid, "mode": "급수시작위치"})
                c.post("/api/module-f/edit/click",
                       json={"sid": sid, "x": (seg[0] + seg[2]) / 2,
                             "y": (seg[1] + seg[3]) / 2, "max_d": 2000})
            c.post("/api/module-f/edit/worst", json={"sid": sid, "k": 30})
            c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
            assert wait(c, sid).get("state") == "done"
            c.post("/api/module-f/design/emit", json={"sid": sid})
            c.post("/api/module-f/convert/run",
                   json={"sid": sid, "outputs": {"full_kfp": False,
                                                 "worst_kfp": True,
                                                 "worst_sdf": False}})
            assert wait(c, sid).get("state") == "done"
            from routes.module_f.jobs import _sess
            sess = _sess(sid)
            tbl = sess["design"]["tables"]
            sdf_p, kfp_p = sess["design_sdf_path"], sess["worst_kfp_path"]

        # F3 — 표기 두 벌
        root = ET.parse(sdf_p).getroot()
        sdf_types = {f.get("type") for f in root.iter("Fitting")}
        assert sdf_types <= {"tee", "elbow", "elbow-45", "gate",
                             "butterfly", "check"}, sdf_types
        assert "tee-run" not in sdf_types and "cross" not in sdf_types

        pd = (json.load(open(kfp_p, encoding="utf-8")).get("pipe_data") or {})
        kfp_types = {x for p in pd.values() for x in (p.get("fittings") or ())}
        assert kfp_types <= {"Elbow", "Elbow 45", "Tee(Branch)",
                             "Tee(Run)"}, kfp_types

        # F6 · F4 — 관경과 등가길이가 표와 같다 (D5)
        dia = {str(r["label"]): int(r["dia"] or 0) for r in tbl.pipes}
        eq = {str(r["label"]): float(r.get("eq_len") or 0) for r in tbl.pipes}
        assert set(map(str, pd)) == set(dia), "표와 .kfp 가 다른 배관망이다"
        for pid, rec in pd.items():
            assert int(rec.get("nominal_mm") or 0) == dia[str(pid)], pid
            assert abs(float(rec.get("equivalent_length") or 0)
                       - eq[str(pid)]) <= 0.001, pid

        # 새 어휘가 실제로 나왔다 — 「아무것도 안 바뀐 통과」를 막는다
        kinds = {str(r["type"]) for r in tbl.fittings}
        assert "tee-run" in kinds, kinds
