# -*- coding: utf-8 -*-
"""덱 「z축 및 피팅류」 3장 — 호(원호) 기호 셋과 수직 전개 기본값 (2026-09-20/21 오너 확정).

  · 기본값: 가지·호 높이차 0.5 · 상향 ① 0.5 · 하향 ① 0.3 ② 0.5 · 상하향 ① 0.3 ② 0.2 ③ 0.5 ④ 0.5
  · 그림 1 우회: 접속 2개 노드 위의 호 짝 → 사이 구간을 0.5 올렸다 내림, 세로 2 · 엘보 4
  · 그림 2 통과 갈래: 주배관 접속점은 평면 4방향이어도 미해결이 아니라 티, 꼭대기 티는 한 번
  · 짝 없는 호는 지어내지 않고 센다
"""
from __future__ import annotations

from routes.module_f.common import _boot
_boot()

from services.cad_import import dto
from services.cad_import.convert.engine import convert_to_kfp
from services.cad_import.design.fitting import build_fittings


def _node(i, x, y, tid="base"):
    rec = {"id": i, "coords": [float(x), float(y), 0.0], "elevation_m": 0.0,
           "type": "기본", "type_id": tid, "category_id": "",
           "k_factor_si": None, "head_spec_name": None, "required_pressure_bar": 0.0}
    if tid == "head":
        rec["type"] = "Head"
        rec["k_factor_si"] = 80.0
    return rec


def _pipe(a, b, length):
    return {"start": a, "end": b, "type": "KSD3507", "diameter": 27.6,
            "nominal_mm": 25, "length_m": float(length), "equivalent_length": 0.0,
            "C": 120.0, "roughness_mm": 0.1, "fittings": [], "flow_lpm": 0.0,
            "velocity_mps": 0.0, "headloss_m": 0.0}


def _line_kfp():
    """A–B–C–D–E 한 줄 (x = 0..4) + E 끝 헤드. 호는 B·D 에 앉는다."""
    nodes = {k: _node(k, x, 0.0) for k, x in (("A", 0), ("B", 1), ("C", 2), ("D", 3), ("E", 4))}
    nodes["H"] = _node("H", 5.0, 0.0, "head")
    pipes = {"AB": _pipe("A", "B", 1.0), "BC": _pipe("B", "C", 1.0), "CD": _pipe("C", "D", 1.0),
             "DE": _pipe("D", "E", 1.0), "EH": _pipe("E", "H", 1.0)}
    return {"nodes_meta_runtime": nodes, "pipe_data": pipes,
            "node_counter": {"N": 0}, "pipe_id_counter": 0}


# 왼쪽 호 ⊂ : 왼쪽 반원을 그렸다(90°→270°) → 열린 쪽은 오른쪽(0°)
ARC_OPEN_RIGHT = {"cx": 1.0, "cy": 0.0, "r": 0.2, "sa": 90.0, "sweep": 180.0}
# 오른쪽 호 ⊃ : 오른쪽 반원(270°→90°) → 열린 쪽은 왼쪽(180°)
ARC_OPEN_LEFT = {"cx": 3.0, "cy": 0.0, "r": 0.2, "sa": 270.0, "sweep": 180.0}
KINDS = {"node_head_kinds": {"H": "상향식"}, "head_kinds": [{"c": [5.0, 0.0], "kind": "상향식"}],
         "sources": [{"xy": [0.0, 0.0]}]}


def test_defaults_follow_the_deck():
    assert dto.BRANCH_DEFAULT_M == 0.5
    assert dto.UPRIGHT_DEFAULT_M == 0.5
    assert (dto.PENDANT_DEFAULT_M, dto.PENDANT_2_DEFAULT_M) == (0.3, 0.5)
    assert (dto.COMBO_1_DEFAULT_M, dto.COMBO_UP_DEFAULT_M,
            dto.COMBO_2_DEFAULT_M, dto.COMBO_3_DEFAULT_M) == (0.3, 0.2, 0.5, 0.5)
    # 화면 칸의 기본 문자열은 숫자와 한 벌이어야 한다(칸을 비우면 그 값이 들어간다).
    for key, _label, text, val in dto.FIELDS:
        if text:
            assert abs(float(text) - float(val)) < 1e-9, key


def test_arc_pair_on_a_through_pipe_lifts_the_run_between_them():
    out = convert_to_kfp({"kfp": _line_kfp(), "ho": [ARC_OPEN_RIGHT, ARC_OPEN_LEFT], **KINDS})
    assert out["ok"] is True, out["blockers"]
    st = out["stats"]
    assert st["n_jog_pairs"] == 1 and st["n_jog_unpaired"] == 0 and st["n_vert_jog"] == 2
    n = out["kfp"]["nodes_meta_runtime"]
    p = out["kfp"]["pipe_data"]
    z = lambda k: float(n[k]["coords"][2])
    # 양 끝(B·D)과 바깥(A·E)은 그대로, 사이(C)는 올라간다
    assert z("A") == 0.0 and z("B") == 0.0 and z("D") == 0.0 and z("E") == 0.0
    assert abs(z("C") - 0.5) < 1e-9
    # 세로관 둘: B 와 D 바로 위로 0.5 (헤드 H 의 상향 스텁은 별개)
    verts = [q for q in p.values() if abs(z(q["start"]) - z(q["end"])) > 1e-9
             and "H" not in (q["start"], q["end"])]
    assert len(verts) == 2 and all(abs(float(q["length_m"]) - 0.5) < 1e-9 for q in verts)
    tops = {q["end"] for q in verts}
    for t in tops:
        assert n[t]["type_id"] == "base" and abs(z(t) - 0.5) < 1e-9
    # 올라간 구간 B'–C–D' 는 평면 길이를 그대로 지닌다(길이 A + 1.0 이 되는 근거)
    run = [q for q in p.values() if q["start"] in tops or q["end"] in tops]
    flat_run = [q for q in run if abs(z(q["start"]) - z(q["end"])) < 1e-9]
    assert len(flat_run) == 2 and all(abs(float(q["length_m"]) - 1.0) < 1e-9 for q in flat_run)
    rep = out["kfp"]["arc_report"]
    assert rep["jog_pairs"] == 1 and rep["rise_m"] == 0.5


def test_jog_bends_read_as_four_elbows():
    out = convert_to_kfp({"kfp": _line_kfp(), "ho": [ARC_OPEN_RIGHT, ARC_OPEN_LEFT], **KINDS})
    kfp = out["kfp"]
    n, p = kfp["nodes_meta_runtime"], kfp["pipe_data"]
    node_xy = {k: (float(v["coords"][0]), float(v["coords"][1])) for k, v in n.items()}
    node_z = {k: float(v["coords"][2]) for k, v in n.items()}
    z = lambda k: node_z[k]
    # 급수원 A 에서의 BFS 부모
    adj = {}
    for pid, q in p.items():
        adj.setdefault(q["start"], []).append((q["end"], pid))
        adj.setdefault(q["end"], []).append((q["start"], pid))
    parents, order = {}, ["A"]
    for cur in order:
        for nxt, _pid in adj.get(cur, ()):
            if nxt not in parents and nxt != "A":
                parents[nxt] = cur
                order.append(nxt)
    bores = {pid: (25, "test") for pid in p}
    fits = build_fittings(kfp, node_xy, bores, parents=parents, node_z=node_z,
                          phys={k: len(adj.get(k, ())) for k in ("A", "B", "C", "D", "E")})
    # 꺾이는 넷: B(가로→세로) · B위(세로→가로) · D위(가로→세로) · D(세로→가로).
    # 엘보는 하류 배관에 단다 — 그 네 배관을 찾아 센다. (헤드 스텁의 엘보 1개는 별개)
    tops = {q["end"] for q in p.values() if q["start"] in ("B", "D") and abs(z(q["start"]) - z(q["end"])) > 1e-9}
    # 엘보는 꺾이는 노드의 «하류» 배관에 실린다: B→B'(B) · B'→C(B') · D'→D 세로관(D') · D→E(D).
    # C→D' 는 C 가 직진이라 부속이 없다.
    bend_pipes = [pid for pid, q in p.items()
                  if (q["start"] in ("B", "D") and q["end"] in tops)              # 세로관 둘
                  or (q["start"] in tops and q["end"] == "C")                    # B'→C
                  or (q["start"] == "D" and q["end"] == "E")]                    # D→E
    assert len(bend_pipes) == 4, bend_pipes
    assert all(fits["per_pipe"][pid]["fittings"] == ["elbow"] for pid in bend_pipes), \
        {pid: fits["per_pipe"][pid]["fittings"] for pid in bend_pipes}
    assert fits["unresolved_kind"] == 0


def test_single_arc_is_counted_and_nothing_is_invented():
    out = convert_to_kfp({"kfp": _line_kfp(), "ho": [ARC_OPEN_RIGHT], **KINDS})
    assert out["ok"] is True
    st = out["stats"]
    assert st["n_jog_pairs"] == 0 and st["n_jog_unpaired"] == 1 and st["n_vert_jog"] == 0
    n = out["kfp"]["nodes_meta_runtime"]
    assert all(abs(float(n[k]["coords"][2])) < 1e-9 for k in ("A", "B", "C", "D", "E"))
    assert out["kfp"]["arc_report"]["unpaired_nodes"] == ["B"]


def test_arc_junction_main_node_is_a_tee_not_unresolved_and_top_tee_counted_once():
    # 주배관 A–M–B, M 에서 세로관 0.5 로 T, T 에서 가지 L·R 양쪽(그림 2 의 실제 모양)
    net = {"pipe_data": {"AM": {"start": "A", "end": "M"}, "MB": {"start": "M", "end": "B"},
                         "MT": {"start": "M", "end": "T"}, "TL": {"start": "T", "end": "L"},
                         "TR": {"start": "T", "end": "R"}},
           "arc_junctions": {"M": {"kind": "T", "top": "T"}}}
    xy = {"A": (0, 0), "M": (1, 0), "B": (2, 0), "T": (1, 0), "L": (1, -1), "R": (1, 1)}
    z = {"A": 0, "M": 0, "B": 0, "T": 0.5, "L": 0.5, "R": 0.5}
    bores = {pid: (25, "test") for pid in net["pipe_data"]}
    parents = {"M": "A", "B": "M", "T": "M", "L": "T", "R": "T"}
    fits = build_fittings(net, xy, bores, parents=parents, node_z=z,
                          phys={"A": 1, "M": 4, "B": 1})   # 평면에서는 4방향 교차였다
    assert fits["unresolved_kind"] == 0
    assert fits["per_pipe"]["MB"]["fittings"] == ["tee-run"]     # 주배관 직진
    assert fits["per_pipe"]["MT"]["fittings"] == ["tee"]         # 세로관으로 분류
    tops = fits["per_pipe"]["TL"]["fittings"] + fits["per_pipe"]["TR"]["fittings"]
    assert tops == ["tee"], tops                                  # 꼭대기 티는 한 번
    assert fits["tee_once"] == [{"node": "T", "pipe": "TL", "skipped": "TR"}]
    assert fits["counts"] == {"tee-run": 1, "tee": 2}


def test_four_way_without_arc_stays_unresolved():
    net = {"pipe_data": {"a": {"start": "s", "end": "n"}, "b": {"start": "n", "end": "h"}}}
    result = build_fittings(net, {"s": (-1, 0), "n": (0, 0), "h": (0, 1)},
                            {"b": (25, "test")}, parents={"n": "s"}, phys={"n": 4})
    assert result["unresolved_kind"] == 1
