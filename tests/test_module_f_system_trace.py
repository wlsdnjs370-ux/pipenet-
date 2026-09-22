# -*- coding: utf-8 -*-
"""[오너 2026-09-22] 계통도 칸 — 레이어 세 묶음 · ★추적 · T 접속 · «같은 층» 뒤처리.

■ 오너 지시 (2026-09-22)

  「계통도 쪽에는 기본값은 "건축 레이어"와 배관망(헤드, 배관, 밸브, 부속류 등) 관련
   레이어만 기본값으로 표시되게 하고 나머지는 아예 숨김처리한 뒤 … 좌측하단에 테이블
   형태로 활성화/비활성화로 표시만 할 수 있게」
  「바꾼 뒤 경로추적 경로가 살짝 틀렸어 … 왜 경로가 저렇게 나왔는지 생각해보고, 저
   경로사이에 있는 피팅류도 어떻게 처리되야 할지 … (저런 경우엔 z축 어느정도
   반영해야겠지)」

■ 여기서 지키는 것

  ⑴ 레이어 세 묶음 — 꺼 둔 레이어 · 구역선 · 문자는 숨김, 배관 · 헤드 · 밸브는 배관망.
  ⑵ ★추적은 숨김 · 헤드가 섞인 도면만 «켜진 배관» 으로 좁힌다. 좁힌 망이 억지 연결을
     부르면 그대로 둔다. 섞인 것이 없는 도면은 지금 그대로다.
  ⑶ T 접속 — 한 배관의 끝이 다른 배관 한가운데에 닿으면 잇는다. X 교차는 안 잇는다.
  ⑷ «같은 층» — 합친 배관도 실측 길이 · 알람밸브 꼬리 · 호 우회 0.5 m(엘보 4) · 부속.
  ⑸ «도면» 기준과 결합(부속 행 없음)은 종전과 같다.
"""
from __future__ import annotations

import math

import pytest

P1 = "6-소화-SP-메인(1차)"      # 이름 사전 PIPE · 자동 추적 키워드(소화 · SP)
BR = "6-소화-가지관"            # PIPE
HD = "6-소화-헤드-측벽형"        # HEAD — 자동에 섞이면 «숨김·헤드» 가 섞인 것
ZN = "6-소화-SP-ZONE"           # 이름은 PIPE 지만 구역선 — 숨김
AR = "A-B1"                     # 건축


def L(layer, x1, y1, x2, y2):
    return {"t": "L", "l": layer, "p": [float(x1), float(y1), float(x2), float(y2)]}


def _parsed(names, hidden=()):
    return {"layers": [{"name": n, "visible": n not in hidden} for n in names]}


def _red_path_scene():
    """B1F 의 그 자리 — 주배관 끝이 세로 주배관 **한가운데** 에 닿는다.

    옆 가지관이 주배관 끝 1.26 m 옆에서 올라가 알람밸브 2.2 m 앞까지 간다.
    T 접속을 안 이으면 경로는 가지관으로 돌아간다(97.9 m) — 세로 주배관이
    주배관 끝과 안 이어져 있기 때문이다. 이으면 세로로 곧장 간다(95 m).
    """
    ents = [
        L(P1, 60000, 0, 10000, 0),            # 주배관(펌프 → 왼쪽)
        L(P1, 10000, -20000, 10000, 30000),   # 세로 주배관 — 주배관 끝이 한가운데에
        L(P1, 10000, 30000, 25000, 30000),    # 윗 가로 → 알람밸브
        L(BR, 11200, 400, 11200, 32000),      # 옆 가지관(돌아가는 길)
        L(BR, 11200, 32000, 24000, 32000),
        L(HD, 40000, 5000, 40300, 5000),      # 헤드 기호 조각
        L(HD, 41000, 5000, 41300, 5000),
        {"t": "PL", "l": ZN, "p": [[0, -30000], [70000, -30000], [70000, 40000],
                                   [0, 40000], [0, -30000]]},
        L(AR, -5000, -35000, 75000, -35000),
    ]
    return ents, _parsed([P1, BR, HD, ZN, AR])


PUMP, AV = (60000.0, 0.0), (25000.0, 30000.0)


def _plan_len(riser):
    xy = {n["label"]: (n["x"], n["y"]) for n in riser["nodes"]}
    return sum(math.dist(xy[p["in"]], xy[p["out"]]) for p in riser["pipes"]) / 1000.0


# ── ⑴ 레이어 세 묶음 ──────────────────────────────────────────────────────
def test_layer_groups_follow_the_owner_rule():
    from routes.module_f.sub_trace import layer_group
    g = lambda n, vis=True, c1="OTHER", c2="OTHER": layer_group(n, visible=vis, cat_name=c1, cat_geo=c2)[0]
    assert g("6-소화-헤드-반경", vis=False, c1="HEAD", c2="HEAD") == "etc"   # CAD 꺼 둠
    assert g("6-소화-SP-ZONE", c1="PIPE", c2="PIPE") == "etc"               # 구역선
    assert g("6-소화-TXT", c1="TEXT", c2="TEXT") == "etc"                   # 문자
    assert g("6-소화-가지관", c1="PIPE", c2="PIPE") == "net"
    assert g("6-소화-헤드-측벽형", c1="HEAD", c2="HEAD") == "net"
    assert g("6-소화-밸브") == "net"                                         # 밸브 이름
    assert g("6-소화-FLX") == "net"
    assert g("A-B1", c1="ARCH", c2="ARCH") == "arch"
    assert g("FF") == "etc"                                                  # 분류 안 됨


def test_layer_rows_count_and_sort_by_group():
    from routes.module_f.sub_trace import layer_rows
    ents, parsed = _red_path_scene()
    rows = layer_rows(ents, parsed)
    by = {r["layer"]: r for r in rows}
    assert by[P1]["group"] == "net" and by[BR]["group"] == "net" and by[HD]["group"] == "net"
    assert by[ZN]["group"] == "etc" and by[AR]["group"] == "arch"
    assert by[P1]["n"] == 3 and by[P1]["auto"]
    order = [r["group"] for r in rows]
    assert order == sorted(order, key=("net", "arch", "etc").index)


# ── ⑶ T 접속 ─────────────────────────────────────────────────────────────
def test_t_junction_splits_only_where_an_end_lands_mid_pipe():
    from routes.module_f.sub_trace import split_t_junctions
    ents = [L(P1, 0, 0, 10000, 0),           # 가로
            L(P1, 5000, 0, 5000, 3000),      # 끝이 가로 한가운데 → T
            L(P1, 7000, -2000, 7000, 2000),  # 가로를 엇갈려 지나감 → X, 안 자른다
            L(P1, 8000, 31, 8000, 2000),     # 끝이 31 mm 떨어짐(눈금) → 안 자른다
            {"t": "T", "l": P1, "p": [1, 1], "v": "150A"}]
    out, st = split_t_junctions(ents, {P1}, tol=2.5)
    assert st == {"split_entities": 1, "cuts": 1}
    pieces = sorted(tuple(e["p"]) for e in out if e["t"] == "L" and e["p"][1] == 0 and e["p"][3] == 0)
    assert pieces == [(0.0, 0.0, 5000.0, 0.0), (5000.0, 0.0, 10000.0, 0.0)]
    assert ents[4] in out and ents[2] in out and ents[3] in out   # 나머지는 같은 객체 그대로


# ── ⑵ ★추적 레이어 ────────────────────────────────────────────────────────
def test_mixed_drawing_traces_only_the_visible_pipe_layers():
    from routes.module_f.sub_trace import layer_rows, plan_trace
    ents, parsed = _red_path_scene()
    tr = plan_trace(ents, layer_rows(ents, parsed))
    assert tr["mode"] == "pipe"
    assert tr["layers"] == sorted([P1, BR])
    assert set(tr["junk"]) == {HD, ZN}
    assert tr["forced"] == 0 and tr["split"]["cuts"] >= 1
    assert tr["graph"][2]["layer_filter_used"] == sorted([P1, BR])


def test_clean_drawing_keeps_todays_tracing():
    """숨김·헤드가 안 섞인 도면(대명동 계통도 꼴)은 ★ 를 안 켠다 — 지금 그대로."""
    from routes.module_f.sub_trace import layer_rows, plan_trace
    ents = [L("HSP", 0, 0, 0, 30000), L("LSP", 500, 0, 500, 30000),
            L("감압밸브", 0, 10000, 500, 10000), L("FF", -9000, 0, 9000, 0)]
    tr = plan_trace(ents, layer_rows(ents, _parsed(["HSP", "LSP", "감압밸브", "FF"])))
    assert tr["mode"] == "auto" and tr["layers"] is None and tr["junk"] == []
    assert "entities" not in tr and "graph" not in tr


def test_fragmented_pipe_layers_fall_back_to_todays_tracing():
    """배관이 다른 레이어에도 그려진 도면(MF-004 꼴) — 좁히면 억지 연결이 생기니 그대로."""
    from routes.module_f.sub_trace import layer_rows, plan_trace
    ents = [L(P1, 0, 0, 10000, 0), L(P1, 60000, 0, 70000, 0),   # 50 m 끊김
            L("fire", 10000, 0, 60000, 0),                        # 끊긴 곳은 fire 에
            L(HD, 1000, 500, 1300, 500)]
    tr = plan_trace(ents, layer_rows(ents, _parsed([P1, "fire", HD])))
    assert tr["mode"] == "auto" and tr["forced"] >= 1 and tr["layers"] is None


def test_red_path_goes_straight_up_the_vertical_main():
    """오너가 빨간 선으로 짚은 길 — T 접속을 이으면 세로 주배관으로 곧장 간다."""
    from routes.module_f.sub_trace import layer_rows, plan_trace
    from routes.module_f.subdrawing import extract_system
    ents, parsed = _red_path_scene()
    tr = plan_trace(ents, layer_rows(ents, parsed))
    lay = set(tr["layers"])
    before = extract_system(ents, PUMP, AV, layer_filter=lay)
    after = extract_system(tr["entities"], PUMP, AV, layer_filter=lay)
    assert _plan_len(before) == pytest.approx(97.9, abs=0.3)     # 가지관으로 돌아감
    assert _plan_len(after) == pytest.approx(95.0, abs=0.01)     # 50 + 30 + 15
    assert any(n["x"] == 10000 and n["y"] == 0 for n in after["nodes"])


def test_drawing_mode_output_is_unchanged():
    """«도면» 기준은 뒤처리를 안 탄다 — 부속 행 · same_level 칸이 없다."""
    from routes.module_f.subdrawing import extract_system
    ents, _ = _red_path_scene()
    r = extract_system(ents, PUMP, AV, layer_filter={P1, BR})
    assert "fittings" not in r and "same_level" not in r
    assert r["elevation_mode"] == "drawing"


# ── ⑷ «같은 층» ──────────────────────────────────────────────────────────
def _same_level(ents, pump, av, layers):
    from routes.module_f.subdrawing import extract_system
    return extract_system(ents, pump, av, layer_filter=layers, elevation_mode="same_level")


def test_same_level_finds_measured_length_of_merged_runs():
    """한 직선 위 조각 셋 → 추출이 배관 하나로 합친다. 종전: «실측 길이를 찾을 수 없습니다»."""
    ents = [L(P1, 0, 0, 3000, 0), L(P1, 3000, 0, 7000, 0), L(P1, 7000, 0, 10000, 0),
            L(P1, 10000, 0, 10000, 4000)]
    r = _same_level(ents, (0, 0), (10000, 4000), {P1})
    assert [round(p["length"], 3) for p in r["pipes"]] == [10.0, 4.0]
    assert all(n["elevation"] == 0 for n in r["nodes"])


def test_same_level_still_refuses_a_run_with_a_real_gap():
    """틈이 있는 구간은 지어내지 않는다 — 조각이 덮지 못하면 종전처럼 멈춘다."""
    from routes.module_f.sub_same_level import measured_lengths
    result = {"nodes": [{"label": "1", "x": 0, "y": 0}, {"label": "10", "x": 10000, "y": 0}],
              "pipes": [{"label": "r1", "in": "1", "out": "10", "length": 10.0}]}
    edge_len = {((0.0, 0.0), (4000.0, 0.0)): 4000.0, ((6000.0, 0.0), (10000.0, 0.0)): 4000.0}
    m = measured_lengths(result, edge_len, {"snap_eps_mm": 50.0})
    assert ((0, 0), (10000, 0)) not in {(a, b) for a, b in m}   # 2 m 틈 — 안 채운다


def test_same_level_red_path_counts_the_tee_and_elbow():
    """주배관 → 세로 주배관: 티(도면 3방향)에서 꺾이니 분류티. 꼭대기는 엘보."""
    from routes.module_f.sub_trace import layer_rows, plan_trace
    ents, parsed = _red_path_scene()
    tr = plan_trace(ents, layer_rows(ents, parsed))
    r = _same_level(tr["entities"], PUMP, AV, set(tr["layers"]))
    assert r["total_pipe_length_m"] == pytest.approx(95.0, abs=0.01)
    assert r["same_level"]["fittings"] == {"tee": 1, "elbow": 1}
    kinds = {(f["pipe"], f["type"]) for f in r["fittings"]}
    up = next(p for p in r["pipes"] if p["in"] != "1" and p["out"] != "10")
    assert (up["label"], "tee") in kinds                    # 세로로 올라가는 배관에 분류티
    assert all(f["type"] in ("tee", "elbow", "elbow-45") for f in r["fittings"])


def test_arc_pair_on_the_path_is_a_half_metre_hop_with_four_elbows():
    """덱 「실제길이 A」 — 마주 보는 원호 사이를 0.5 m 올린다. 엘보 4 (+ 모서리 1)."""
    ents = [L(P1, 0, 0, 40000, 0), L(P1, 40000, 0, 40000, 10000),
            {"t": "A", "l": P1, "c": [20000, 0], "r": 175, "a": [90, 270]},    # ( — 동쪽이 열림
            {"t": "A", "l": P1, "c": [26000, 0], "r": 175, "a": [270, 90]}]    # ) — 서쪽이 열림
    r = _same_level(ents, (0, 0), (40000, 10000), {P1})
    rep = r["same_level"]
    assert rep["hops"]["pairs"] == 1 and rep["hops"]["unpaired"] == 0
    z = {(n["x"], n["y"], n["elevation"]) for n in r["nodes"]}
    assert {(20000, 0, 0.0), (20000, 0, 0.5), (26000, 0, 0.5), (26000, 0, 0.0)} <= z
    lens = sorted(round(p["length"], 3) for p in r["pipes"])
    assert lens == [0.5, 0.5, 6.0, 10.0, 14.0, 20.0]
    assert r["total_pipe_length_m"] == pytest.approx(51.0)
    assert rep["fittings"] == {"elbow": 5}                  # 우회 4 + 모서리 1
    verts = [p for p in r["pipes"] if p.get("hop")]
    assert sorted(p["elev"] for p in verts) == [-0.5, 0.5]


def test_lonely_arc_is_counted_not_invented():
    ents = [L(P1, 0, 0, 40000, 0), L(P1, 40000, 0, 40000, 10000),
            {"t": "A", "l": P1, "c": [20000, 0], "r": 175, "a": [90, 270]}]
    r = _same_level(ents, (0, 0), (40000, 10000), {P1})
    assert r["same_level"]["hops"] == {"pairs": 0, "unpaired": 1, "rise_m": 0.5}
    assert all(n["elevation"] == 0 for n in r["nodes"])


def test_av_tail_wiggle_through_a_double_line_is_straightened():
    """알람밸브 조립 상세(입면 이중선)를 오르내린 꼬리 → 곧장. 진짜 입상관은 남긴다."""
    ents = [L(P1, 0, 0, 10000, 0), L(P1, 10000, 0, 10000, 1000),        # 주배관 · 입상
            L(P1, 10000, 1000, 10100, 1000), L(P1, 10100, 1000, 10100, 1800),
            L(P1, 10100, 1800, 10250, 1800), L(P1, 10250, 1800, 10250, 1000),
            L(P1, 10250, 1000, 10400, 1000)]
    r = _same_level(ents, (0, 0), (10400, 1000), {P1})
    tail = r["same_level"]["av_tail"]
    assert tail and tail["removed_pipes"] == 4 and tail["straight_m"] == pytest.approx(0.3)
    xy = [(n["x"], n["y"]) for n in r["nodes"]]
    assert (10000, 1000) in xy and (10100, 1800) not in xy            # 입상관은 그대로
    assert r["total_pipe_length_m"] == pytest.approx(10.0 + 1.0 + 0.1 + 0.3)


def test_a_straight_or_single_bend_approach_is_left_alone():
    ents = [L(P1, 0, 0, 10000, 0), L(P1, 10000, 0, 10000, 500)]
    r = _same_level(ents, (0, 0), (10000, 500), {P1})
    assert r["same_level"]["av_tail"] is None


# ── ⑸ 결합 — 같은 층 부속 행을 싣는다 · 도면 기준은 종전 ─────────────────────
def test_merge_carries_same_level_fittings_only():
    from test_module_f_system_layout import heads   # 같은 꼴의 평면도 표
    from routes.module_f.merge import merge_network
    from routes.module_f.subdrawing import extract_system
    ents = [L("PIPE", 0, 0, 3000, 0), L("PIPE", 3000, 0, 3000, 4000)]
    same = extract_system(ents, (0, 0), (3000, 4000), layer_filter={"PIPE"},
                          elevation_mode="same_level")
    got = merge_network(heads(), riser=same, mode="hsp_pump")
    sys_rows = [f for f in got["combined"].fittings if str(f.get("pipe", "")).startswith("r")]
    assert [f["type"] for f in same["fittings"]] == ["elbow"]
    assert any(f["type"] == "elbow" for f in sys_rows)
    assert any("계통도 부속" in s for s in got["steps"])
    drawing = extract_system(ents, (0, 0), (3000, 4000), layer_filter={"PIPE"})
    got2 = merge_network(heads(), riser=drawing, mode="hsp_pump")
    assert not any("계통도 부속" in s for s in got2["steps"])
    assert not any(f["type"] == "elbow" and str(f["pipe"]).startswith("r")
                   for f in got2["combined"].fittings)
