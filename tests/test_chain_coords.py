# -*- coding: utf-8 -*-
"""[회랑 사슬좌표 §4-3] 좌표를 «재지 않고 만든다» — 식으로 못박는다.

지시서 `ModuleF_회랑_사슬좌표_지시서.md` · 그림 7·8.

■ 규칙

    p(자식) = p(부모) + L_e · u_e

    L_e  평면 배관 = 양 끝 board 노드의 유클리드 거리 · 세로 토막 = 입력값
    u_e  평면 배관 = board 방향을 8 방향(0·45·…·315°)으로 · 세로 = ±z

■ 판정 (§4-1)

    C1 사슬 닫힘 · C2 길이 = 좌표 거리 · C3 1:1 · C4 방향 ∈ 8 방향

■ 왜 각도를 3°·87°·−2°·40° 로 두나

  도면은 딱 맞아떨어지지 않는다(그림 7 ①). 8 방향으로 돌리면 끝점이 밀리고,
  그 밀림이 **쌓인다**(그림 8 ③). 직교만인 망은 밀림이 0 이라 그것으로는
  규칙이 도는지 알 수 없다 — 그래서 일부러 «어중간한» 각을 넣는다.

■ 상하향식

  오너 지시(2026-09-14): **로직만 두고 실제 검증은 상향식·하향식 두 가지로.**
  그래서 상하향식은 여기 합성 망으로만 확인한다(§4-2 갈음) — 실도면 회귀는
  `scripts/_probe_plan_preserve.py` 가 대명동(하향식)·B1F(상향식)로 돈다.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.design.chain import (chain_coords,     # noqa: E402
                                              snap_dir8)


# ═══════════════════════ 합성 망 만들기 (그림 7 의 런 + 티 + 헤드)
def _pt(x, y):
    return (float(x), float(y))


def _net(angles_len, *, extra_heads=(), declared=()):
    """한 줄짜리 런 — (각도°, 길이mm) 목록으로 board 와 kfp 를 함께 만든다.

    kfp 좌표는 «아무 값» 이어도 된다(사슬이 다시 만든다). 일부러 board 와
    다른 값을 넣어 — 사슬이 **정말** 다시 만드는지 보이게 한다.
    """
    board = [_pt(0, 0)]
    x, y = 0.0, 0.0
    for a, L in angles_len:
        x += L * math.cos(math.radians(a))
        y += L * math.sin(math.radians(a))
        board.append(_pt(x, y))
    nodes, pipes, node_ref, edge_ref = {}, {}, {}, {}
    for i, p in enumerate(board):
        nid = f"N{i}"
        nodes[nid] = {"id": nid, "coords": [p[0] / 1000.0 + 1.0 + 0.011,
                                            p[1] / 1000.0 + 1.0 - 0.007, 0.0],
                      "type": "Node", "type_id": "base"}
        node_ref[nid] = i
    # 뿌리 = 접속점(알람밸브). `require_anchor` 는 `type_id='pump'` 로 찾는다 —
    #   없으면 «물이 어디로 들어오는지 모른다» 며 멈춘다(지어내지 않는다).
    nodes["N0"]["type_id"] = "pump"
    nodes["N0"]["io_node"] = "Input"
    for i in range(len(board) - 1):
        pid = f"P{i + 1}"
        pipes[pid] = {"start": f"N{i}", "end": f"N{i + 1}",
                      "length_m": 0.0, "C": 120, "roughness_mm": 0.0}
        edge_ref[pid] = (i, i + 1)
    for rec in extra_heads:
        nodes.update(rec["nodes"])
        pipes.update(rec["pipes"])
    kfp = {"nodes_meta_runtime": nodes, "pipe_data": pipes}
    return kfp, board, node_ref, edge_ref


def _run(kfp, board, node_ref, edge_ref, **kw):
    return chain_coords(kfp, edge_ref=edge_ref, node_ref=node_ref,
                        board_pts=board, origin_mm=(0.0, 0.0), **kw)


def _xy(kfp, nid):
    c = kfp["nodes_meta_runtime"][nid]["coords"]
    return (float(c[0]), float(c[1]))


def _len(kfp, pid):
    return float(kfp["pipe_data"][pid]["length_m"])


# 그림 7 의 네 토막 — 도면이 딱 맞아떨어지지 않는 모양
RUN = [(3.0, 3000.0), (87.0, 2200.0), (-2.0, 2600.0), (40.0, 1600.0)]


# ═══════════════════════ C2 — 길이는 도면 값 그대로
def test_길이가_도면_값_그대로다():
    kfp, board, nref, eref = _net(RUN)
    r = _run(kfp, board, nref, eref)
    assert r["ok"], r
    for i, (_a, L) in enumerate(RUN):
        got = _len(kfp, f"P{i + 1}") * 1000.0
        assert abs(got - L) <= 1.0, (i, got, L)


def test_길이가_좌표_거리와_같다():
    """★C2 — 표의 길이를 좌표에서 다시 재도 참값이어야 한다(그림 8 아래)."""
    kfp, board, nref, eref = _net(RUN)
    _run(kfp, board, nref, eref)
    for pid, pr in kfp["pipe_data"].items():
        a, b = _xy(kfp, pr["start"]), _xy(kfp, pr["end"])
        d = math.dist(a, b) * 1000.0
        assert abs(d - _len(kfp, pid) * 1000.0) <= 1.0, pid


# ═══════════════════════ C4 — 방향은 8 방향뿐
def test_방향이_8방향이다():
    kfp, board, nref, eref = _net(RUN)
    r = _run(kfp, board, nref, eref)
    want = [0.0, 90.0, 0.0, 45.0]          # 3°·87°·−2°·40° 가 가는 자리
    for i, w in enumerate(want):
        a, b = _xy(kfp, f"N{i}"), _xy(kfp, f"N{i + 1}")
        ang = math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])) % 360.0
        assert abs(((ang - w + 180.0) % 360.0) - 180.0) <= 1e-6, (i, ang, w)
    assert r["dev_max"] <= 22.5 and r["dev_over_22_5"] == 0


# ═══════════════════════ C1 — 사슬이 닫힌다 · 끝점 이탈은 해석값과 같다
def test_사슬이_닫히고_끝점이_해석값만큼_밀린다():
    kfp, board, nref, eref = _net(RUN)
    _run(kfp, board, nref, eref)
    # 사슬을 손으로 이어 본다 — 8 방향 · 도면 길이
    x = y = 0.0
    for a, L in RUN:
        ux, uy, _dev = snap_dir8(math.cos(math.radians(a)),
                                 math.sin(math.radians(a)))
        x += L * ux
        y += L * uy
    got = _xy(kfp, f"N{len(RUN)}")
    p0 = _xy(kfp, "N0")
    assert abs((got[0] - p0[0]) * 1000.0 - x) <= 1e-6
    assert abs((got[1] - p0[1]) * 1000.0 - y) <= 1e-6
    # 도면 끝점과의 이탈 — 0 이 아니고(각이 어중간하니) 상한 안이다
    slip = math.dist((x, y), (board[-1][0], board[-1][1]))
    bound = sum(L * math.sin(math.radians(snap_dir8(
        math.cos(math.radians(a)), math.sin(math.radians(a)))[2]))
        for a, L in RUN)
    assert slip > 1.0, slip
    assert slip <= bound + 1e-6, (slip, bound)


def test_직교만인_망은_이탈이_0이다():
    """★그림 8 ③ — 직교·45° 배관은 회전각 0 이라 밀림이 없다."""
    ortho = [(0.0, 3000.0), (90.0, 2200.0), (180.0, 1000.0), (45.0, 1400.0)]
    kfp, board, nref, eref = _net(ortho)
    r = _run(kfp, board, nref, eref)
    assert r["dev_max"] <= 1e-9
    end = _xy(kfp, f"N{len(ortho)}")
    p0 = _xy(kfp, "N0")
    assert math.dist(((end[0] - p0[0]) * 1000.0, (end[1] - p0[1]) * 1000.0),
                     board[-1]) <= 1e-6


# ═══════════════════════ 선언 길이(신축배관 접기)는 건드리지 않는다
def test_선언_길이_배관은_그대로_둔다():
    kfp, board, nref, eref = _net(RUN)
    kfp["pipe_data"]["P2"]["length_m"] = 9.999
    _run(kfp, board, nref, eref, declared_pipes=("P2",))
    assert _len(kfp, "P2") == 9.999, "선언값을 덮었다"
    assert abs(_len(kfp, "P1") * 1000.0 - RUN[0][1]) <= 1.0


# ═══════════════════════ 세로 토막 — 입력값 그대로 · ±z
def _head_leg(base_nid, hid, dz, x_m, y_m, z0=0.0, L=0.3):
    """base 에서 dz 만큼 오르내리는 세로 토막 하나 (헤드)."""
    return {"nodes": {hid: {"id": hid, "coords": [x_m, y_m, z0 + dz],
                            "type": "Head", "type_id": "head"}},
            "pipes": {f"V{hid}": {"start": base_nid, "end": hid,
                                  "length_m": L, "C": 120}}}


def test_세로_토막은_입력값_그대로다():
    kfp, board, nref, eref = _net(RUN)
    b = _xy(kfp, "N2")
    kfp["nodes_meta_runtime"]["H1"] = {
        "id": "H1", "coords": [b[0], b[1], -0.3],
        "type": "Head", "type_id": "head"}
    kfp["pipe_data"]["V1"] = {"start": "N2", "end": "H1", "length_m": 0.3,
                              "C": 120}
    _run(kfp, board, nref, eref)
    a, h = _xy(kfp, "N2"), _xy(kfp, "H1")
    assert math.dist(a, h) <= 1e-9, "세로가 옆으로 갔다"
    za = kfp["nodes_meta_runtime"]["N2"]["coords"][2]
    zh = kfp["nodes_meta_runtime"]["H1"]["coords"][2]
    assert abs((zh - za) + 0.3) <= 1e-9, (za, zh)
    assert abs(_len(kfp, "V1") - 0.3) <= 1e-9


# ═══════════════════════ 상하향식 ③ — **로직만** (오너: 실검증은 상·하향만)
def test_상하향식_옆_토막은_8방향으로_간다():
    """상하향식 ③ 은 가지 방향으로 옆으로 뻗는다(그림 2′ 오른쪽).

    board 쌍이 없는 «엔진이 만든 가로 토막» 이라, 사슬은 길이를 그대로 두고
    방향만 8 방향으로 맞춘다. 실도면 검증은 상향식·하향식으로만 한다
    (오너 2026-09-14) — 여기서 로직이 도는 것만 못박는다.
    """
    kfp, board, nref, eref = _net(RUN)
    b = _xy(kfp, "N2")
    # ③ = N2 에서 37° 로 0.3 m (board 쌍 없음) → 45° 로 가야 한다
    kx = b[0] + 0.3 * math.cos(math.radians(37.0))
    ky = b[1] + 0.3 * math.sin(math.radians(37.0))
    kfp["nodes_meta_runtime"]["K1"] = {"id": "K1", "coords": [kx, ky, 0.0],
                                       "type": "Node", "type_id": "base"}
    kfp["pipe_data"]["E1"] = {"start": "N2", "end": "K1", "length_m": 0.3,
                              "C": 120}
    r = _run(kfp, board, nref, eref)
    assert r["kinds"]["engine"] == 1, r["kinds"]
    a, kxy = _xy(kfp, "N2"), _xy(kfp, "K1")
    ang = math.degrees(math.atan2(kxy[1] - a[1], kxy[0] - a[0])) % 360.0
    assert abs(ang - 45.0) <= 1e-6, ang
    assert abs(math.dist(a, kxy) - 0.3) <= 1e-9, "③ 길이가 바뀌었다"


# ═══════════════════════ §2 멈춤 — 고리가 있으면 닿지 못한 노드를 낸다
def test_고리가_있으면_말한다():
    """회랑은 나무여야 한다. 임의로 끊어 나무로 만들지 않는다(§5)."""
    kfp, board, nref, eref = _net(RUN)
    # 뿌리에서 떨어진 조각을 하나 둔다 — BFS 가 못 닿는다
    kfp["nodes_meta_runtime"]["X1"] = {"id": "X1", "coords": [9.0, 9.0, 0.0],
                                       "type": "Node", "type_id": "base"}
    kfp["nodes_meta_runtime"]["X2"] = {"id": "X2", "coords": [9.5, 9.0, 0.0],
                                       "type": "Node", "type_id": "base"}
    kfp["pipe_data"]["PX"] = {"start": "X1", "end": "X2", "length_m": 0.5}
    r = _run(kfp, board, nref, eref)
    assert set(r["unreached"]) == {"X1", "X2"}, r["unreached"]


# ═══════════════════════ §5 금지 — 손대면 안 되는 파일
def test_금지_파일을_안_바꿨다():
    """전체망이 그 코드를 쓴다 — 사슬은 회랑에서만 돈다(§1-2 · §5)."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import")
    for rel in ("convert/planar.py", "convert/engine.py",
                "design/sdf_post.py", "design/worst.py"):
        t = (src / rel).read_text(encoding="utf-8")
        assert "chain_coords" not in t, f"{rel} 이 사슬을 부른다"
    api = (_ROOT / "routes" / "module_f" / "api_edit.py").read_text(
        encoding="utf-8")
    assert "chain_coords" not in api


def test_사슬은_expand_worst_에서만_불린다():
    t = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
         / "design" / "restrict.py").read_text(encoding="utf-8")
    assert t.count("chain_coords(") == 1
    i = t.index("def expand_worst(")
    assert t.index("chain_coords(") > i, "회랑 밖에서 부른다"
