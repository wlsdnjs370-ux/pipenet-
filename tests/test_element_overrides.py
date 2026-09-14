# -*- coding: utf-8 -*-
"""[요소속성 수정카드 §5-2] 사람이 덮은 값이 «가리킨 그 자리» 에만 닿는가.

지시서 `ModuleF_요소속성_수정카드_지시서.md` · 그림 15 의 규칙 여섯을 시험으로
못박는다. 여기서 지키는 것:

  ① **안정 키 왕복** — 담고 꺼낸 키가 같은 것을 가리킨다.
  ② **번역이지 판정이 아니다** — 키 → 이번 이름은 빌드마다 다시 매겨진다.
  ③ **못 옮긴 것은 말한다**(규칙 4) — 조용히 안 버린다.
  ④ **원값을 채운다**(규칙 5) — 「무엇에서 무엇으로」가 남는다.
  ⑤ **그 값만 바뀐다** — 옆 배관은 건드리지 않는다.
  ⑥ **길이는 좌표를 움직인다** — 회랑은 사슬이 길이로 좌표를 만든다.
  ⑦ **저장소를 둘 두지 않는다**(규칙 6 · §6) — 옛 관경 목록을 흡수한다.
  ⑧ **통합의 계통도·기계실**도 같은 저장소로 고친다(§4 · 오너 승인).

★실도면 시험(대명동 K=30)은 무거우므로 한 벌만 돈다 — 나머지는 최소 표본으로
  규약만 잰다. 규약이 깨지는 것은 대개 표본에서 먼저 드러난다.
"""
from __future__ import annotations

import json
import math
import os
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "core"),
           os.path.join(_ROOT, "cad_project_editor_g"),
           os.path.join(_ROOT, "scripts"),
           os.path.join(_ROOT, "tests")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from routes.module_f import overrides as ov  # noqa: E402


# ═══════════════════════════════════════════ ① 안정 키 왕복
def test_안정_키는_담고_꺼내도_같은_자리를_가리킨다():
    """키는 JSON 으로 파일에 남는다 — 왕복에서 갈래가 바뀌면 남의 자리를 덮는다."""
    cases = [
        ov.key_pipe(2472, 1243),          # 정렬은 키가 한다
        ov.key_node(77),
        ov.key_head(12),
        ov.key_vert("head", 12, 2),
        ov.key_merge("sys", "r3"),
        ov.key_merge("mr", "m1"),
    ]
    for k in cases:
        back = ov.key_from_json(json.loads(json.dumps(ov.key_to_json(k))))
        assert back == k, f"{k} 가 왕복에서 {back} 로 바뀌었다"


def test_배관_키는_노드쌍을_정렬해_담는다():
    """(a,b) 와 (b,a) 는 같은 배관이다 — 두 줄로 담기면 뒤엣것이 이긴다."""
    assert ov.key_pipe(9, 3) == ov.key_pipe(3, 9) == ("pipe", 3, 9)


def test_모르는_키는_조용히_통과하지_않는다():
    for bad in (None, [], ["nope", 1], ["pipe"], "pipe"):
        assert ov.key_from_json(bad) is None


# ═══════════════════════════════════════════ ② 값 검증 (기준 8)
def test_고칠_수_없는_속성은_거절한다():
    val, why = ov.validate("pipe", "주인없는칸", 1)
    assert val is None and "고칠 수 없는" in why


def test_범위_밖_값은_거절한다():
    for kind, field, bad in (("pipe", "length", -1),
                             ("pipe", "dia", 0),
                             ("head", "k_factor_si", 0),
                             ("pipe", "roughness_mm", -0.1)):
        val, why = ov.validate(kind, field, bad)
        assert val is None, f"{kind}.{field} 가 {bad} 를 받았다"
        assert why


def test_수가_아닌_값은_거절한다():
    for bad in ("", "칠점일", float("nan"), float("inf")):
        val, why = ov.validate("pipe", "length", bad)
        assert val is None and why


def test_길이만_사유가_필수다():
    """전부 필수로 하면 사람이 아무 글자나 적게 되고 사유 칸이 뜻을 잃는다."""
    assert ("pipe", "length") in ov.REASON_REQUIRED
    assert ("vert", "length") in ov.REASON_REQUIRED
    assert ("pipe", "dia") not in ov.REASON_REQUIRED
    assert ("node", "elevation") not in ov.REASON_REQUIRED


# ═══════════════════════════════════════════ ③ 저장소
def test_같은_자리_같은_속성은_하나뿐이다():
    k = ov.key_pipe(1, 2)
    rows = ov.put([], k, "pipe", "length", 3.0, reason="ㄱ", at="t1")
    rows = ov.put(rows, k, "pipe", "length", 5.0, reason="ㄴ", at="t2")
    assert len(rows) == 1
    assert rows[0]["new"] == 5.0 and rows[0]["reason"] == "ㄴ"


def test_다른_속성은_나란히_산다():
    k = ov.key_pipe(1, 2)
    rows = ov.put([], k, "pipe", "length", 3.0)
    rows = ov.put(rows, k, "pipe", "dia", 65)
    assert {r["field"] for r in rows} == {"length", "dia"}
    rows = ov.drop(rows, k, "length")
    assert [r["field"] for r in rows] == ["dia"]


def test_파일로_왕복한다(tmp_path, monkeypatch):
    """서버를 껐다 켜도 값이 남는다 — 원값·사유·시각까지."""
    monkeypatch.setattr(ov, "path_for",
                        lambda k: os.path.join(str(tmp_path),
                                               f"{k}_수리계산수정.json"))
    k = ov.key_pipe(11, 22)
    rows = ov.put([], k, "pipe", "length", 2.5, old=7.5,
                  reason="현장 실측", at="2026-09-14 16:00")
    path = ov.write_file("시험도면", rows)
    assert os.path.exists(path)
    back = ov.read_file("시험도면")
    assert len(back) == 1
    r = back[0]
    assert ov.key_from_json(r["key"]) == k
    assert (r["old"], r["new"]) == (7.5, 2.5)
    assert r["reason"] == "현장 실측" and r["at"] == "2026-09-14 16:00"


def test_파일_이름은_도면_이름을_따른다(monkeypatch):
    import services.cad_import.pipeline.handoff as _h
    monkeypatch.setattr(_h, "default_edits_dir", lambda: "/tmp/그어딘가")
    p = ov.path_for("어떤 도면")
    assert os.path.basename(p) == "어떤 도면_수리계산수정.json"


# ═══════════════════════════════════════════ ④ 번역 · 못 옮긴 것
class _FakeBoard:
    def __init__(self, pts=(), disks=(), hnodes=()):
        self.pts = list(pts)
        self.disks = [list(d) for d in disks]
        self.hnodes = [set(h) for h in hnodes]


def _fake_got(*, pipes, nodes, edge_ref, node_ref, origin=(1000.0, 1000.0)):
    return {
        "kfp": {"pipe_data": pipes, "nodes_meta_runtime": nodes},
        "edge_ref": edge_ref, "node_ref": node_ref, "origin_mm": origin,
    }


def _tiny():
    """배관 둘 · 절점 셋 · 헤드 하나 — 번역이 도는 최소 망."""
    nodes = {
        # ★사슬을 다시 걸려면 **접속점(급수원)** 이 있어야 한다 — 물이 어디서
        #   들어오는지 모르면 어느 쪽을 붙박아 둘지 정할 수 없다. 폴백을 두지
        #   않는 것이 이 저장소의 규약이라(`require_anchor`), 표본에도 둔다.
        "N1": {"coords": [0.0, 0.0, 0.0], "type_id": "pump"},
        "N2": {"coords": [1.0, 0.0, 0.0], "type_id": "base"},
        "N3": {"coords": [2.0, 0.0, 0.0], "type_id": "head",
               "k_factor_si": 80.0},
    }
    # ★`C`·거칠기·관종을 실제로 들려 둔다 — 진짜 망에는 있다(`emit_sdf` 가
    #   `p["c"]` 를 반드시 읽으므로 없으면 산출이 KeyError 로 죽는다).
    #   빈 표본으로 재면 「원값 None」이 정상처럼 보여 결함을 놓친다.
    pipes = {
        "P1": {"start": "N1", "end": "N2", "length_m": 1.0, "C": 120,
               "roughness_mm": 0.1, "equivalent_length": 0.0, "type": "Sch40"},
        "P2": {"start": "N2", "end": "N3", "length_m": 1.0, "C": 120,
               "roughness_mm": 0.1, "equivalent_length": 0.0, "type": "Sch40"},
    }
    got = _fake_got(pipes=pipes, nodes=nodes,
                    edge_ref={"P1": [10, 11], "P2": [11, 12]},
                    node_ref={"N1": 10, "N2": 11, "N3": 12})
    board = _FakeBoard(pts=[(0, 0)] * 13, disks=[(0, 0, 50)],
                       hnodes=[{12}])
    return got, board


def test_키에서_이번_이름으로_번역한다():
    got, board = _tiny()
    idx = ov.build_index(got, board)
    assert idx["pipe"][ov.key_pipe(10, 11)] == "P1"
    assert idx["node"][ov.key_node(11)] == "N2"
    assert idx["head"][ov.key_head(0)] == "N3", "헤드를 원 번호로 못 되짚었다"


def test_헤드는_좌표가_아니라_이음으로_되짚는다():
    """회랑 좌표는 사슬로 «만든» 것이라 도면 mm 와 어긋난다.

    실측(대명동 K=30): 헤드가 가장 가까운 원에서 최소 199 · 중앙 913 ·
    최대 1,998 mm 떨어져 있었다. 좌표에 자를 대면 헤드 키가 하나도 안 생긴다.
    여기서는 원을 **10 m 밖**에 두고도 이음으로 닿는지를 본다.
    """
    got, board = _tiny()
    board.disks = [(10_000.0, 10_000.0, 50.0)]      # 좌표로는 절대 못 찾는다
    idx = ov.build_index(got, board)
    assert idx["head"][ov.key_head(0)] == "N3"


def _tiny_vert():
    """헤드에 세로 토막 둘을 매단 망 — 상하향식이면 한 헤드에 ①②③④ 가 걸린다."""
    got, board = _tiny()
    nd = got["kfp"]["nodes_meta_runtime"]
    pr = got["kfp"]["pipe_data"]
    nd["N4"] = {"coords": [2.0, 0.0, -0.3], "type_id": "base"}
    nd["N5"] = {"coords": [2.0, 0.0, -0.9], "type_id": "base"}
    pr["P3"] = {"start": "N3", "end": "N4", "length_m": 0.3}
    pr["P4"] = {"start": "N4", "end": "N5", "length_m": 0.6}
    return got, board


def test_세로_토막_키가_뿌리_헤드를_따라간다():
    """토막 이름(P3·P4)은 재계산마다 바뀐다 — 뿌리 원 번호와 z 순서로 담는다."""
    got, board = _tiny_vert()
    idx = ov.build_index(got, board)
    # 역할은 (아래 z, 위 z) 오름차순 — 가장 아래 토막이 1번이다.
    assert idx["vert"][ov.key_vert("head", 0, 1)] == "P4"   # (−0.9, −0.3)
    assert idx["vert"][ov.key_vert("head", 0, 2)] == "P3"   # (−0.3,  0.0)
    assert idx["roots"]["P3"] == ov.key_vert("head", 0, 2)
    assert idx["roots"]["P4"] == ov.key_vert("head", 0, 1),         "두 번째 토막이 뿌리를 못 찾아 키 없이 떨어졌다"


def test_세로_토막_길이도_좌표를_움직인다():
    got, board = _tiny_vert()
    before = {k: list(v["coords"])
              for k, v in got["kfp"]["nodes_meta_runtime"].items()}
    rows = ov.put([], ov.key_vert("head", 0, 2), "vert", "length", 1.2,
                  reason="시험", at="t")
    n, missed, _rep = ov.apply_to_kfp(got, board, rows)
    assert not missed and n >= 0
    after = got["kfp"]["nodes_meta_runtime"]
    assert got["kfp"]["pipe_data"]["P3"]["length_m"] == pytest.approx(1.2)
    assert math.dist(before["N4"], after["N4"]["coords"]) > 1e-9
    assert math.dist(before["N1"], after["N1"]["coords"]) < 1e-9,         "급수원 쪽이 따라 흔들렸다"


def test_그_자리가_없으면_못_옮겼다고_말한다():
    """규칙 4 — 조용히 버리지 않는다."""
    got, board = _tiny()
    idx = ov.build_index(got, board)
    rows = [
        {"key": ov.key_to_json(ov.key_pipe(10, 11)), "kind": "pipe",
         "field": "length", "new": 3.0},
        {"key": ov.key_to_json(ov.key_pipe(900, 901)), "kind": "pipe",
         "field": "length", "new": 3.0},
        {"key": ["망가진키"], "kind": "pipe", "field": "length", "new": 3.0},
    ]
    applied, missed = ov.resolve(rows, idx)
    assert len(applied) == 1
    whys = [m["why"] for m in missed]
    assert len(missed) == 2
    assert any("범위에 없" in w for w in whys)
    assert any("읽지 못" in w for w in whys)


def test_resolve_는_사본이_아니라_원본을_준다():
    """적용하는 쪽이 채운 원값이 저장소에 그대로 남아야 한다(규칙 5).

    사본을 주면 채운 원값이 조용히 버려져 카드가 「원값 없음」을 보인다 —
    한 번 그렇게 만들었다가 고쳤다.
    """
    got, board = _tiny()
    idx = ov.build_index(got, board)
    rows = [{"key": ov.key_to_json(ov.key_pipe(10, 11)), "kind": "pipe",
             "field": "length", "new": 3.0}]
    (_k, r, _name), = ov.resolve(rows, idx)[0]
    r["old"] = 1.0
    assert rows[0]["old"] == 1.0, "채운 원값이 저장소에 안 남는다"


# ═══════════════════════════════════════════ ⑤ 적용 — kfp 메타 · 길이
def test_길이를_덮으면_그_배관만_바뀐다():
    got, board = _tiny()
    rows = ov.put([], ov.key_pipe(10, 11), "pipe", "length", 4.0,
                  reason="시험", at="t")
    n, missed, rep = ov.apply_to_kfp(got, board, rows)
    pr = got["kfp"]["pipe_data"]
    assert not missed
    assert pr["P1"]["length_m"] == pytest.approx(4.0)
    assert pr["P2"]["length_m"] == pytest.approx(1.0), "옆 배관이 따라 바뀌었다"
    assert rows[0]["old"] == pytest.approx(1.0), "원값을 안 채웠다"
    assert n >= 1 and rep is not None


def test_길이를_덮으면_좌표가_따라_움직인다():
    """오너(2026-09-14) 「20m 배관을 2m로 바꾸었는데, 아이소가 그대로인건
    말이안되니까」 — 회랑은 사슬이 길이로 좌표를 만든다."""
    got, board = _tiny()
    before = [list(v["coords"]) for v in got["kfp"]["nodes_meta_runtime"].values()]
    rows = ov.put([], ov.key_pipe(10, 11), "pipe", "length", 4.0,
                  reason="시험", at="t")
    ov.apply_to_kfp(got, board, rows)
    nd = got["kfp"]["nodes_meta_runtime"]
    after = [list(v["coords"]) for v in nd.values()]
    moved = sum(1 for a, b in zip(before, after) if math.dist(a[:2], b[:2]) > 1e-9)
    assert moved, "길이를 바꿨는데 좌표가 하나도 안 움직였다"
    # C2 — 길이 = 좌표 거리
    for pid, p in got["kfp"]["pipe_data"].items():
        a = nd[str(p["start"])]["coords"]
        b = nd[str(p["end"])]["coords"]
        assert math.dist(a, b) == pytest.approx(float(p["length_m"]), abs=1e-6), \
            f"{pid} 의 길이와 좌표 거리가 어긋났다"


def test_두_번_적용해도_같은_결과다():
    """멱등 — 표를 두 번 확정해도 값이 두 번 줄어들지 않는다."""
    got, board = _tiny()
    rows = ov.put([], ov.key_pipe(10, 11), "pipe", "length", 4.0,
                  reason="시험", at="t")
    ov.apply_to_kfp(got, board, rows)
    once = dict(got["kfp"]["pipe_data"]["P1"])
    ov.apply_to_kfp(got, board, rows)
    assert got["kfp"]["pipe_data"]["P1"]["length_m"] == pytest.approx(
        once["length_m"])


def test_메타에만_있는_값도_덮는다():
    got, board = _tiny()
    rows = ov.put([], ov.key_head(0), "head", "k_factor_si", 115.2)
    rows = ov.put(rows, ov.key_pipe(11, 12), "pipe", "roughness_mm", 0.05)
    n, missed, _rep = ov.apply_to_kfp(got, board, rows)
    assert not missed and n >= 2
    assert got["kfp"]["nodes_meta_runtime"]["N3"]["k_factor_si"] == 115.2
    assert got["kfp"]["pipe_data"]["P2"]["roughness_mm"] == 0.05
    assert rows[0]["old"] == 80.0, "K 값 원값을 안 채웠다"


# ═══════════════════════════════════════════ ⑥ 적용 — 표 칸
class _Tbl:
    def __init__(self, nodes, pipes):
        self.nodes, self.pipes = nodes, pipes


def _tiny_tbl():
    """표의 노드 좌표는 **mm 정수**다(§T3) — kfp 는 m. 단위가 갈리는 자리."""
    return _Tbl(
        nodes=[{"label": "1", "x": 0, "y": 0, "elevation": 0.0},
               {"label": "2", "x": 1000, "y": 0, "elevation": 0.0},
               {"label": "3", "x": 2000, "y": 0, "elevation": 0.0}],
        pipes=[{"label": "P1", "dia": 50, "dia_src": "규칙", "length": 1.0},
               {"label": "P2", "dia": 25, "dia_src": "규칙", "length": 1.0}])


def test_표고는_표_칸에서_덮인다():
    """★한때 m 좌표를 mm 칸에 맞대 **한 칸도 안 맞았고**(실측 0/241), 그런데도
    못 옮긴 목록에 안 떠서 표고 수정이 조용히 사라졌다(규칙 4 위반)."""
    got, board = _tiny()
    tbl = _tiny_tbl()
    rows = ov.put([], ov.key_node(11), "node", "elevation", 4.321)
    n, missed = ov.apply_to_tables(tbl, got, board, rows)
    assert not missed, f"못 옮겼다고 한다: {missed}"
    assert n == 1
    assert tbl.nodes[1]["elevation"] == 4.321
    assert tbl.nodes[0]["elevation"] == 0.0, "옆 절점이 따라 바뀌었다"
    assert rows[0]["old"] == 0.0


def test_헤드도_표고를_고칠_수_있다():
    """상하향 **종류**는 손질이 주인이지만(위상에 가깝다), 표고는 표의 한 칸이다.
    다른 절점은 다 고치는데 노즐만 못 고치면 그것은 규칙이 아니라 구멍이다."""
    got, board = _tiny()
    tbl = _tiny_tbl()
    rows = ov.put([], ov.key_head(0), "head", "elevation", -0.3)
    n, missed = ov.apply_to_tables(tbl, got, board, rows)
    assert not missed and n == 1
    assert tbl.nodes[2]["elevation"] == -0.3
    assert rows[0]["old"] == 0.0


def test_헤드_K값과_필요압력은_사유가_필요없다():
    """사유 필수는 길이뿐이다 — 전부 필수로 하면 사유 칸이 뜻을 잃는다."""
    for f in ("k_factor_si", "required_pressure_bar", "elevation"):
        assert ("head", f) not in ov.REASON_REQUIRED
        assert ov.validate("head", f, 1.0)[1] is None


def test_관경은_표_칸과_근거를_함께_바꾼다():
    got, board = _tiny()
    tbl = _tiny_tbl()
    rows = ov.put([], ov.key_pipe(10, 11), "pipe", "dia", 65)
    n, missed = ov.apply_to_tables(tbl, got, board, rows)
    assert not missed and n == 1
    assert tbl.pipes[0]["dia"] == 65
    assert tbl.pipes[0]["dia_src"] == "사람", "누가 정했는지가 안 남는다"
    assert tbl.pipes[1]["dia"] == 25


def test_관경_원값은_규칙이_낸_값이지_덮인_값이_아니다():
    """★실측(대명동): 도면 글씨로 50A 이던 배관을 100A 로 덮었더니 카드가
    「원값 100 → 100」 을 보였다.

    §6 으로 관경의 두 문을 한 벌로 모은 뒤, `decide_bores` 가 **이미 사람 값을
    넣은** 표가 여기로 들어온다. 그 행을 원값이라 읽으면 규칙이 낸 값이 영영
    사라진다(규칙 5 위반). 권위는 `bore_overrides[pid]["orig_dia"]` 에 있다.
    """
    got, board = _tiny()
    tbl = _tiny_tbl()
    tbl.pipes[0]["dia"] = 100           # decide_bores 가 이미 덮어 놓은 상태
    tbl.pipes[0]["dia_src"] = "사람"
    tbl.bore_overrides = {"P1": {"dia": 100, "orig_dia": 50,
                                 "orig_src": "text", "a": 10, "b": 11}}
    rows = ov.put([], ov.key_pipe(10, 11), "pipe", "dia", 100)
    n, missed = ov.apply_to_tables(tbl, got, board, rows)
    assert not missed and n == 1
    assert rows[0]["old"] == 50, f"원값이 {rows[0]['old']} 로 남았다 (50 이어야)"


def test_원값은_덮인_값으로_덮이지_않는다():
    """표를 두 번 확정해도 「원값 → 새값」이 그대로여야 한다.

    빌드마다 엔진이 새 망을 내므로 보통은 원값을 다시 적는 편이 낫지만
    (도면을 고치면 원값도 따라 고쳐진다), 같은 망에 두 번 적용하면 지금 값이
    **이미 사람 값**이다 — 그때만 건너뛴다.
    """
    for kind, field, new in (("pipe", "c", 100.0),
                             ("pipe", "roughness_mm", 0.05),
                             ("pipe", "type", "Sch10"),
                             ("head", "k_factor_si", 115.2),
                             ("node", "elevation", 4.321)):
        got, board = _tiny()
        tbl = _tiny_tbl()
        key = {"pipe": ov.key_pipe(10, 11), "head": ov.key_head(0),
               "node": ov.key_node(11)}[kind]
        rows = ov.put([], key, kind, field, new)
        _n, _m, rep = ov.apply_to_kfp(got, board, rows)
        ov.apply_to_tables(tbl, got, board, rows, rep)
        first = rows[0]["old"]
        _n, _m, rep2 = ov.apply_to_kfp(got, board, rows)
        ov.apply_to_tables(tbl, got, board, rows, rep2)
        assert rows[0]["old"] == first,             f"{kind}.{field} 원값이 {first} → {rows[0]['old']} 로 덮였다"
        assert rows[0]["old"] != new, f"{kind}.{field} 원값이 새값과 같다"


def test_고칠_수_있다고_말한_칸은_전부_어딘가에_닿는다():
    """카드는 `FIELDS` 를 전부 「고칠 수 있습니다」로 보인다. 적용 자리는 넷으로
    갈려 있어, 어느 자리도 안 맡는 (갈래, 속성)이 있으면 사람이 값을 넣었는데
    아무 일도 안 일어난다 — 못 옮김 목록에도 안 뜨면 규칙 4 위반이다.

    ★여기서는 **수리계산 갈래만** 전수로 본다(통합 갈래는 결합망이 필요하다 —
      `scripts/_probe_override_coverage.py` 가 그쪽까지 돈다).
    """
    val = {"length": 4.0, "dia": 65, "type": "Sch10", "c": 100.0,
           "roughness_mm": 0.05, "equivalent_length": 2.5,
           "elevation": -1.25, "k_factor_si": 115.2,
           "required_pressure_bar": 1.75}
    key_of = {"pipe": ov.key_pipe(10, 11), "node": ov.key_node(11),
              "head": ov.key_head(0), "vert": ov.key_vert("head", 0, 1)}
    quiet = []
    for (kind, field) in sorted(ov.FIELDS):
        if kind in ("sys", "mr"):
            continue
        got, board = _tiny_vert()
        tbl = _tiny_tbl()
        before = json.dumps([got["kfp"], tbl.nodes, tbl.pipes],
                            sort_keys=True, default=str)
        rows = ov.put([], key_of[kind], kind, field, val[field],
                      reason="전수", at="t")
        _n1, m1, rep = ov.apply_to_kfp(got, board, rows)
        _n2, m2 = ov.apply_to_tables(tbl, got, board, rows, rep)
        after = json.dumps([got["kfp"], tbl.nodes, tbl.pipes],
                           sort_keys=True, default=str)
        if after == before and not (list(m1) + list(m2)):
            quiet.append((kind, field))
    assert not quiet, f"조용히 사라지는 칸: {quiet}"


# ═══════════════════════════════════════════ ⑦ 통합 — 계통도 · 기계실
def _load_fx(fname, name):
    """옆 시험의 표본을 그대로 쓴다 — 표본이 갈리면 두 시험이 다른 것을 잰다."""
    import importlib.util
    path = os.path.join(_ROOT, "tests", fname)
    spec = importlib.util.spec_from_file_location(f"_fx_{fname}", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return getattr(mod, name)


def _merged(machineroom=None):
    from routes.module_f.merge import merge_network
    got = merge_network(_load_fx("test_module_f_merge.py", "_sample")(),
                        riser=_load_fx("test_module_f_merge.py", "_riser")(),
                        machineroom=machineroom, mode="lsp_gravity")
    # ★진짜 망의 배관 행에는 `c` 가 있다 — `emit_sdf` 가 `p["c"]` 를 반드시
    #   읽으므로 없으면 결합 산출이 KeyError 로 죽는다. 표본에는 없어서,
    #   `c` 를 고치는 시험이 「원값 없음」으로 통과해 버린다.
    for r in (got.get("combined").pipes if got.get("combined") else ()):
        r.setdefault("c", 120)
    return got


def test_통합에서_계통도_배관을_고친다():
    got = _merged()
    where = ov.merge_parts(got)
    lab = next(p for p, k in where["pipe"].items() if k == "system")
    rows = ov.put([], ov.key_merge("sys", lab), "sys", "length", 12.5,
                  reason="현장 실측", at="t")
    n, missed = ov.apply_to_merge(got, rows)
    assert not missed and n == 1
    row = next(r for r in got["combined"].pipes if str(r["label"]) == lab)
    assert float(row["length"]) == pytest.approx(12.5)
    assert rows[0]["old"] is not None, "원값을 안 채웠다"


def test_이음매는_고치지_않는다():
    """두 도면을 잇는 배관은 어느 도면 것도 아니다 — 어느 쪽을 고친 것인지
    말할 수 없는 값은 두지 않는다."""
    got = _merged()
    where = ov.merge_parts(got)
    seams = [p for p, k in where["pipe"].items() if k == "seam"]
    for lab in seams:
        rows = ov.put([], ov.key_merge("sys", lab), "sys", "length", 9.0,
                      reason="시험", at="t")
        n, missed = ov.apply_to_merge(got, rows)
        assert n == 0 and missed, f"{lab} 이음매가 덮였다"


def test_통합은_회랑_요소를_두_번_덮지_않는다():
    """회랑은 설계 표에서 이미 덮여 결합으로 흘러든다 — 여기서 또 덮으면
    한쪽만 고치는 날 두 화면이 갈린다."""
    got = _merged()
    before = [dict(r) for r in got["combined"].pipes]
    rows = ov.put([], ov.key_pipe(10, 11), "pipe", "length", 99.0,
                  reason="시험", at="t")
    n, missed = ov.apply_to_merge(got, rows)
    assert (n, missed) == (0, [])
    assert [dict(r) for r in got["combined"].pipes] == before


def test_기계실_요소도_같은_저장소로_고친다():
    """★`mr` 갈래는 여태 «못 옮김» 쪽으로만 지나갔다 — 기계실이 붙은 판이
    표본에 없었기 때문이다. 그 빈 자리로 「mr 은 아예 안 먹힌다」가 지나갈 수
    있으므로, 기계실을 실제로 붙여 놓고 잰다.
    """
    mr = _load_fx("test_module_f_seam_coords.py", "_machineroom")()
    got = _merged(machineroom=mr)
    where = ov.merge_parts(got)
    labs = [p for p, k in where["pipe"].items() if k == "machineroom"]
    assert labs, "기계실이 안 붙었다 — 표본을 다시 본다"
    lab = labs[0]
    row0 = next(r for r in got["combined"].pipes if str(r["label"]) == lab)
    was = float(row0["length"])
    rows = ov.put([], ov.key_merge("mr", lab), "mr", "length", 9.75,
                  reason="현장 실측", at="t")
    n, missed = ov.apply_to_merge(got, rows)
    assert not missed and n == 1, missed
    assert float(row0["length"]) == pytest.approx(9.75)
    assert rows[0]["old"] == pytest.approx(was)
    # 계통도 키로는 기계실 배관을 못 건드린다 — 갈래를 섞으면 안 된다
    got2 = _merged(machineroom=mr)
    bad = ov.put([], ov.key_merge("sys", lab), "sys", "length", 1.0,
                 reason="시험", at="t")
    n2, m2 = ov.apply_to_merge(got2, bad)
    assert n2 == 0 and m2 and "machineroom" in m2[0]["why"]


def test_통합_원값도_덮인_값으로_덮이지_않는다():
    got = _merged()
    where = ov.merge_parts(got)
    lab = next(p for p, k in where["pipe"].items() if k == "system")
    rows = ov.put([], ov.key_merge("sys", lab), "sys", "c", 100.0)
    ov.apply_to_merge(got, rows)
    first = rows[0]["old"]
    ov.apply_to_merge(got, rows)          # 같은 망에 두 번
    assert rows[0]["old"] == first != 100.0,         f"원값이 {first} → {rows[0]['old']} 로 덮였다"


def test_없는_라벨은_못_옮겼다고_말한다():
    got = _merged()
    rows = ov.put([], ov.key_merge("sys", "없는라벨"), "sys", "length", 3.0,
                  reason="시험", at="t")
    n, missed = ov.apply_to_merge(got, rows)
    assert n == 0 and len(missed) == 1
    assert "라벨" in missed[0]["why"]


# ═══════════════════════════════════════════ ⑧ §6 — 저장소를 둘 두지 않는다
def test_카드로_고친_관경이_옛_관경_목록과_한_벌이_된다():
    """관경은 문이 둘이다(F-11c 목록 · 이 카드). `decide_bores` 에 닿는 값은
    한 벌이어야 한다 — 안 그러면 숫자는 맞는데 **부속이 옛 관경으로** 정해진다.
    """
    from routes.module_f.api_design import _bore_ov_map
    sess = {
        "bore_overrides": [{"a": 3, "b": 9, "dia": 80, "note": "옛 문"}],
        "element_overrides": ov.put(
            [], ov.key_pipe(10, 11), "pipe", "dia", 65, reason="새 문"),
        "_ov_loaded_for": "",
    }
    got = _bore_ov_map(sess)
    assert got[(3, 9)] == (80, "옛 문"), "옛 목록이 사라졌다"
    assert got[(10, 11)] == (65, "새 문"), "카드로 고친 관경이 안 닿는다"


def test_같은_자리를_두_문으로_고치면_카드가_이긴다():
    from routes.module_f.api_design import _bore_ov_map
    sess = {
        "bore_overrides": [{"a": 3, "b": 9, "dia": 80, "note": "옛 문"}],
        "element_overrides": ov.put(
            [], ov.key_pipe(3, 9), "pipe", "dia", 100, reason="새 문"),
        "_ov_loaded_for": "",
    }
    assert _bore_ov_map(sess)[(3, 9)] == (100, "새 문")


# ═══════════════════════════════════════════ ⑨ 실도면 — 대명동 K=30
#
#   표본은 규약만 잰다. 「진짜 도면에서도 서는가」는 한 벌 돌려 봐야 안다 —
#   회랑 사슬·티 겹침·세로 토막이 다 걸린 망에서 **그 배관만** 바뀌는지.
#   도면이 없는 환경에서는 건너뛴다(픽스처 정책은 `tests/conftest.py`).

_DXF = os.path.join(_ROOT, "routes", "제출용[최종]",
                    "1. 입력도면 대명동 단위세대 평면도.dxf")


@pytest.mark.skipif(not os.path.exists(_DXF), reason="대명동 도면이 없다")
def test_실도면에서_길이와_C값을_고치면_그_값만_바뀐다():
    """끝에서 끝까지 — 라우트로 고치고, 표·좌표·파일을 센다.

    화면이 밟는 길을 그대로 밟는다(HTTP): 표 확정 → 고치기 → 표 재확정.
    """
    import argparse

    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    import importlib
    from _probe_candidate_drop import setup, wait
    from _workdir_iso import isolated_workdir
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True

    # ★작업폴더를 격리한다. 이 시험은 수정을 **파일로** 남기는데, 격리하지
    #   않으면 그 파일이 진짜 작업폴더에 남아 다음에 그 도면을 여는 사람이
    #   «내가 안 넣은 수정 4건» 을 보게 된다 — 실제로 한 번 그렇게 남겼다.
    with isolated_workdir(prefix="elov_"), srv.app.test_client() as c:
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
        assert c.post("/api/module-f/edit/worst",
                      json={"sid": sid, "k": 30}).get_json()["ok"]

        def rebuild():
            c.post("/api/module-f/design/build", json={"sid": sid, "k": 30})
            assert wait(c, sid).get("state") == "done"
            return c.get(f"/api/module-f/design/preview?sid={sid}").get_json()

        j0 = rebuild()
        pipes0 = {str(p["label"]): p for p in j0["view"]["pipes"]}
        rows0 = {str(r["label"]): r for r in j0["tables"]["pipes"]}
        # 키를 가진 평면 배관 중 가장 긴 것 — 바뀌는 것이 눈에 보이게
        lab = max((p for p in j0["view"]["pipes"]
                   if (p.get("key") or [None])[0] == "pipe"),
                  key=lambda p: float(rows0[str(p["label"])]["length"]))
        key, L0 = lab["key"], float(rows0[str(lab["label"])]["length"])
        c0 = rows0[str(lab["label"])].get("c")
        nodes0 = {str(n["label"]): (n["x"], n["y"]) for n in j0["view"]["nodes"]}

        # 사유 없이 길이를 바꾸면 거절해야 한다 (기준 8)
        bad = c.post("/api/module-f/design/override",
                     json={"sid": sid, "kind": "pipe", "key": key,
                           "field": "length", "new": L0 / 5}).get_json()
        assert bad.get("ok") is False and "사유" in str(bad.get("message"))

        ok = c.post("/api/module-f/design/override",
                    json={"sid": sid, "kind": "pipe", "key": key,
                          "field": "length", "new": round(L0 / 5, 3),
                          "reason": "현장 실측"}).get_json()
        assert ok["ok"], ok
        assert c.post("/api/module-f/design/override",
                      json={"sid": sid, "kind": "pipe", "key": key,
                            "field": "c", "new": 100}).get_json()["ok"]

        j1 = rebuild()
        rows1 = {str(r["label"]): r for r in j1["tables"]["pipes"]}
        lab1 = next(p for p in j1["view"]["pipes"] if p.get("key") == key)
        r1 = rows1[str(lab1["label"])]
        assert float(r1["length"]) == pytest.approx(round(L0 / 5, 3), abs=1e-3)
        assert float(r1["c"]) == pytest.approx(100.0)
        assert not j1.get("ov_missed"), j1.get("ov_missed")

        # ⑤ 그 값만 바뀐다 — 이름이 아니라 **키**로 맞댄다(이름은 옮겨 다닌다)
        key_of = {str(p["label"]): tuple(p["key"])
                  for p in j0["view"]["pipes"] if p.get("key")}
        len0 = {key_of[k]: float(v["length"])
                for k, v in rows0.items() if k in key_of}
        key_of1 = {str(p["label"]): tuple(p["key"])
                   for p in j1["view"]["pipes"] if p.get("key")}
        len1 = {key_of1[k]: float(v["length"])
                for k, v in rows1.items() if k in key_of1}
        moved = [k for k in len0
                 if k in len1 and k != tuple(key)
                 and abs(len0[k] - len1[k]) > 1e-6]
        assert not moved, f"옆 배관 {len(moved)}개가 따라 바뀌었다: {moved[:3]}"

        # ⑥ 길이가 좌표를 움직인다 — 아이소가 따라 변한다
        nodes1 = {str(n["label"]): (n["x"], n["y"]) for n in j1["view"]["nodes"]}
        shifted = sum(1 for k in nodes0 if k in nodes1
                      and math.dist(nodes0[k], nodes1[k]) > 1e-9)
        assert shifted, "길이를 5분의 1로 줄였는데 좌표가 하나도 안 움직였다"

        # ④ 원값·사유·시각이 남는다 (규칙 5)
        saved = {str(r["field"]): r for r in (j1.get("overrides") or [])}
        assert saved["length"]["old"] == pytest.approx(L0, abs=1e-3)
        assert saved["length"]["reason"] == "현장 실측"
        assert saved["length"]["at"]
        assert saved["c"]["old"] == pytest.approx(float(c0))

        # 지우면 원값으로 (기준 4)
        for f in ("length", "c"):
            c.post("/api/module-f/design/override",
                   json={"sid": sid, "kind": "pipe", "key": key,
                         "field": f, "remove": True})
        j2 = rebuild()
        rows2 = {str(r["label"]): r for r in j2["tables"]["pipes"]}
        lab2 = next(p for p in j2["view"]["pipes"] if p.get("key") == key)
        assert float(rows2[str(lab2["label"])]["length"]) == pytest.approx(
            L0, abs=1e-3)
