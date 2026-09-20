# -*- coding: utf-8 -*-
"""[통합 격자·활성·위상 §4-3] 위상 수정 — 지우고·더하고·다시 매긴다.

지시서 `ModuleF_통합_격자활성_위상수정_지시서.md` §3-3 · §3-6 · §4.

여기서 지키는 것 넷:

  1. **거절은 이유와 함께** — 갈라지는 배관 · 차수 ≥3 · 급수원 · 기준점 ·
     이음매. 조용히 통과시키면 산출에서야 터진다(§4 M6).
  2. **총연장은 보존된다** — 노드를 넣어도 빼도 배관 길이 합이 같다(M5).
     0.001 m 도 새면 안 된다. 반올림 잔차는 긴 쪽이 받는다.
  3. **멱등** — 같은 목록을 두 번 적용하면 완전히 같은 망(M3). 목록을 비우면
     원래 망으로 돌아온다.
  4. **이름은 다시 매기되 주소는 그대로**(§3-6 · D7) — 가운데에 넣으면 뒤
     번호가 한 칸씩 밀리고, 안정 키는 안 움직인다.
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys

import pytest

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
for _p in (_ROOT, os.path.join(_ROOT, "cad_project_editor_g"),
           os.path.join(_ROOT, "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from routes.module_f import overrides as ov            # noqa: E402


# ─────────────────────────────────────────────── 합성 망
def _node(nid, x, y, z=0.0, tid="base"):
    return {"id": nid, "coords": [float(x), float(y), float(z)],
            "elevation_m": float(z), "type": "기본", "type_id": tid,
            "category_id": "", "k_factor_si": None,
            "head_spec_name": None, "required_pressure_bar": 0.0}


def _pipe(a, b, length, dia=50):
    return {"start": a, "end": b, "type": "KSD 3507",
            "diameter": float(dia), "nominal_mm": int(dia),
            "length_m": float(length), "equivalent_length": 0.0,
            "C": 120, "roughness_mm": 0.045, "fittings": [],
            "flow_lpm": 0.0, "velocity_mps": 0.0, "headloss_m": 0.0}


def _chain():
    """회랑 5절점 한 줄 + 끝에 헤드 둘. N1 이 급수원(Input)이다.

        N1 ─P1─ N2 ─P2─ N3 ─P3─ N4 ─P4─ H1
                        └─P5─ H2
    """
    # ★급수원은 `type_id="pump"` 가 가린다(`design/anchor`) — io_node 가
    #   아니다. 여기를 틀리면 BFS 뿌리가 dict 순서로 정해져, 상류/하류가
    #   통째로 뒤집힌 채 시험이 «통과» 한다.
    nodes = {"N1": _node("N1", 0, 0, tid="pump"), "N2": _node("N2", 1, 0),
             "N3": _node("N3", 2, 0), "N4": _node("N4", 3, 0),
             "H1": _node("H1", 4, 0, 0.3, "head"),
             "H2": _node("H2", 2, 1, 0.3, "head")}
    pipes = {"P1": _pipe("N1", "N2", 1.0), "P2": _pipe("N2", "N3", 1.0),
             "P3": _pipe("N3", "N4", 1.0), "P4": _pipe("N4", "H1", 1.0),
             "P5": _pipe("N3", "H2", 1.0)}
    return {"nodes_meta_runtime": nodes, "pipe_data": pipes}


def _got(kfp=None, *, eref=None, nref=None):
    return {"kfp": kfp or _chain(),
            "edge_ref": dict(eref or {"P1": (0, 1), "P2": (1, 2),
                                      "P3": (2, 3), "P4": (3, 4),
                                      "P5": (2, 5)}),
            "node_ref": dict(nref or {"N1": 0, "N2": 1, "N3": 2,
                                      "N4": 3, "H1": 4, "H2": 5}),
            "origin_mm": (0.0, 0.0)}


class _Board:
    """`build_index` 가 보는 것만 흉내 낸다 — 좌표와 disk 목록."""

    def __init__(self, n=8):
        self.pts = [(float(i) * 1000.0, 0.0) for i in range(n)]
        self.disks = []
        self.hnodes = {}


def _op(op_id, op, target, payload=None, reason="시험"):
    return {"id": op_id, "op": op, "target": list(target),
            "payload": dict(payload or {}), "reason": reason, "at": ""}


def _fake_table(dia=50):
    """표가 선 «뒤» 를 흉내 낸다 — 기기 행이 붙을 자리만 있으면 된다."""
    from services.cad_import.design.tables import PipeTablesG

    t = PipeTablesG()
    t.pipes = [{"label": "P1", "in": "1", "out": "2", "dia": dia,
                "length": 1.0, "c": 120, "type": "KSD 3507"}]
    t.pipe_labels = {"P1": "P1"}
    return t


def _total(kfp):
    return round(sum(float(r.get("length_m") or 0.0)
                     for r in kfp["pipe_data"].values()), 6)


# ─────────────────────────────────────────────── §4 M6 — 거절 여섯
def test_갈라지는_배관은_거절한다():
    """나무에서는 어느 배관을 지워도 아래가 끊긴다 — 그래서 전부 거절이다."""
    got = _got()
    n, missed, _r = ov.apply_ops_to_kfp(
        got, _Board(), [_op("a1", "delete", ("pipe", 1, 2))])
    assert n == 0
    assert "끊깁니다" in missed[0]["why"], missed
    assert len(got["kfp"]["pipe_data"]) == 5     # 아무것도 안 지워졌다


def test_끊기는_쪽이_하나면_길을_알려_준다():
    """「안 됩니다」로 끝내지 않는다 — 절점을 지우면 된다고 말한다."""
    got = _got()
    _n, missed, _r = ov.apply_ops_to_kfp(
        got, _Board(), [_op("a1", "delete", ("pipe", 3, 4))])
    assert "그 절점을 지우면" in missed[0]["why"], missed


def test_차수_3_절점은_거절한다():
    got = _got()
    n, missed, _r = ov.apply_ops_to_kfp(
        got, _Board(), [_op("a1", "delete", ("node", 2))])   # N3 — 차수 3
    assert n == 0
    assert "3개 붙어" in missed[0]["why"], missed


def test_급수원은_거절한다():
    got = _got()
    n, missed, _r = ov.apply_ops_to_kfp(
        got, _Board(), [_op("a1", "delete", ("node", 0))])
    assert n == 0
    assert "급수원" in missed[0]["why"], missed


def test_기준점과_이음매는_통합에서_거절한다():
    """통합에만 있는 두 거절 — 기준점(라벨 10)과 이음매 배관."""
    from routes.module_f.merge import ANCHOR_LABEL

    got = _merged_fixture()
    n, missed, _r = ov.apply_ops_to_merge(
        got, [_op("a1", "delete", ("sys", ANCHOR_LABEL))])
    assert n == 0 and "기준점" in missed[0]["why"], missed

    n2, missed2, _r2 = ov.apply_ops_to_merge(
        got, [_op("a2", "delete", ("sys", "SEAM"))])
    assert n2 == 0 and "이음매" in missed2[0]["why"], missed2


# ─────────────────────────────────────────────── ⑩ 고리
def _loop_got():
    """N2 ─ N3 를 한 바퀴 돌게 이어 둔다 — 지워도 안 갈라진다."""
    kfp = _chain()
    kfp["pipe_data"]["P6"] = _pipe("N2", "N4", 2.0)
    got = _got(kfp)
    got["edge_ref"]["P6"] = (1, 3)
    return got


def test_고리가_있으면_지워도_된다():
    got = _loop_got()
    n, missed, rep = ov.apply_ops_to_kfp(
        got, _Board(), [_op("a1", "delete", ("pipe", 1, 2))])
    assert n == 1, missed
    assert "P2" not in got["kfp"]["pipe_data"]
    # 드문 일이라 보고에 센다 — 나중에 「왜 지워졌나」를 되짚을 재료다.
    assert rep["loop_pass"] == 1


def test_헤드는_지운_것만_센다():
    """D6 — 표 메타의 「사용자가 지운 헤드 N」으로 사람이 기준개수를 읽는다.

    거절된 삭제까지 세면 그 수가 거짓이 되고, **지운 뒤에** 세려 들면 이미
    지워져 「헤드였나」를 물을 자리가 없다(결합은 노즐 행을 함께 지운다).
    """
    got = _got()
    n, missed, rep = ov.apply_ops_to_kfp(
        got, _Board(), [_op("h1", "delete", ("node", 4))])   # H1 — 헤드
    assert (n, rep["heads_removed"]) == (1, 1), missed

    got2 = _got()
    n2, _m2, rep2 = ov.apply_ops_to_kfp(
        got2, _Board(), [_op("x", "delete", ("node", 2))])   # 차수 3 — 거절
    assert (n2, rep2["heads_removed"]) == (0, 0)


def test_헤드를_지우면_노즐도_함께_사라진다():
    """통합에서는 노즐이 **행**이라 절점만 지우면 고아가 남는다."""
    got = _merged_fixture()
    got["combined"].nozzles.append({"label": "1", "in": "12", "out": "@/1"})
    got["parts"]["plan"] = ["11", "12"]
    net = ov._MergeNet(got)
    assert net.is_head("12")
    net.drop_node("12")
    assert got["combined"].nozzles == []


# ─────────────────────────────────────────────── §4 M5 — 총연장
def test_차수_2_절점을_지우면_배관이_합쳐진다():
    got = _got()
    before = _total(got["kfp"])
    n, missed, _r = ov.apply_ops_to_kfp(
        got, _Board(), [_op("a1", "delete", ("node", 1))])   # N2 — 차수 2
    assert n == 1, missed
    pipes = got["kfp"]["pipe_data"]
    assert "N2" not in got["kfp"]["nodes_meta_runtime"]
    assert len(pipes) == 4
    keep = pipes["P1"]
    assert {keep["start"], keep["end"]} == {"N1", "N3"}
    assert keep["length_m"] == pytest.approx(2.0)
    assert _total(got["kfp"]) == pytest.approx(before, abs=1e-3)
    # 관경·관종·C 는 **상류 것**을 쓴다 — 그것이 규칙이다.
    assert keep["C"] == 120 and keep["nominal_mm"] == 50


def test_노드를_넣어도_총연장은_같다():
    got = _got()
    before = _total(got["kfp"])
    n, missed, rep = ov.apply_ops_to_kfp(
        got, _Board(), [_op("a1", "add_node", ("pipe", 0, 1), {"t": 0.37})])
    assert n == 1, missed
    pipes = got["kfp"]["pipe_data"]
    assert len(pipes) == 6
    new_pid = rep["added"]["pipes"]["a1"]
    assert pipes["P1"]["length_m"] == pytest.approx(0.37, abs=1e-6)
    assert pipes[new_pid]["length_m"] == pytest.approx(0.63, abs=1e-6)
    assert _total(got["kfp"]) == pytest.approx(before, abs=1e-6)
    # 관경·C 는 복사다 — 한 배관을 자른 것이니 같은 관이다.
    assert pipes[new_pid]["nominal_mm"] == pipes["P1"]["nominal_mm"]
    assert pipes[new_pid]["C"] == pipes["P1"]["C"]
    # 새 절점은 내분점에 선다(표고도 함께).
    xyz = got["kfp"]["nodes_meta_runtime"][rep["added"]["nodes"]["a1"]]["coords"]
    assert xyz[0] == pytest.approx(0.37, abs=1e-9)


def test_반올림_잔차는_긴_쪽이_받는다():
    """0.333 + 0.667 ≠ 1.000 이 되는 일이 실제로 난다 — 합을 먼저 지킨다."""
    for total, t in ((1.0, 1 / 3), (7.777, 0.5), (0.001, 0.5), (10.0, 0.37)):
        a, b = ov.split_lengths(total, t)
        assert round(a + b, 3) == round(total, 3), (total, t, a, b)


# ─────────────────────────────────────────────── §3-3-3 기기
def test_기기를_더하면_등가길이가_라이브러리_값이다():
    from services.cad_import.design.fitting import load_equivalent_lengths

    got = _got()
    n, missed, rep = ov.apply_ops_to_kfp(
        got, _Board(),
        [_op("a1", "add_equip", ("pipe", 0, 1),
             {"lib_id": "VALVE_GATE", "t": 0.5})])
    assert n == 1, missed
    assert len(rep["equip"]) == 1

    tbl = _fake_table()
    got_n = ov.apply_ops_to_tables(tbl, rep)
    assert got_n == 1
    row = tbl.equipment[-1]
    want = load_equivalent_lengths()["VALVE_GATE"][50]
    assert row["eq_len"] == pytest.approx(want)
    assert row["rel_pos"] == pytest.approx(0.5)
    # ★노드의 `fitting_id` 는 건드리지 않는다(§3-3-3). 기기는 «배관 위» 다.
    assert all("fitting_id" not in m
               for m in got["kfp"]["nodes_meta_runtime"].values())


def test_등가길이를_못_구하면_0_이_아니라_미해결이다():
    """0 은 「손실이 없다」는 주장이다 — 모르는 것과 같은 값으로 두지 않는다."""
    got = _got()
    _n, _m, rep = ov.apply_ops_to_kfp(
        got, _Board(),
        [_op("a1", "add_equip", ("pipe", 0, 1),
             {"lib_id": "VALVE_ALARM", "t": 0.5})])
    tbl = _fake_table()
    missed = []
    ov.apply_ops_to_tables(tbl, rep, missed=missed)
    assert tbl.equipment[-1]["eq_len"] is None
    assert len(missed) == 1 and "등가길이를 못 구했습니다" in missed[0]["why"]


def test_기기_목록은_서버가_라이브러리에서_뽑는다():
    cat = ov.equip_catalog()
    ids = [r["id"] for r in cat]
    assert "VALVE_GATE" in ids and "VALVE_ALARM" in ids
    # 스트레이너는 id 가 uuid 라 **분류**로 고른다 — 이름으로 박으면 못 찾는다.
    assert any(r["category"] == "strainer" for r in cat), cat


# ─────────────────────────────────────────────── §4 M3 — 멱등
def test_같은_목록을_두_번_적용하면_같은_망이다():
    ops = [_op("a1", "add_node", ("pipe", 0, 1), {"t": 0.37}),
           _op("a2", "delete", ("node", 1))]
    outs = []
    for _ in range(2):
        got = _got()
        ov.apply_ops_to_kfp(got, _Board(), ops)
        outs.append(json.dumps(got["kfp"], sort_keys=True, default=str))
    assert outs[0] == outs[1]


def test_목록을_비우면_원래_망이다():
    base = json.dumps(_got()["kfp"], sort_keys=True, default=str)
    got = _got()
    ov.apply_ops_to_kfp(got, _Board(), [])
    assert json.dumps(got["kfp"], sort_keys=True, default=str) == base


def test_순서는_배열이_아니라_규칙이_정한다():
    """삭제 → 노드추가 → 기기추가. 목록에 거꾸로 담아도 같은 순서로 돈다."""
    rows = [_op("c", "add_equip", ("pipe", 0, 1), {"lib_id": "VALVE_GATE"}),
            _op("b", "add_node", ("pipe", 0, 1), {"t": 0.5}),
            _op("a", "delete", ("node", 1))]
    assert [r["op"] for r in ov.ops_sorted(rows)] == [
        "delete", "add_node", "add_equip"]


# ─────────────────────────────────────────────── 저장소 — version 1·2
def test_version1_파일을_읽어도_터지지_않는다(tmp_path, monkeypatch):
    p = tmp_path / "K_수리계산수정.json"
    p.write_text(json.dumps({"version": 1, "items": [{"key": ["node", 3]}]}),
                 encoding="utf-8")
    monkeypatch.setattr(ov, "path_for", lambda key: str(p))
    assert len(ov.read_file("K")) == 1
    assert ov.read_file_ops("K") == []        # ops 없음으로 읽는다


def test_값만_써도_위상이_지워지지_않는다(tmp_path, monkeypatch):
    """한 파일을 둘이 쓴다 — 한쪽이 모르고 덮으면 다른 쪽이 사라진다."""
    p = tmp_path / "K_수리계산수정.json"
    monkeypatch.setattr(ov, "path_for", lambda key: str(p))
    ops = [_op("a1", "delete", ("node", 3))]
    ov.write_file("K", [{"key": ["node", 3], "field": "dia", "new": 50}], ops)
    assert json.loads(p.read_text(encoding="utf-8"))["version"] == 2
    # 값만 아는 옛 호출자가 저장한다 — 위상은 그대로 있어야 한다.
    ov.write_file("K", [{"key": ["node", 3], "field": "dia", "new": 65}])
    assert len(ov.read_file_ops("K")) == 1


def test_둘_다_비면_파일을_지운다(tmp_path, monkeypatch):
    p = tmp_path / "K_수리계산수정.json"
    monkeypatch.setattr(ov, "path_for", lambda key: str(p))
    ov.write_file("K", [{"key": ["node", 3]}], [_op("a1", "delete", ("node", 3))])
    assert p.exists()
    ov.write_file("K", [], [])
    assert not p.exists()


def test_새_요소의_주소는_그것을_만든_id_다():
    got = _got()
    _n, _m, rep = ov.apply_ops_to_kfp(
        got, _Board(), [_op("zz", "add_node", ("pipe", 0, 1), {"t": 0.5})])
    assert rep["added"]["nodes"]["zz"] == "UXzz"
    assert ov.key_from_json(["add", "zz"]) == ("add", "zz")


def test_쪼갠_조각은_주소록에서_뺀다():
    """두 조각이 같은 board 쌍을 쓴다 — 주소는 **원 조각**이 지킨다(§5)."""
    got = _got()
    ov.apply_ops_to_kfp(got, _Board(),
                        [_op("a1", "add_node", ("pipe", 0, 1), {"t": 0.5})])
    idx = ov.build_index(got, _Board())
    assert idx["pipe"][ov.key_pipe(0, 1)] == "P1"      # 새 조각이 아니다


# ─────────────────────────────────────────────── §3-3-1 ⑤ 새 요소의 값 수정
def test_방금_만든_절점의_값을_고칠_수_있다():
    """§3-3-1 ⑤ — 순서가 ①값 → ②삭제 → ③노드 → ④기기 → ⑤**새 요소의 값** 인
    이유가 이것이다. 만들기 전에는 가리킬 자리가 없다.

    ★이 자리가 비어 있었다: 새 절점은 board 에 대응이 없어 어떤 안정 키로도
      안 잡혔고, 값을 고쳐 두면 「그 자리가 이번 계산 범위에 없습니다」로
      **조용히** 떨어졌다.
    """
    got = _got()
    _n, _m, rep = ov.apply_ops_to_kfp(
        got, _Board(), [_op("zz", "add_node", ("pipe", 0, 1), {"t": 0.5})])
    idx = ov.build_index(got, _Board())
    # 절점도 배관도 그 id 로 가리켜진다 — 갈래(kind)가 둘을 가른다.
    assert idx["node"][ov.key_added("zz")] == rep["added"]["nodes"]["zz"]
    assert idx["pipe"][ov.key_added("zz")] == rep["added"]["pipes"]["zz"]

    rows = ov.put([], ov.key_added("zz"), "node", "elevation", 1.234,
                  reason="시험")
    applied, missed = ov.resolve(rows, idx)
    assert not missed and len(applied) == 1, missed


def test_통합에서_만든_절점도_그_id_로_가리켜진다():
    """라벨은 결합할 때마다 새로 매겨진다 — id 만이 안정된 주소다."""
    got = _merged_fixture()
    n, missed, rep = ov.apply_ops_to_merge(
        got, [_op("qq", "add_node", ("sys", "r1"), {"t": 0.4})])
    assert n == 1, missed
    born = got["ops_added"]
    assert born["merge_by_id"]["qq"] == rep["added"]["nodes"]["qq"]

    rows = ov.put([], ov.key_added("qq"), "sys", "elevation", 2.5,
                  reason="시험")
    n2, m2 = ov.apply_to_merge(got, rows)
    assert (n2, m2) == (1, []), m2
    lab = born["merge_by_id"]["qq"]
    row = next(r for r in got["combined"].nodes if str(r["label"]) == lab)
    assert row["elevation"] == 2.5


def test_통합이_안_만든_요소는_통합이_손대지_않는다():
    """회랑에서 만든 요소는 설계 표에서 이미 고쳐져 흘러든다 — 두 번 덮지 않는다."""
    got = _merged_fixture()
    rows = ov.put([], ov.key_added("없는id"), "sys", "length", 9.9,
                  reason="시험")
    assert ov.apply_to_merge(got, rows) == (0, [])


# ─────────────────────────────────────────────── §3-6 이름 다시 매기기
def test_가운데에_넣으면_뒤_번호가_한_칸씩_민다():
    from services.cad_import.design.tables import build_design_tables

    def labels(got):
        tbl = build_design_tables(got["kfp"], {"heads": [], "loads": {}},
                                  got["edge_ref"], [],
                                  default_schedule="KSD 3507")
        return [r["label"] for r in tbl.pipes], tbl

    before, tbl0 = labels(_got())
    assert before == ["P1", "P2", "P3", "P4", "P5"]
    got = _got()
    ov.apply_ops_to_kfp(got, _Board(),
                        [_op("a1", "add_node", ("pipe", 0, 1), {"t": 0.5})])
    after, tbl1 = labels(got)
    assert len(after) == 6
    assert after == ["P1", "P2", "P3", "P4", "P5", "P6"]
    # ★뒤 번호가 **한 칸씩** 민다. 「밀렸다」만 보면 뒤죽박죽이어도 통과하니,
    #   민 거리가 정확히 1 인지 본다(⑧ 이 요구하는 것이 그것이다).
    #   자른 배관(P1) 자신은 앞 조각이라 제자리다.

    def _n(lab):
        return int(str(lab)[1:])

    for pid in ("P2", "P3", "P4", "P5"):
        assert _n(tbl1.pipe_labels[pid]) == _n(tbl0.pipe_labels[pid]) + 1, (
            pid, tbl0.pipe_labels, tbl1.pipe_labels)
    assert tbl1.pipe_labels["P1"] == tbl0.pipe_labels["P1"] == "P1"
    # 주소(pid)는 그대로 살아 있다 — 저장된 수정이 옆 배관으로 안 옮겨간다.
    assert set(tbl1.pipe_labels) >= {"P1", "P2", "P3", "P4", "P5"}


def test_부속표가_가리키는_배관이_배관표에_전부_있다():
    """§4 M10 — 고아 0. 이름을 다시 매기면서 한쪽만 고치면 여기서 걸린다."""
    from services.cad_import.design.tables import build_design_tables

    got = _got()
    tbl = build_design_tables(got["kfp"], {"heads": [], "loads": {}},
                              got["edge_ref"], [],
                              default_schedule="KSD 3507")
    have = {str(r["label"]) for r in tbl.pipes}
    for r in tbl.fittings:
        assert str(r["pipe"]) in have, (r, sorted(have))
    for r in tbl.equipment:
        assert str(r["pipe"]) in have, (r, sorted(have))


def test_같은_도면을_두_번_만들면_이름이_같다():
    from services.cad_import.design.tables import build_design_tables

    out = []
    for _ in range(2):
        got = _got()
        tbl = build_design_tables(got["kfp"], {"heads": [], "loads": {}},
                                  got["edge_ref"], [],
                                  default_schedule="KSD 3507")
        out.append([r["label"] for r in tbl.pipes])
    assert out[0] == out[1]


# ─────────────────────────────────────────────── 통합 — 합성 결합망
def _merged_fixture():
    """회랑 3 + 라이저 3 + 이음매 하나. 기준점은 라벨 10 이다."""
    from routes.module_f.merge import ANCHOR_LABEL

    class _HT:
        pass

    c = _HT()
    c.nodes = [
        {"label": "1", "x": 0, "y": 0, "elevation": 0.0, "io_node": "Input"},
        {"label": ANCHOR_LABEL, "x": 1000, "y": 0, "elevation": 0.0,
         "io_node": "No"},
        {"label": "11", "x": 2000, "y": 0, "elevation": 0.0, "io_node": "No"},
        {"label": "12", "x": 3000, "y": 0, "elevation": 0.0, "io_node": "No"},
    ]
    c.pipes = [
        {"label": "r1", "in": "1", "out": ANCHOR_LABEL, "dia": 100,
         "length": 1.0, "c": 120, "elev": 0.0},
        {"label": "SEAM", "in": ANCHOR_LABEL, "out": "11", "dia": 100,
         "length": 1.0, "c": 120, "elev": 0.0},
        {"label": "P1", "in": "11", "out": "12", "dia": 50,
         "length": 1.0, "c": 120, "elev": 0.0},
    ]
    c.nozzles, c.fittings, c.equipment = [], [], []
    return {"combined": c,
            "parts": {"system": ["1", ANCHOR_LABEL], "plan": ["11", "12"],
                      "machineroom": []}}


def test_통합에서_계통도_배관을_지운다():
    got = _merged_fixture()
    # r1 은 급수원과 기준점을 잇는다 — 지우면 갈라지니 거절이어야 한다.
    n, missed, _r = ov.apply_ops_to_merge(
        got, [_op("a1", "delete", ("sys", "r1"))])
    assert n == 0 and "끊깁니다" in missed[0]["why"], missed


def test_통합에서_회랑_요소는_손대지_않는다():
    """회랑 위상은 04 에서 이미 먹었다 — 여기서 또 먹으면 두 번 지운다."""
    got = _merged_fixture()
    n, missed, _r = ov.apply_ops_to_merge(
        got, [_op("a1", "delete", ("pipe", 0, 1))])
    assert (n, missed) == (0, [])
    assert len(got["combined"].pipes) == 3


def test_통합에서_노드를_넣으면_길이_합이_같다():
    got = _merged_fixture()
    before = round(sum(float(r["length"]) for r in got["combined"].pipes), 6)
    n, missed, rep = ov.apply_ops_to_merge(
        got, [_op("a1", "add_node", ("sys", "r1"), {"t": 0.4})])
    assert n == 1, missed
    after = round(sum(float(r["length"]) for r in got["combined"].pipes), 6)
    assert after == pytest.approx(before, abs=1e-3)
    assert len(got["combined"].nodes) == 5
    assert rep["added"]["nodes"]["a1"] == "13"      # ⑧ 맨 끝 번호 + 1


# ─────────────────────────────────────────────── ⑨ 두 산출이 같은 말을
def test_회랑_수정은_04_와_통합_양쪽에_실린다():
    """§4 M9 — 적용 자리가 **하나**라 두 파일이 갈릴 수가 없다.

    회랑 대상 op 는 `apply_ops_to_kfp` 에서만 먹고, 그 결과로 만든 표가
    그대로 결합으로 흘러든다. 그래서 여기서 확인할 것은 「통합 쪽이 회랑
    op 를 **다시** 먹지 않는다」이다 — 두 번 먹으면 두 배로 지운다.
    """
    corridor = _op("a1", "delete", ("node", 1))
    got = _got()
    n1, _m1, _r1 = ov.apply_ops_to_kfp(got, _Board(), [corridor])
    assert n1 == 1
    mg = _merged_fixture()
    n2, m2, _r2 = ov.apply_ops_to_merge(mg, [corridor])
    assert (n2, m2) == (0, [])


# ─────────────────────────────────────────────── §3-1-3 배율 유추 (화면)
_STUB = "const S = {};\n"


def _js_fn(name, head):
    js = open(os.path.join(_ROOT, "static", "module_f.js"),
              encoding="utf-8").read()
    i = js.index(head)
    return js[i:js.index("\n  }\n", i) + 4], name


def _run_infer(pipes, at, body):
    node = shutil.which("node")
    if node is None:
        pytest.skip("node 가 없다 — 화면 코드를 돌릴 수 없다")
    src, _ = _js_fn("inferMeterScale", "  function inferMeterScale(")
    prog = "\n".join([f"const PIPES = {json.dumps(pipes)};",
                      f"const AT = {json.dumps(at)};", _STUB, src, body])
    out = subprocess.run([node, "-e", prog], capture_output=True, text=True,
                         encoding="utf-8", errors="replace")
    assert out.returncode == 0, out.stderr[-800:]
    return json.loads(out.stdout)


def _scaled(k):
    """길이 1 m 짜리 수평 배관 넷 — 좌표만 k 배로 늘린다."""
    at = {str(i): [i * k, 0] for i in range(5)}
    pipes = [{"in": str(i), "out": str(i + 1), "length": 1.0, "elev": 0.0}
             for i in range(4)]
    return pipes, at


def test_배율을_표에서_되찾는다():
    got = _run_infer(*_scaled(3.7),
                     body="console.log(JSON.stringify("
                          "inferMeterScale(PIPES, AT)));")
    assert got["s"] == pytest.approx(3.7, rel=1e-6), got
    assert got["n"] == 4 and got["off"] == 0


def test_쓸_배관이_셋_미만이면_안_깐다():
    """지어내지 않는다 — 두 개로 낸 «배율» 은 배율이 아니다."""
    pipes, at = _scaled(3.7)
    got = _run_infer(pipes[:2], at,
                     body="console.log(JSON.stringify("
                          "inferMeterScale(PIPES, AT)));")
    assert got is None, got


def test_수직_배관만_있으면_못_구함이다():
    """표고차가 길이의 25% 를 넘는 배관은 «수평» 이 아니라 자에서 뺀다."""
    at = {str(i): [0, 0] for i in range(5)}
    pipes = [{"in": str(i), "out": str(i + 1), "length": 1.0, "elev": 1.0}
             for i in range(4)]
    got = _run_infer(pipes, at,
                     body="console.log(JSON.stringify("
                          "inferMeterScale(PIPES, AT)));")
    assert got is None, got


# ─────────────────────────────────────────────── 라우트 — 문 하나
def _client():
    """`_fail` 은 `{"ok": false, "message": …}` 로 답한다 — 화면이 읽는 칸이
    그 하나라, 시험도 같은 칸을 본다."""
    import importlib

    for p in (_ROOT, os.path.join(_ROOT, "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    srv = importlib.import_module("대조 서버")
    srv.app.config["TESTING"] = True
    c = srv.app.test_client()
    with c.session_transaction() as s:
        s["authed"] = True
    return c


def _sid(tmp_path, monkeypatch):
    from routes.module_f.jobs import _new_session

    sess = _new_session()
    sess["key"] = "시험도면"
    monkeypatch.setattr(
        ov, "path_for",
        lambda key: str(tmp_path / f"{key}_수리계산수정.json"))
    return sess["id"], sess


def test_사유가_없으면_저장하지_않는다(tmp_path, monkeypatch):
    """왜 망을 고쳤는지가 안 남으면 그 계산서를 다음 사람이 못 믿는다."""
    c = _client()
    sid, _s = _sid(tmp_path, monkeypatch)
    r = c.post("/api/module-f/element/op",
               json={"sid": sid, "op": "delete", "target": ["node", 3],
                     "payload": {}})
    assert r.status_code >= 400, r.get_json()
    assert "사유" in (r.get_json() or {}).get("message", "")


def test_모르는_수정은_거절한다(tmp_path, monkeypatch):
    c = _client()
    sid, _s = _sid(tmp_path, monkeypatch)
    r = c.post("/api/module-f/element/op",
               json={"sid": sid, "op": "move", "target": ["node", 3],
                     "reason": "시험"})
    assert r.status_code >= 400
    assert "모르는 위상 수정" in (r.get_json() or {}).get("message", "")


def test_한_문에_쌓이고_파일로_남는다(tmp_path, monkeypatch):
    """어느 화면에서 눌렀든 같은 목록이다(⑨) — 라우트가 하나인 이유."""
    c = _client()
    sid, sess = _sid(tmp_path, monkeypatch)
    r = c.post("/api/module-f/element/op",
               json={"sid": sid, "op": "delete", "target": ["sys", "r3"],
                     "reason": "계통도 쪽"})
    assert r.status_code == 200, r.get_json()
    d = r.get_json()
    assert len(d["ops"]) == 1 and d["ops"][0]["op"] == "delete"
    rid = d["ops"][0]["id"]
    assert len(rid) == 8

    r2 = c.post("/api/module-f/element/op",
                json={"sid": sid, "op": "add_node", "target": ["pipe", 3, 9],
                      "payload": {"t": 0.4}, "reason": "회랑 쪽"})
    assert len(r2.get_json()["ops"]) == 2
    # 파일에 남는다 — 서버를 껐다 켜도 그대로다(D2 write-through).
    assert len(ov.read_file_ops("시험도면")) == 2

    # 되돌리기 — Ctrl+Z 가 부르는 그 길이다.
    r3 = c.post("/api/module-f/element/op", json={"sid": sid, "remove": rid})
    assert len(r3.get_json()["ops"]) == 1
    assert len(ov.read_file_ops("시험도면")) == 1
    _ = sess


def test_고를_수_있는_기기를_서버가_준다(tmp_path, monkeypatch):
    c = _client()
    sid, _s = _sid(tmp_path, monkeypatch)
    d = c.get(f"/api/module-f/element/op?sid={sid}").get_json()
    ids = [r["id"] for r in d["equip"]]
    assert "VALVE_GATE" in ids and len(ids) >= 10, ids


def test_거절은_저장_전에_난다(tmp_path, monkeypatch):
    """목록에 넣어 두고 「표 확정」에서 터뜨리면, 그때는 왜 막혔는지 화면에
    남아 있지 않다."""
    c = _client()
    sid, sess = _sid(tmp_path, monkeypatch)
    sess["merged"] = _merged_fixture()
    r = c.post("/api/module-f/element/op",
               json={"sid": sid, "op": "delete", "target": ["sys", "SEAM"],
                     "reason": "시험"})
    assert r.status_code >= 400
    assert "이음매" in (r.get_json() or {}).get("message", "")
    assert ov.load_ops(sess) == []          # 저장되지 않았다


def test_사람이_더한_기기는_kfp_등가길이에_더해진다():
    """§3-3-3 · §4-2-6 — `.sdf` 에는 `<Equipment>` 로, `.kfp` 에는 그 배관의
    등가길이에 **더해서**. 안 더하면 같은 망인데 두 파일의 손실이 달라진다."""
    from services.cad_import.design.emit import emit_design_kfp
    from services.cad_import.design.tables import build_design_tables

    got = _got()
    _n, _m, rep = ov.apply_ops_to_kfp(
        got, _Board(),
        [_op("a1", "add_equip", ("pipe", 0, 1),
             {"lib_id": "VALVE_GATE", "t": 0.5})])
    tbl = build_design_tables(got["kfp"], {"heads": [], "loads": {}},
                              got["edge_ref"], [],
                              bores={p: (50, "시험") for p in got["kfp"]["pipe_data"]},
                              default_schedule="KSD 3507")
    ov.apply_ops_to_tables(tbl, rep)
    add = [e for e in tbl.equipment if e.get("op_id") == "a1"]
    assert len(add) == 1 and add[0]["eq_len"] is not None, tbl.equipment

    import tempfile
    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "t.kfp")
        emit_design_kfp(tbl, got, out)
        pd = json.load(open(out, encoding="utf-8"))["pipe_data"]
    # 그 배관(pid P1)의 등가길이가 기기 값만큼 늘었다.
    base = [float(r.get("eq_len") or 0.0) for r in tbl.pipes
            if str(r["label"]) == tbl.pipe_labels["P1"]][0]
    assert float(pd["P1"]["equivalent_length"]) == pytest.approx(
        round(base + add[0]["eq_len"], 3), abs=1e-6), pd["P1"]


def test_등가길이를_못_구한_기기는_sdf_에_0_으로_안_실린다():
    """§4 M8 — 0 은 「손실이 없다」는 주장이다. 안 싣고 세어 올린다."""
    from services.cad_import.design.emit import tables_to_network
    from services.cad_import.design.tables import PipeTablesG

    t = PipeTablesG()
    t.nodes = [{"label": "1", "x": 0, "y": 0, "elevation": 0.0,
                "io_node": "Input"},
               {"label": "2", "x": 1000, "y": 0, "elevation": 0.0,
                "io_node": "No"}]
    t.pipes = [{"label": "P1", "in": "1", "out": "2", "dia": 25,
                "length": 1.0, "elev": 0.0, "c": 120, "status": "Normal",
                "type": "KSD 3507"}]
    t.equipment = [{"pipe": "P1", "in": "1", "out": "2", "label": "1",
                    "desc": "게이트 밸브", "eq_len": None, "rel_pos": 0.5,
                    "op_id": "z1"}]
    net = tables_to_network(t, project_title="시험")
    eq = getattr(net.pipes["P1"], "equipment", None) or []
    assert eq == [], eq


# ─────────────────────────────────────────────── 검토가 잡아낸 셋
def test_부속이_달린_배관에도_기기_등가길이가_kfp_에_더해진다():
    """★부속 라벨이 **배관 이름을 덮어쓰고** 있었다.

    `for kind in by_pipe.get(lab, ()): lab = fitting_label(...)` — 그 뒤 기기
    합산이 `lab` 으로 배관을 찾으므로, 부속이 하나라도 달린 배관에서는 사람이
    더한 기기의 등가길이가 **통째로 빠졌다**. 부속이 없는 배관에서만 우연히
    맞아서 첫 시험은 통과했다 — 그래서 여기서는 **부속을 달아** 둔다.
    """
    from services.cad_import.design.emit import emit_design_kfp
    from services.cad_import.design.tables import build_design_tables

    got = _got()
    _n, _m, rep = ov.apply_ops_to_kfp(
        got, _Board(),
        [_op("a1", "add_equip", ("pipe", 0, 1),
             {"lib_id": "VALVE_GATE", "t": 0.5})])
    tbl = build_design_tables(
        got["kfp"], {"heads": [], "loads": {}}, got["edge_ref"], [],
        bores={p: (50, "시험") for p in got["kfp"]["pipe_data"]},
        default_schedule="KSD 3507")
    ov.apply_ops_to_tables(tbl, rep)
    lab = tbl.pipe_labels["P1"]
    # 그 배관에 부속을 하나 달아 둔다 — 덮어쓰기가 살아 있으면 여기서 걸린다.
    tbl.fittings.append({"pipe": lab, "in": "1", "out": "2",
                         "type": "elbow", "count": "1"})
    eq = [e for e in tbl.equipment if e.get("op_id") == "a1"][0]["eq_len"]
    assert eq is not None

    import tempfile
    with tempfile.TemporaryDirectory() as d:
        out = os.path.join(d, "t.kfp")
        emit_design_kfp(tbl, got, out)
        pd = json.load(open(out, encoding="utf-8"))["pipe_data"]
    base = [float(r.get("eq_len") or 0.0) for r in tbl.pipes
            if str(r["label"]) == lab][0]
    assert pd["P1"]["fittings"] == ["Elbow"], pd["P1"]["fittings"]
    assert float(pd["P1"]["equivalent_length"]) == pytest.approx(
        round(base + eq, 3), abs=1e-6), pd["P1"]


def test_결합표_기기_목록이_비어_있어도_단_것이_남는다():
    """★`or []` 는 빈 목록에서 **새 리스트**를 만든다 — 기기 목록은 보통
    비어 있으므로, 사람이 단 밸브가 「적용했습니다」라고 보고된 채 산출물에서
    사라졌다."""
    got = _merged_fixture()
    assert got["combined"].equipment == []          # 보통의 상태
    n, missed, _r = ov.apply_ops_to_merge(
        got, [_op("e1", "add_equip", ("sys", "r1"),
                  {"lib_id": "VALVE_GATE", "t": 0.5})])
    assert n == 1, missed
    # ★결합표 **그 객체**에 남아야 한다 — 어댑터가 들고 있던 사본이 아니라.
    assert len(got["combined"].equipment) == 1, got["combined"].equipment
    assert got["combined"].equipment[0]["lib"] == "VALVE_GATE"


def test_기기는_담당_헤드_수로_붙을_관을_고른다():
    """★`tree_loads` 는 kfp 이름 키인데 표 라벨로 찾고 있었다 — 담당 헤드 수가
    전부 0 이 되어 FX·알람밸브가 「물이 지나는 관」이 아니라 호칭경·이름 순으로
    붙었다. 두 이름이 같은 글자꼴이라 예외도 빈 값도 안 났다."""
    from services.cad_import.design.tables import build_design_tables

    got = _got()
    # N3(표에서 차수 3) 에 알람밸브를 찍는다. 붙을 후보는 P2(상류·담당 많음) ·
    # P3 · P5(곁가지·담당 0). 담당 헤드 수를 못 읽으면 곁가지에 붙는다.
    tbl = build_design_tables(
        got["kfp"], {"heads": [], "loads": {}}, got["edge_ref"], [],
        bores={"P1": (50, "시"), "P2": (50, "시"), "P3": (25, "시"),
               "P4": (25, "시"), "P5": (25, "시")},
        valve_nodes=["N3"],
        tree_loads={"P1": 3, "P2": 3, "P3": 1, "P4": 1, "P5": 1},
        default_schedule="KSD 3507")
    av = [e for e in tbl.equipment if e["desc"] == "A/V"][0]
    assert av["pipe"] == tbl.pipe_labels["P2"], (av["pipe"], tbl.pipe_labels)
