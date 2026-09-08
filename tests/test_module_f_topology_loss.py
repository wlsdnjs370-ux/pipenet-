# -*- coding: utf-8 -*-
"""[위상 손실] 통합 경계 다섯 결함 — 지시서 `ModuleF_위상손실_수정지시서.md`.

평면도·계통도·기계실 **추출 셋은 정상**이고, 깨지던 곳은 두 경계였다.
여기서 그 다섯을 못박는다.

  D1 기계실이 통합 좌표계로 안 옮겨진다 (항상)
  D2 고아 참조 검사가 «?» 를 면제한다        → 세기
  D3 메타 없는 절점이 원점에 생긴다          → 세기
  D4 배관 개명이 부속·기기를 안 데려간다
  D5 결합 뒤 검사가 하나도 없다              → 신설(보고)

■ D1 실측 (대명동 3장 · `scripts/_f_topology_probe.py`)

    bbox span   986,199 mm → 37,261 mm   (평면도 단독 36,150 수준)
    emit 배율   0.003042 → 0.080513
    기계실 좌표 원본 그대로  11/12 → 0
    plan_edges  0 → 33

  원인: `prepend_machine_room_to_riser` 뒤의 `rt.nodes[0]` 은 기계실 수원(m1)
  인데 그것을 `pump_junction_label` 로 넘겼다. `stitch` 는 그 라벨을 기계실을
  **이미 제외한** 목록에서 찾으므로 `pump_node` 가 항상 None → 기계실이 원
  DXF 좌표에 방치됐다.
"""
from __future__ import annotations

import importlib.util
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _fx(name):
    """옆 시험의 표본 — `tests` 는 패키지가 아니라 경로로 읽는다."""
    path = _ROOT / "tests" / "test_module_f_merge.py"
    spec = importlib.util.spec_from_file_location("_mf_fx", str(path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return getattr(mod, name)


def _machineroom(x0=900000.0, y0=400000.0, n=4):
    """기계실 경로 — **원 DXF 좌표**(수십만 mm). 라이저·평면도와 다른 좌표계."""
    labels = [f"m{i + 1}" for i in range(n)]
    nodes = [{"label": lab, "x": x0 + i * 1500.0, "y": y0 + i * 800.0,
              "elevation": 0.0,
              "io_node": "Input" if i == 0 else "No"}
             for i, lab in enumerate(labels)]
    pipes = [{"label": f"mp{i}", "in": labels[i], "out": labels[i + 1],
              "dia": 100, "length": 1.7} for i in range(n - 1)]
    return {"nodes": nodes, "pipes": pipes,
            "conn_node_label": labels[-1],
            "conn_xy": [nodes[-1]["x"], nodes[-1]["y"]],
            "plan_edges": [[x0, y0, x0 + 2000.0, y0],
                           [x0 + 2000.0, y0, x0 + 2000.0, y0 + 1200.0]]}


def _merge(machineroom=None):
    from routes.module_f.merge import merge_network
    return merge_network(_fx("_sample")(), riser=_fx("_riser")(),
                         machineroom=machineroom, mode="lsp_gravity")


# ─────────────────────────────── D1
def test_D1_기계실이_통합_좌표계로_옮겨진다():
    """★기계실 노드가 원 DXF 좌표로 남으면 통합망 bbox 가 수십 배로 뛴다."""
    mr = _machineroom()
    got = _merge(mr)
    assert got["attached"] is True
    raw = {str(n["label"]): (n["x"], n["y"]) for n in mr["nodes"]}
    at = {str(n["label"]): (float(n.get("x", 0)), float(n.get("y", 0)))
          for n in got["combined"].nodes}
    unmoved = [lab for lab in raw if lab in at and at[lab] == raw[lab]]
    assert unmoved == [], f"원 좌표에 방치된 기계실 노드: {unmoved}"


def test_D1_기계실_평면_형상이_렌더된다():
    """`plan_laid` 가 비면 기계실 평면이 화면에서 통째로 사라진다."""
    got = _merge(_machineroom())
    assert list(getattr(got["combined"], "machine_room_plan_edges", []) or [])


def test_D1_통합_bbox_가_평면도_수준으로_남는다():
    """기계실이 붙어도 캔버스 배율이 무너지지 않아야 한다."""
    a = _merge(None)["combined"]
    b = _merge(_machineroom())["combined"]

    def span(c):
        xs = [float(n.get("x", 0)) for n in c.nodes]
        ys = [float(n.get("y", 0)) for n in c.nodes]
        return max(max(xs) - min(xs), max(ys) - min(ys))

    assert span(b) < span(a) * 3, f"기계실이 bbox 를 {span(b) / span(a):.1f}배로"


def test_D1_붙는_자리는_라이저의_Input_이다():
    """★기계실 수원(m1)이 아니라 라이저 Input 이다 — 두 곳이 같은 규칙을 쓴다."""
    got = _merge(_machineroom())
    assert got["pump_junction"] == "1", got["pump_junction"]


# ─────────────────────────────── D2
def test_D2_끝점_없는_배관을_센다():
    """«?» 를 면제하는 대신 **세어서** 올린다 — 예외로 바꾸지 않는다."""
    from routes.module_f.merge import to_head_tables
    tbl = _fx("_sample")()
    tbl.pipes = list(tbl.pipes) + [
        {"label": "PX", "in": "1", "out": "?", "dia": 25, "length": 1.0}]
    ht = to_head_tables(tbl)
    assert len(ht.dangling) == 1
    assert ht.dangling[0][0] == "배관" and ht.dangling[0][3] == "?"


def test_D2_건수를_단계기록에_올린다():
    tbl = _fx("_sample")()
    tbl.pipes = list(tbl.pipes) + [
        {"label": "PX", "in": "?", "out": "?", "dia": 25, "length": 1.0}]
    from routes.module_f.merge import merge_network
    got = merge_network(tbl, riser=_fx("_riser")(), mode="lsp_gravity")
    assert any("끝점 없는 배관" in s for s in got["steps"]), got["steps"]


def test_D2_없으면_말하지_않는다():
    got = _merge(None)
    assert not any("끝점 없는" in s for s in got["steps"])


def test_D2_모르는_절점은_여전히_예외다():
    """«?» 와 달리 «옮기다 빠뜨린» 자리는 성격이 다르다 — 그대로 올린다."""
    import pytest
    from routes.module_f.merge import MergeError, to_head_tables
    tbl = _fx("_sample")()
    tbl.pipes = list(tbl.pipes) + [
        {"label": "PX", "in": "1", "out": "없는절점", "dia": 25, "length": 1.0}]
    with pytest.raises(MergeError, match="없는 절점"):
        to_head_tables(tbl)


# ─────────────────────────────── D3
def test_D3_메타_없는_절점을_모아_보고한다():
    """첫 건에서 죽이면 전체 규모를 못 본다 — 모아서 한 번에."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "design" / "tables.py").read_text(encoding="utf-8")
    assert "_missing_meta" in src
    assert "★좌표 메타 없는 절점" in src
    i = src.index("def xy(nid)")
    assert "_missing_meta.add" in src[i:i + 400], "폴백이 조용히 원점을 준다"


# ─────────────────────────────── D4
def test_D4_개명이_부속과_기기를_데려간다():
    """★개명된 배관의 부속은 emit 색인(`fittings_by_pipe`)에서 사라진다."""
    tbl = _fx("_sample")()
    # 라이저 배관 라벨(r0)과 같은 이름을 평면도에 심는다 → 평면도 쪽이 개명된다.
    tbl.pipes = [dict(p, label="r0") if i == 0 else p
                 for i, p in enumerate(tbl.pipes)]
    tbl.fittings = [{"pipe": "r0", "in": "1", "out": "2",
                     "type": "Elbow 90", "count": 1}]
    tbl.equipment = [{"pipe": "r0", "in": "1", "out": "2",
                      "label": "E1", "desc": "TEST", "eq_len": 1.0}]
    from routes.module_f.merge import merge_network
    c = merge_network(tbl, riser=_fx("_riser")(), mode="lsp_gravity")["combined"]
    labels = {str(p["label"]) for p in c.pipes}
    assert "r0_2" in labels, "개명이 안 일어났다 — 시험 전제가 깨졌다"
    orphan_f = [f for f in c.fittings if str(f.get("pipe")) not in labels]
    orphan_e = [e for e in c.equipment
                if e.get("pipe") and str(e.get("pipe")) not in labels]
    assert orphan_f == [] and orphan_e == [], (orphan_f, orphan_e)
    assert any(str(f.get("pipe")) == "r0_2" for f in c.fittings)
    assert any(str(e.get("pipe")) == "r0_2" for e in c.equipment)


def test_D4_개명이_없으면_원본_그대로다():
    c = _merge(None)["combined"]
    tbl = _fx("_sample")()
    assert [f.get("pipe") for f in c.fittings][:len(tbl.fittings)] == \
        [f.get("pipe") for f in tbl.fittings]


# ─────────────────────────────── D5
def test_D5_결합_뒤_검사가_붙는다():
    from routes.module_f.merge import combined_summary
    got = _merge(_machineroom())
    ck = got["checks"]
    assert ck["combined"] is True
    assert ck["components"] == 1, ck["component_sizes"]
    assert ck["dangling_pipes_n"] == 0
    assert ck["orphan_fittings"] == [] and ck["orphan_equipment"] == []
    assert len(ck["inputs"]) == 1, ck["inputs"]
    assert ck["bbox"] and ck["bbox"]["span_x"] > 0
    assert combined_summary(got)["checks"] is ck


def test_D5_두_기준점이_한_점이다():
    """S740 — 라이저 AV 는 헤드망 AV 자리로 snap 된다. 0 이 아니면 사선이 뜬다."""
    got = _merge(None)
    gap = got["checks"]["anchor_gap"]
    assert gap and gap.startswith("0.0 mm"), gap


def test_D5_검사는_예외가_아니다():
    """보고만 한다 — 지금 돌던 실행이 검사 때문에 실패하면 안 된다."""
    src = (_ROOT / "routes" / "module_f" / "merge.py").read_text(
        encoding="utf-8")
    i = src.index("def check_combined(")
    body = src[i:src.index("\ndef ", i + 10)]
    assert "raise" not in body, body[:200]


def test_D5_평면도_단독이면_검사가_비어_있다():
    """계통도가 없으면 결합망 자체가 없다 — 없는 것을 검사하지 않는다."""
    from routes.module_f.merge import check_combined, merge_network
    got = merge_network(_fx("_sample")(), mode="lsp_gravity")
    assert check_combined(got) == {"combined": False}
