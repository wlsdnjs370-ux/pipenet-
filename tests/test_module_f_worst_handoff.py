# -*- coding: utf-8 -*-
"""[최불리 인계] 손질이 고른 K개를 수리계산이 그대로 받는다.

지시서 `ModuleF_최불리선정_인계_지시서.md`.

■ 증상 — 손질의 선정을 수리계산이 버렸다

      routes/module_f/api_design.py
          got = select_and_expand(payload, es.board, k=cfg["k"],
                                  selected_source=sel)     # ★only_heads 없음

  `select_and_expand` 는 `only_heads=None` 이면 `cand = wet`(도면 전체의 물닿는
  헤드)에서 K 개를 **다시 뽑는다.** 손질 코드는 K 를 이미 맞춰 주고 있었다
  (`sess["worst_k"]`) — **K 는 맞췄는데 선정 결과를 빠뜨린 것**이다.

■ 실측 (대명동 · 영역 안 헤드 16 · K=12 · `scripts/_probe_worst_handoff.py`)

                          조치 전            인계 후
      설계가 본 후보       111 / 도면 111  →  **12** / 도면 111
      손질↔설계 헤드 대응   **0/12**        →  **11/12** (중앙 오차 14.0 mm)
      기준 헤드 어긋남      7,256 mm        →  **14.0 mm**
      최원 유하거리         42.79 vs 57.58  →  **42.79 = 42.79**

  ★12개 중 1개는 표에 못 왔다 — 물닿음 판정은 통과했는데 제한 전개에서 떨어진
    자리다(260307.6, −228066.2). **다른 헤드로 채우지 않고 말한다.**

■ 인덱스 공간 (§2-2 — 이것부터 확인했다)

      손질  w["heads"]      = {hi for hi, d in enumerate(b.disks) …}
      payload  data["hcov"] = [list(d) for d in b.disks]
      planar   wet_head_idx = enumerate(hcov) 의 인덱스

  셋 다 `board.disks` 순서다. 조치 후 11/12 가 좌표로 1:1 맞은 것이 그 실증이다.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _src(name="api_design.py"):
    return (_ROOT / "routes" / "module_f" / name).read_text(encoding="utf-8")


# ─────────────────────────────── ① 손질 선정을 넘긴다
def test_손질_선정을_only_heads_로_넘긴다():
    """★이 한 줄이 없어서 수리계산이 도면 전체에서 다시 뽑았다."""
    src = _src()
    i = src.index("got = select_and_expand(")
    seg = src[max(0, i - 1400):i + 300]
    assert "only_heads=only" in seg, seg[-400:]
    assert 'sess.get("worst")' in seg, "손질 선정을 안 읽는다"
    assert 'w_sel.get("heads")' in seg


def test_선정이_없으면_전체에서_뽑되_말한다():
    """§2-4 — 조용히 다르게 동작하는 갈래를 두지 않는다."""
    from routes.module_f.api_design import _worst_handoff_note
    got = {"candidate_heads": 30}
    note = _worst_handoff_note(got, [], 30, 30)
    assert note["from_edit"] is False
    assert any("도면 전체에서" in m for m in note["messages"]), note


def test_선정이_있으면_그_사실을_말한다():
    from routes.module_f.api_design import _worst_handoff_note
    note = _worst_handoff_note({"candidate_heads": 12}, list(range(12)), 12, 12)
    assert note["from_edit"] is True and note["picked"] == 12
    assert not note["messages"], note          # 다 받았으면 할 말이 없다


# ─────────────────────────────── ② K 를 조용히 깎지 않는다
def test_K_가_어긋나면_손질_것을_믿는다():
    """§2-5 — 기준개수의 입력칸은 손질에 하나뿐이라는 것이 이 저장소의 결정이다."""
    src = _src()
    i = src.index("k_use = int(cfg[\"k\"])")
    seg = src[i:i + 400]
    assert 'sess.get("worst_k")' in seg, seg
    from routes.module_f.api_design import _worst_handoff_note
    note = _worst_handoff_note({"candidate_heads": 12}, list(range(12)), 12, 30)
    assert any("손질 값 12" in m for m in note["messages"]), note


def test_전개가_못_붙이면_세어서_말한다():
    """§2-3 — 개수만 세면 사람은 어디를 고쳐야 할지 모른다."""
    from routes.module_f.api_design import _worst_handoff_note
    note = _worst_handoff_note({"candidate_heads": 9}, list(range(12)), 12, 12)
    assert note["not_attachable"] == 3
    assert any("붙이지 못했습니다" in m for m in note["messages"])
    assert any("다른 헤드로 채우지 않았습니다" in m for m in note["messages"])


# ─────────────────────────────── ③ 표를 보고 다시 센다
class _T:
    def __init__(self, nodes, nozzles):
        self.nodes = nodes
        self.nozzles = nozzles


class _B:
    def __init__(self, disks):
        self.disks = disks


def test_표에_못_온_헤드를_좌표로_짚는다():
    """★물닿음 판정을 통과하고도 제한 전개에서 떨어질 수 있다(실측 12 → 11)."""
    from routes.module_f.api_design import _handoff_after_table
    disks = [(1000.0, 1000.0, 40.0), (2000.0, 1000.0, 40.0),
             (3000.0, 1000.0, 40.0)]
    # 표에는 앞의 둘만 왔다 — 세 번째가 빠졌다.
    nodes = [{"label": "1", "x": 1000.0, "y": 1000.0},
             {"label": "2", "x": 2000.0, "y": 1000.0}]
    tbl = _T(nodes, [{"in": "1"}, {"in": "2"}])
    got = {"handoff": {"from_edit": True, "picked": 3, "messages": []},
           "_picked": [0, 1, 2], "origin_mm": (1000.0, 1000.0)}
    _handoff_after_table(got, tbl, _B(disks))
    h = got["handoff"]
    assert h["missing"] == 1 and h["in_table"] == 2, h
    assert [m["disk"] for m in h["missing_heads"]] == [2], h["missing_heads"]
    assert any("표에 오지 못했습니다" in m for m in h["messages"])


def test_짝짓기는_1대1이다():
    """★「각자 가장 가까운 것」만 보면 두 헤드가 같은 노즐을 짚어도 통과한다.

    실측에서 그렇게 한 번 놓쳤다 — 표는 11개인데 «다 찾음» 이 나왔다.
    """
    from routes.module_f.api_design import _handoff_after_table
    #  두 헤드가 거의 같은 자리인데 표에는 노즐이 하나뿐이다.
    disks = [(1000.0, 1000.0, 40.0), (1000.0, 1010.0, 40.0)]
    tbl = _T([{"label": "1", "x": 1000.0, "y": 1000.0}], [{"in": "1"}])
    got = {"handoff": {"from_edit": True, "picked": 2, "messages": []},
           "_picked": [0, 1], "origin_mm": (1000.0, 1000.0)}
    _handoff_after_table(got, tbl, _B(disks))
    assert got["handoff"]["missing"] == 1, got["handoff"]
    assert len(got["handoff"]["missing_heads"]) == 1


def test_다_왔으면_아무_말도_안_한다():
    from routes.module_f.api_design import _handoff_after_table
    disks = [(1000.0, 1000.0, 40.0)]
    tbl = _T([{"label": "1", "x": 1000.0, "y": 1000.0}], [{"in": "1"}])
    got = {"handoff": {"from_edit": True, "picked": 1, "messages": []},
           "_picked": [0], "origin_mm": (1000.0, 1000.0)}
    _handoff_after_table(got, tbl, _B(disks))
    assert got["handoff"]["missing"] == 0
    assert not got["handoff"]["messages"]


def test_선정이_없으면_표를_안_본다():
    """손질에서 안 골랐으면 «못 온 헤드» 라는 개념 자체가 없다."""
    from routes.module_f.api_design import _handoff_after_table
    got = {"handoff": {"from_edit": False, "messages": []}}
    _handoff_after_table(got, _T([], []), _B([]))
    assert "missing" not in got["handoff"]


# ─────────────────────────────── ④ 화면과 금지 사항
def test_화면이_인계를_말한다():
    js = (_ROOT / "static" / "module_f.js").read_text(encoding="utf-8")
    i = js.index("function handoffLines(")
    seg = js[i:i + 1600]
    assert "손질에서 고른" in seg and "표에 못 온 헤드" in seg, seg[:400]
    assert "다른 헤드로 채우지 않았습니다" in seg
    assert "handoffLines(s.handoff)" in js, "요약이 그것을 안 부른다"


def test_손질_화면은_안_건드린다():
    """§5 — 손질은 이미 옳게 동작한다. 넘기는 인자 하나만 더한다."""
    import subprocess
    r = subprocess.run(
        ["git", "diff", "--stat", "HEAD", "--",
         "routes/module_f/api_edit.py", "routes/module_f/remote30.py"],
        cwd=str(_ROOT), capture_output=True, encoding="utf-8",
        errors="replace")
    assert not (r.stdout or "").strip(), r.stdout


def test_선정_규칙은_안_바꿨다():
    """§5 — `select_and_expand`·`worst_k_heads` 의 규칙은 그대로다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "design" / "restrict.py").read_text(encoding="utf-8")
    assert "cand = wet if only_heads is None else (set(only_heads) & wet)" in src
