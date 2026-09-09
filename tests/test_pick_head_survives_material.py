# -*- coding: utf-8 -*-
"""[찍기] 배관을 «하나 더» 찍어도 찍은 헤드가 살아남는다.

■ 무엇이 문제였나 (대명동 · 사용자 지목으로 드러남)

  「02.찍기에서 수동으로 정의한 배관이랑 헤드를 04.수리계산에서 반영을 못한다」

  원인은 한 줄이었다::

      board.py  _click_line()
          cleared = self.clear_heads() if self.heads else 0

  배관 클릭 **한 번**이 찍은 헤드를 통째로 지웠다. 엔진은 그 수를 보고로
  돌려줬지만(`헤드해제`) `static/module_f.js` 에 그 이름은 **0 건**이었다.
  그래서 화면은 아무 말도 하지 않았다.

  실측(대명동 · 추천 채택 뒤 배관 한 묶음 추가)::

      스펙 헤드 칸        1  →  **0**
      손질판 헤드      111  →  **0**        ← 수리계산에 노즐이 하나도 안 온다

■ 왜 지웠던 것은 옳았나 — 그래서 «지우기» 가 아니라 «다시 찍기» 다

  헤드 판정은 문양 지문(`pick_mark_fp`)을 쓰고 그 지문은 **재료 집합에
  딸려 있다**. 재료가 바뀌면 옛 지문은 그 도면의 답이 아니다. 그러니 버리는
  대신 **같은 자리를 새 재료로 다시 찍는다** — `pick/auto` 가 이미 쓰는
  관용구다(「묶음의 실제 선분 중점으로 정상 클릭 경로를 태운다」).

  조치 뒤 실측: 헤드 칸 1 → **1** · 손질판 헤드 0 → **24**
  (24 는 사람이 보탠 그 재료가 문양 판정을 바꾼 결과다 — 찍은 대로다.)
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


class _W:
    """가장 작은 세계 — 배관 두 묶음과 헤드 원 하나."""

    def __init__(self):
        self.segs = [
            ("PIPE", 1, (0.0, 0.0), (1000.0, 0.0)),
            ("PIPE", 1, (1000.0, 0.0), (2000.0, 0.0)),
            ("PIPE2", 3, (0.0, 500.0), (1000.0, 500.0)),
        ]
        self.circles = [("HEAD", 2, 1000.0, 0.0, 40.0)]
        self.arcs = []
        self.arc_ang = []
        self.texts = []


def _board():
    from services.cad_import.pick.board import Board
    kn = {"small_r": 100.0, "small_len": 300.0}
    return Board(_W(), kn)


def _pick_pipe_then_head(b):
    b.apply_click("재료", 500.0, 0.0)          # PIPE×1
    b.complete_materials()
    b.apply_click("헤드", 1040.0, 0.0)         # 원 테두리
    return b


# ─────────────────────────────── ★핵심 — 배관을 더 찍어도 헤드가 남는다
def test_배관을_더_찍어도_헤드가_남는다():
    b = _pick_pipe_then_head(_board())
    assert b.heads, "표본이 헤드를 못 찍었다 — 시험 전제가 깨졌다"
    n0 = len(b.heads)
    rep = b.apply_click("재료", 500.0, 500.0)   # PIPE2×3 을 «더» 찍는다
    assert rep is not None and rep["동작"] == "추가"
    assert len(b.mat) == 2, b.mat
    assert len(b.heads) == n0, f"헤드가 사라졌다 {n0} → {len(b.heads)}"


def test_되살린_수를_보고에_담는다():
    """조용히 하지 않는다 — 화면이 읽을 수 있어야 말할 수 있다."""
    b = _pick_pipe_then_head(_board())
    rep = b.apply_click("재료", 500.0, 500.0)
    assert rep["헤드되살림"] == len(b.heads)
    assert rep["헤드잃음"] == 0


def test_화면이_잃은_헤드를_말한다():
    """★보고에 담기만 하고 아무도 안 읽으면 종전과 똑같다.

    종전 `헤드해제` 가 정확히 그랬다 — 엔진은 세어 보냈고 화면은 그 이름을
    한 번도 쓰지 않았다.
    """
    js = (_ROOT / "static" / "module_f.js").read_text(encoding="utf-8")
    i = js.index("async function pickClick(")
    seg = js[i:i + 2000]
    assert '헤드잃음' in seg, "화면이 잃은 헤드를 안 읽는다"
    assert "헤드 칸을 다시 찍어" in seg, "무엇을 하면 되는지 안 말한다"


def test_재료를_취소해도_헤드가_남는다():
    """켜기만이 아니라 «끄기» 도 재료 변경이다 — 같은 길로 간다."""
    b = _board()
    b.apply_click("재료", 500.0, 0.0)
    b.apply_click("재료", 500.0, 500.0)
    b.complete_materials()
    b.apply_click("헤드", 1040.0, 0.0)
    n0 = len(b.heads)
    assert n0
    b.apply_click("재료", 500.0, 500.0)          # 취소
    assert len(b.mat) == 1
    assert len(b.heads) == n0


def test_클릭_기록이_어긋나지_않는다():
    """되살림은 헤드 클릭 기록을 다시 쌓는다 — 되돌리기가 성립해야 한다."""
    b = _pick_pipe_then_head(_board())
    b.apply_click("재료", 500.0, 500.0)
    heads_cl = [c for c in b.clicks if c["모드"] == "헤드"]
    assert len(heads_cl) == 1, [c["모드"] for c in b.clicks]
    # 마지막 클릭(되살린 헤드)을 되돌리면 헤드가 빠진다.
    b.undo()
    assert not b.heads


def test_지우기로_되돌아가지_않는다():
    """★`clear_heads()` 한 줄로 돌아가면 이 사고가 그대로 재발한다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "pick" / "board.py").read_text(encoding="utf-8")
    i = src.index("    def _click_line(")
    seg = src[i:i + 1500]
    assert "_replay_head_clicks(" in seg, "다시 태우지 않는다"
    j = src.index("def _replay_head_clicks(")
    assert "재료 집합에 딸려" in src[j:j + 1800], "왜 다시 찍는지를 안 적었다"


def test_손질이_헤드_0_을_말한다():
    """헤드 0 으로 넘어가면 수리계산 표가 빈다 — 그 자리에서 말한다."""
    src = (_ROOT / "routes" / "module_f" / "api_pick.py").read_text(
        encoding="utf-8")
    i = src.index("if not es.board.disks:")
    assert "노즐이 하나도 오지 않습니다" in src[i:i + 400]


def test_잃은_칸을_부풀려_세지_않는다():
    """★«칸» 으로 센다 — 클릭 기록의 픽 개수를 그냥 더하면 안 된다.

    같은 헤드 키가 여러 기록에 들어 있어(한 자리를 두 번 찍거나, 추가 뒤
    취소하거나) 실서버 실측에서 195 로 부풀었고, 실제로 잃은 것이 하나도
    없는데 화면이 「192칸을 잃었습니다」라고 말할 뻔했다. 틀린 수를 말하면
    사람을 엉뚱한 데로 보낸다.
    """
    b = _pick_pipe_then_head(_board())
    b.apply_click("헤드", 1040.0, 0.0)          # 같은 자리를 다시 → 취소
    b.apply_click("헤드", 1040.0, 0.0)          # 다시 → 추가
    assert len([c for c in b.clicks if c["모드"] == "헤드"]) == 3
    rep = b.apply_click("재료", 500.0, 500.0)
    assert rep["헤드잃음"] == 0, rep
    assert rep["헤드되살림"] == len(b.heads)


def test_되살림은_키_기준이다():
    """중복을 걷어 내는 자리가 한 곳이어야 두 수가 안 갈린다."""
    src = (_ROOT / "cad_project_editor_g" / "services" / "cad_import"
           / "pick" / "board.py").read_text(encoding="utf-8")
    i = src.index("def _replay_head_clicks(")
    seg = src[i:i + 2600]
    assert "want = set()" in seg and "head_key(" in seg
    assert "len(want - now)" in seg, "차집합이 아니라 뺄셈이면 또 부풀 수 있다"
