# -*- coding: utf-8 -*-
"""[최불리] 전개가 못 붙이는 헤드는 «다음 순위» 로 채운다 — K 가 진짜 K 이도록.

■ 무엇이 문제였나 (2026-09-09 사용자 지적)

  「애초에 고른다는 선택지 없이, 사전에 체크한 30개의 헤드와 거기 연결된
   배관만 가져오면 되는 것 아닌가」

  옳은 지적인데 코드는 **순서가 거꾸로**였다::

      손질     최불리 K개를 고른다            ← board 도달로만 센다
      수리계산 cand = 고른 것 ∩ 붙는 헤드     ← 여기서 처음 걸러진다
               → 못 붙는 만큼 **그냥 줄었다**

  실측(대명동 골든 K=10): 고른 10개 중 2개를 전개가 배관에 붙이지 못해
  표에 **8개**만 왔다.

■ 왜 손질에서 미리 안 거르나 — 재 보고 접었다

  그 판정(`attachable_heads`)이 곧 **전체망 전개 한 번**이다. 실측(B1F ·
  절점 22,575 · 헤드 3,105)::

      손질 최불리      0.3s
      수리계산       234.2s      ← 첫 실행
      수리계산 재실행  117.3s      ← 판마다 한 번만 재는 캐시가 든 뒤

  `/edit/worst` 는 진행표시 없는 **동기** 요청이라 여기서 재면 손질 화면이
  2분 얼어붙는다. 총합은 어차피 같고 **자리만 옮긴다.** 그래서 손질은 빠르게
  두고, 이미 그 값을 재는 수리계산이 다음 순위를 채운다.

■ 조치

  손질이 «고른 K개» 와 함께 그 K개를 뽑은 **후보 범위**(`worst_cand` —
  영역·도면 장으로 가둔 것)를 실어 보낸다. 수리계산은 못 붙는 헤드를 만나면
  그 범위에서 **같은 규칙**(유하거리 긴 순서)으로 K개를 다시 채운다.
  채운 헤드는 지어낸 것이 아니라 **그다음으로 불리한 헤드**다.

  조치 뒤: 표 확정 · 헤드 **10** (종전 8). 골든 불변.
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _src(rel):
    return (_ROOT / rel).read_text(encoding="utf-8")


# ─────────────────────────────── 손질은 «범위» 도 함께 넘긴다
def test_손질이_후보_범위를_실어_보낸다():
    s = _src("routes/module_f/api_edit.py")
    assert 'sess["worst_cand"]' in s, "후보 범위를 안 남긴다"
    i = s.index('sess["worst_cand"]')
    assert "「도면 전체가 후보」" in s[max(0, i - 300):i + 200]


def test_손질은_전개를_돌리지_않는다():
    """★손질에서 재면 큰 도면에서 화면이 2분 얼어붙는다(B1F 실측 117초).

    `/edit/worst` 는 진행표시 없는 동기 요청이다. 여기에 전체망 전개를
    들이면 «빠른 화면» 이라는 성질 자체가 사라진다 — 총합은 어차피 같고
    자리만 옮기는 것이라 이득도 없다.
    """
    s = _src("routes/module_f/api_edit.py")
    assert "wet_heads(" not in s, "손질이 전개 탐침을 돌린다"
    assert "attachable_heads" not in s
    # 왜 안 하는지가 적혀 있어야 다음 사람이 다시 넣지 않는다.
    assert "117초" in s and "동기 요청" in s


# ─────────────────────────────── 수리계산이 채운다
def test_수리계산이_다음_순위로_채운다():
    s = _src("routes/module_f/api_design.py")
    i = s.index("filled = 0")
    seg = s[i:i + 1200]
    assert "wet_heads(" in s[max(0, i - 900):i], "붙는 헤드를 안 잰다"
    assert 'sess.get("worst_cand")' in seg, "후보 범위를 안 쓴다"
    assert "short < k_use" in seg, "모자랄 때만 채우는 조건이 없다"
    assert "다음 순위" in seg, "무엇을 했는지 안 말한다"


def test_모자라지_않으면_손질_선정을_그대로_쓴다():
    """★다 붙으면 손질이 고른 그 K개다 — 멀쩡한 것을 다시 뽑지 않는다."""
    s = _src("routes/module_f/api_design.py")
    i = s.index("filled = 0")
    seg = s[i:i + 1200]
    # 채우는 갈래는 `short < k_use` 안에서만 `only` 를 바꾼다.
    #
    # ★창을 400 → 900 으로 넓혔다(2026-09-10). 그 사이에 §2-5 의 「채우지
    #   않기」 갈래가 들어와 확인할 줄이 창 밖으로 밀렸다. 지키려는 것은
    #   «조건 안에서만 바꾼다» 이지 «몇 자 안에 있다» 가 아니다 —
    #   창 크기가 시험의 참·거짓을 가르면 안 된다(같은 이유로 인계 시험도
    #   한 번 넓혔다).
    j = seg.index("if wet and short < k_use:")
    assert "only = (set(pool) & wet) if pool else None" in seg[j:j + 900]
    assert seg[:j].count("only =") == 0, "조건 밖에서 후보를 갈아치운다"


# ─────────────────────────────── 두 화면이 같은 자를 쓴다
def test_제한_전개도_head_xy_를_넘긴다():
    """★없으면 «같은 자리» 판정이 좌표가 아니라 절점으로 떨어져 손질과 다른
    K개를 고른다 — 두 화면이 다른 말을 하게 된다."""
    s = _src("cad_project_editor_g/services/cad_import/design/restrict.py")
    i = s.index("worst = worst_k_heads(")
    seg = s[i:i + 400]
    assert 'head_xy=getattr(b, "disks", None)' in seg


# ─────────────────────────────── 탐침은 판마다 한 번
def test_탐침을_판마다_한_번만_잰다():
    s = _src("routes/module_f/attach.py")
    assert "def board_stamp(" in s and "def wet_heads(" in s
    assert '_wet_probe' in s
    # 지문에 판을 바꾸는 손잡이가 다 들어 있어야 한다.
    for name in ("sources", "valves", "disk_kinds", "joins", "deletes",
                 "edge_len_mm"):
        assert name in s, name


def test_탐침_결과를_전개가_다시_안_잰다():
    s = _src("cad_project_editor_g/services/cad_import/design/restrict.py")
    assert "probe=None" in s
    i = s.index("if probe is None:")
    assert "가장 비싼 줄" in s[max(0, i - 500):i]


def test_캐시가_빗나가면_말한다():
    """★조용히 빗나가면 비용만 늘고 아무도 모른다 — 실제로 한 번 그랬다."""
    s = _src("routes/module_f/attach.py")
    assert "판이 그대로라 다시 재지 않습니다" in s
    assert "전체망 전개를 한 번 돕니다" in s and "판이 바뀜" in s


def test_지문이_판_내용까지_본다():
    """개수만 보면 «지우고 같은 수만큼 그린» 판을 같다고 읽는다."""
    from routes.module_f.attach import board_stamp

    class _B:
        pts = [(0.0, 0.0), (1.0, 1.0)]
        edges = [(0, 1)]
        disks = [(1.0, 1.0, 42.0)]
        sources = [0]
        valves = [0]
        disk_kinds = ["하향식"]
        joins = []
        deletes = []
        edge_len_mm = {}

    class _ES:
        key = "k"
        board = _B()

    es = _ES()
    a = board_stamp(es)
    es.board.disks = [(9.0, 9.0, 42.0)]      # 개수는 같고 «자리» 만 다르다
    assert board_stamp(es) != a
