# -*- coding: utf-8 -*-
"""[최불리규칙복원 §2-2] 관이 스쳐 지나가는 헤드는 «붙인다» — 감추지 않는다.

지시서 `ModuleF_최불리규칙_복원_지시서.md` §2-2.

■ 무엇이 문제였나

  `pass_under` = 관이 헤드 원 안을 지나가는데 그 자리에 접속 노드가 없다.
  실제 배관은 가지관에 티로 물려 있고, 도면이 그것을 «관이 원을 지나가게»
  그린 것뿐이다. 상향식은 이미 5단계에서 그 관을 헤드 중심에서 쪼개 준다
  (`stage5_split_through_uprights` · 2026-08-12 오너 승인). **하향식·상하향식
  에만 그 문이 닫혀 있었다** — 비대칭이다.

  닫혀 있는 동안 이 헤드들은 「못 붙는다」는 이유로 최불리 후보에서 빠졌고,
  하필 가장 먼 헤드가 말단 가지관 끝이라 이 모양이 되기 쉬워서 「먼 순서
  그대로 K 개」가 뒤집혔다(실측 대명동: 상위 30 중 5개가 빠졌다).

■ 조치 — 새 판정을 만들지 않는다 (§6)

  `attach_heads_center` 가 ①(중심)·②(테두리 끝)로 못 붙인 헤드를 모아
  **그 함수를 그대로 다시 부르고**, 그 결과로 ①을 다시 시도한다.

■ ★그 함수의 문을 열어야 했던 이유 — 실측

  `stage5_split_through_uprights` 는 제 안에서 `has_near_node`(중심 ≤ARM_CTR
  **또는 테두리 ±HEAD_TOUCH**)로 «이미 붙은 헤드»를 건너뛴다. 그런데
  `pass_under` 는 **정의상 테두리에 노드가 있는** 경우다(관이 스쳐 지나가며
  폴리선 꼭짓점을 남긴다). 그래서 그냥 부르면 통째로 건너뛴다::

      대명동 · 못 붙은 12개 · near=True · thru=True → **12개 전부 건너뜀 ·
      쪼갠 관 0개**   (`scripts/_probe_passunder_split.py`)

  부르는 쪽이 **이미 «못 붙었다»고 가른** 헤드만 넘길 때는 그 문을 열도록
  `assume_unattached` 를 뒀다. 기본값은 False — 5단계 상향식 경로는 한 바이트도
  안 바뀐다(그것을 아래 시험이 지킨다).
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.pipeline import flow as fw          # noqa: E402


def _src(rel):
    return (_ROOT / rel).read_text(encoding="utf-8")


# ═══════════════════════ 인공 망 — 관이 헤드 밑을 스쳐 지나간다
#
#   (0,0) ── (400,0) ──────── (1000,0)
#                  ◯ 헤드 중심 (450,0) · r=50
#
#   (400,0) 은 원 «테두리 위» 이고 차수 2 다 — 「선이 끝나는 자리」가 아니라서
#   ②가 안 열린다. 중심에는 노드가 없어 ①도 안 열린다 → `pass_under`.
HR = 50.0


def _net():
    pts = [(0.0, 0.0), (400.0, 0.0), (1000.0, 0.0)]
    edges = {(0, 1), (1, 2)}
    hcov = [(450.0, 0.0, HR)]
    return pts, edges, hcov


def test_스쳐_지나가던_헤드가_붙는다():
    pts, edges, hcov = _net()
    why: dict = {}
    p2, e2, centers, _w, _m = fw.attach_heads_center(pts, edges, hcov, why=why)
    assert centers[0] is not None, f"못 붙었다 — 사유 {why}"
    assert not why, f"붙었는데 사유가 남았다: {why}"
    # 중심에 노드가 생겼고, 그 노드가 망에 물려 있다.
    cx, cy = p2[centers[0]][0], p2[centers[0]][1]
    assert abs(cx - 450.0) < 1e-6 and abs(cy) < 1e-6, (cx, cy)
    assert any(centers[0] in e for e in e2), "중심 노드가 떠 있다"


def test_두_번_돌려도_같다():
    """★멱등 — 두 화면·두 단계가 같은 함수를 부른다. 부를 때마다 늘어나면 안 된다."""
    pts, edges, hcov = _net()
    p2, e2, c2, _w, _m = fw.attach_heads_center(pts, edges, hcov)
    p3, e3, c3, _w3, _m3 = fw.attach_heads_center(p2, e2, hcov)
    assert len(p3) == len(p2) and set(e3) == set(e2), "두 번째에 망이 늘었다"
    assert c3 == c2


def test_붙일_관이_없으면_안_붙인다():
    """★없는 배관을 지어내지 않는다 — 원 안을 지나는 관이 있을 때만이다."""
    pts = [(0.0, 0.0), (400.0, 0.0)]
    edges = {(0, 1)}
    hcov = [(5000.0, 5000.0, HR)]          # 멀리 떨어진 헤드
    why: dict = {}
    _p, _e, centers, _w, _m = fw.attach_heads_center(
        pts, edges, hcov, why=why)
    assert centers[0] is None and why.get(0) == "no_center", why


# ═══════════════════════ §6 — 그 함수를 «재사용» 한다
def test_새_판정을_안_만들었다():
    s = _src("cad_project_editor_g/services/cad_import/pipeline/flow.py")
    i = s.index("def attach_heads_center(")
    seg = s[i:s.index("\ndef ", i + 10)]
    assert "stage5_split_through_uprights(" in seg, "그 함수를 안 부른다"
    assert "assume_unattached=True" in seg


def test_5단계_기본값은_종전_그대로다():
    """★기본값이 바뀌면 상향식 5단계가 움직이고 골든이 통째로 흔들린다."""
    s = _src("cad_project_editor_g/services/cad_import/pipeline/flow.py")
    assert ("def stage5_split_through_uprights(pts, edges, ups,"
            " assume_unattached=False):") in s
    # 이미 붙은 헤드는 기본 호출에서 여전히 건너뛴다.
    pts = [(0.0, 0.0), (450.0, 0.0), (1000.0, 0.0)]   # 중심에 노드가 있다
    edges = {(0, 1), (1, 2)}
    ups = [(450.0, 0.0, HR)]
    _p, _e, n_split = fw.stage5_split_through_uprights(pts, edges, ups)
    assert n_split == 0, "이미 붙은 헤드를 또 쪼갰다"


def test_문을_열면_스쳐가는_관을_쪼갠다():
    pts, edges, hcov = _net()
    ups = [(450.0, 0.0, HR)]
    _p, _e, n0 = fw.stage5_split_through_uprights(pts, edges, ups)
    assert n0 == 0, "★테두리 노드 때문에 기본 호출은 건너뛴다(실측 그대로)"
    _p2, _e2, n1 = fw.stage5_split_through_uprights(
        pts, edges, ups, assume_unattached=True)
    assert n1 == 1, "문을 열어도 안 쪼갠다"
