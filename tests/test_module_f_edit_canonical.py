# -*- coding: utf-8 -*-
"""[손질정본] 손질이 정본이고 수리계산은 받아 쓰기만 한다.

지시서 `ModuleF_손질정본_지시서.md`.

■ 계약 (§0)

    ① 프로그램이 배관망을 자동으로 정의한다
    ② 사람이 손질에서 더 정의한다
    ③ 그 위에서 알람밸브·영역으로 기준개수 K개를 뽑는다
    ④ 수리계산은 ③이 정한 것을 **그대로** 받는다

  정본은 ②·③ 의 결과다. 수리계산은 스스로 다시 고르지 않는다.

■ 왜 갈렸나 (§1)

  「헤드가 배관에 붙었나」를 두 함수가 다른 자로 잰다 — 최불리는 느슨한 자
  (`b.hnodes`)를 쓰고 전개는 엄격한 자를 쓴다. 그래서 손질 선정에 «전개가 못
  붙이는 헤드» 가 섞이고, 수리계산이 그것을 걸러 내며 다른 헤드로 채웠다.
  **거르는 자리가 거꾸로였다.**

■ ★조치 (§2-1) 는 **철회**됐다 (2026-09-13)

  「거르기를 손질로 옮긴다」가 이 문서의 §2-1 이었고, 그것이 규칙을 깼다 —
  후보를 깎으면 「먼 순서 그대로 K 개」가 성립하지 않는다. 철회했다.
  대체 문서 `ModuleF_최불리규칙_복원_지시서.md`.

  **남은 것만 이 파일이 지킨다** — 되먹임 없음(기준 6·7)은 그대로 유효하다.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
for _p in (str(_ROOT), str(_ROOT / "core"),
           str(_ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)


def _src(rel):
    return (_ROOT / rel).read_text(encoding="utf-8")


# ═══════════════════════ §0 · 기준 6 — 되먹임이 없다
def test_design_build_이_선정을_안_쓴다():
    """★수리계산 결과를 손질로 되돌리면 «사람이 정의한 것» 이 사라진다.

    전에 이 방향으로 고쳤다가 되돌린 적이 있다(§0). 코드로 막아 둔다.
    """
    s = _src("routes/module_f/api_design.py")
    # sess["worst"] 에 **쓰는** 자리가 없어야 한다. 읽기(get)는 허용.
    writes = re.findall(r'sess\[\s*["\']worst["\']\s*\]\s*=', s)
    assert not writes, f"/design/build 가 선정을 덮어쓴다: {writes}"
    assert "_adopt_final_worst" not in s, "되먹임 함수가 남아 있다"


def test_되먹임_함수가_저장소에_없다():
    for rel in ("routes/module_f/selection.py", "routes/module_f/remote30.py",
                "routes/module_f/views.py"):
        assert "_adopt_final_worst" not in _src(rel), rel


# ═══════════════════════ 기준 7 — 평면 보기 코드 무변경
def test_worst_view_가_표를_안_본다():
    """★평면이 수리계산을 따라가면 ②가 의미를 잃는다(§0)."""
    s = _src("routes/module_f/remote30.py")
    for bad in ("_design_corridor", "net_from", "from_design",
                'sess.get("design")', "tables"):
        assert bad not in s, f"remote30 이 설계 표를 본다: {bad}"


def test_views_가_표를_안_본다():
    s = _src("routes/module_f/views.py")
    for bad in ("design", "net_from", "from_design"):
        assert bad not in s, f"views.py 가 설계 표를 본다: {bad}"


def test_평면_JS_가_표를_안_본다():
    js = _src("static/module_f.js")
    assert "net_from" not in js and "from_design" not in js
    # 표 확정 뒤 손질 상태를 다시 받아 «표 기준» 으로 갈아 끼우지 않는다.
    i = js.index('const d = await post("/api/module-f/design/build"')
    seg = js[i:i + 1400]
    assert "/api/module-f/edit/state" not in seg, "표 확정이 평면을 갈아 끼운다"


# ═══════════════════════ ★§2-1 은 폐기됐다 — 후보를 깎지 않는다
#
#   여기 있던 `test_손질이_노즐로_오는_헤드만_후보로_삼는다` 를 걷어냈다.
#   그 시험이 지키던 조치(`only = before & keep`)가 **지금 증상의 원인**이었다:
#   `worst_k_heads` 는 `only_heads` 밖을 순위 계산에서 아예 건너뛰므로,
#   「못 붙는다」고 뺀 헤드가 2등이면 32등이 그 자리에 온다.
#
#       실측(대명동 K=30 · `scripts/_probe_candidate_drop.py`):
#         빠진 33개 중 **상위 30 안에 5개** — 2등 51.69m · 5등 49.49m ·
#         19등 · 24등 · 26등. 전부 `pass_under`.
#
#   대체 문서: `ModuleF_최불리규칙_복원_지시서.md`.
#   규칙을 지키는 자리는 `tests/test_worst_rank_invariant.py` (불변식 ①②③) ·
#   `tests/test_worst_no_hiding.py` (막기·백필 금지) 로 옮겼다.


def test_shared_를_판정하지_않고_들고_온다():
    """★§6 판정 규칙 변경 금지 — 이미 있는 `shared_head_idx` 를 옮길 뿐이다."""
    s = _src("cad_project_editor_g/services/cad_import/design/restrict.py")
    assert '"shared": set(built.get("shared_head_idx") or ())' in s
    # 실패 경로도 같은 모양이어야 부르는 쪽이 갈래를 안 만든다.
    assert '"shared": set()' in s


def test_판정_규칙은_한_글자도_안_바꿨다():
    """§6 — 이 지시서는 «어느 자를 쓰느냐» 를 바꾸지 자를 깎지 않는다."""
    s = _src("cad_project_editor_g/services/cad_import/pipeline/flow.py")
    assert "if d <= ARM_CTR or abs(d - hr) <= tol:" in s          # head_nodes
    assert "elif abs(d - hr) <= tol and len(nbr[n]) == 1:" in s   # attach
    assert "if abs(do - hr) <= 2.0:" in s                         # 현 걸름


# ═══════════════════════ ★§2-2(뺀 헤드 배너)는 퇴역 — 2026-09-14 리팩터링
#
#   여기 있던 `test_뺀_헤드를_응답에_싣는다` · `test_화면이_사유별로_할_일을_말한다`
#   는 걷어냈다. 속도 조치(고른 뒤 K개만 전개)로 서버의 `not_attachable` 이
#   **영구히 빈 dict** 가 되어, 그 응답 필드·배너(renderNotAttachable)는 한 번도
#   뜰 수 없는 코드였다 — 항상 0 인 통계는 없는 것보다 나쁘다(있다고 믿게 한다).
#   같은 의무는 산 경로가 진다:
#     · 뽑힌 K 안에 못 붙는 헤드 → §2-3 이 400 의 `not_attached` 로 자리·사유·
#       할 일을 싣고 화면(renderBlocked)이 그린다
#       — tests/test_worst_no_hiding.py 가 지킨다.
#     · 도면 전체의 «안 붙는 헤드» 통계 → 수리계산의 «제외 사유».
# ═══════════════════════ §2-3 · 기준 3 — K 미달은 손질에서 막는다
def test_K_미달_게이트는_그대로다():
    """후보가 좁아지면 `reachable` 이 저절로 «표에 오는 헤드 수» 가 된다."""
    s = _src("routes/module_f/api_edit.py")
    assert 'if w["reachable"] < k:' in s, "게이트가 사라졌다"
    i = s.index('if w["reachable"] < k:')
    assert "설계면적이 성립하지 않습니다" in s[i:i + 900]


# ═══════════════════════ §2-4 — 백필은 안전망으로만 남는다
def test_백필은_남기되_조용하지_않다():
    s = _src("routes/module_f/api_design.py")
    i = s.index("if wet and short < k_use:")
    seg = s[i:i + 900]
    assert "[★비정상]" in seg, "정상 흐름이 아님을 안 말한다"
    assert "zone_confined_pool(" in s, "안전망을 지웠다"
