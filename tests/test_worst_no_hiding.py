# -*- coding: utf-8 -*-
"""[최불리규칙복원 §2-3·§6] 못 붙는 헤드를 «감추지» 않는다 — 막고 짚는다.

지시서 `ModuleF_최불리규칙_복원_지시서.md`.

  규칙을 깨던 세 가지는 전부 「못 붙는 헤드」를 **감추려는** 시도였다:

      ① 후보 깎기   붙는 것만 후보로 → 가장 먼 헤드가 안 뽑힌다  (§2-1 철회)
      ② 백필        뽑은 뒤 다른 헤드로 바꿔 넣기 → 두 화면이 갈린다 (§6 금지)
      ③ 되먹임      수리계산 결과를 평면으로 되돌리기 → 정의가 사라진다 (§6 금지)

  감추지 말고 **붙이거나(§2-2), 멈추고 짚는다(§2-3).**
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


def _between(s, a, b, start=0):
    i = s.index(a, start)
    return s[i:s.index(b, i)]


# ═══════════════════════ 기준 1·3 — 후보는 영역·장 제한만 받는다
def test_후보는_영역과_장으로만_좁힌다():
    """★사람이 손질에서 이은 헤드는 **반드시** 후보에 든다(기준 3).

    그것이 참이려면 후보를 정하는 자리에 «영역·장» 말고 다른 자가 없어야
    한다. 여기서 `only` 에 쓰는 줄을 전부 세어 확인한다.
    """
    s = _src("routes/module_f/api_edit.py")
    body = _between(s, "def _compute_worst", "w = _worst_k_heads(")
    code = [ln for ln in body.splitlines()
            if not ln.lstrip().startswith("#")]
    # `only` 가 **왼쪽에** 오는 줄만 — 튜플 대입(`only, sheet_no = …`)도 잡는다.
    #   `only_heads=only` 는 `only\b` 뒤가 밑줄이라 안 걸린다.
    writes = [ln.strip() for ln in code
              if re.match(r"\s*only\b[^=]*=[^=]", ln)]
    # 허용되는 것: 초기화 · 도면 장 · 영역 — 이 셋뿐이다.
    assert writes == [
        "only, sheet_no = None, None",
        "only = {hi for hi, d in enumerate(b.disks)",
        "only = in_zone if only is None else (only & in_zone)",
    ], writes


# ═══════════════════════ 기준 5 — K 안에 못 붙는 헤드가 있으면 막는다
def test_뽑힌_K_에_못_붙는_헤드가_있으면_막는다():
    s = _src("routes/module_f/api_edit.py")
    seg = _between(s, "wet_ok = {int(i) for i in probe[\"wet\"]}",
                   "return {\"k\": len(w[\"heads\"])")
    assert "blocked = [int(h) for h in w[\"heads\"] if int(h) not in wet_ok]" in seg
    assert "if blocked:" in seg
    # 막을 때 선정을 **지운다** — 안 지우면 옛 선정으로 표가 선다.
    assert 'sess["worst"] = None' in seg
    # 채우지도 빼지도 않는다.
    assert "zone_confined_pool" not in seg and "keep" not in seg
    # 자리·사유·할 일을 함께 낸다.
    for k in ('"why"', '"todo"', '"xy"', '"not_attached"'):
        assert k in seg, k


def test_사유별_할_일이_한_벌이다():
    """★문구가 두 벌이면 한쪽만 고쳐지는 날이 온다 — 서버가 주고 화면은 그린다."""
    s = _src("routes/module_f/api_edit.py")
    i = s.index("ATTACH_TODO = {")
    seg = s[i:s.index("}", i)]
    for why in ("dry", "center_dry", "no_center", "chord_only",
                "pass_under", "shared"):
        assert f'"{why}"' in seg, why
    js = _src("static/module_f.js")
    assert "function renderBlocked(" in js
    assert "renderBlocked(err.data && err.data.not_attached)" in js
    # 화면은 서버가 준 `todo` 를 그대로 쓴다(제 문구를 만들지 않는다).
    box = _between(js, "function renderBlocked(", "function renderRankBroken(")
    assert "r.todo" in box
    assert 'id="ed-blocked"' in _src("templates/module_f.html")


def test_오류에_자료가_실려_온다():
    """★문장만 던지면 «어느 헤드가·왜·뭘 하면 되는지» 가 버려진다."""
    js = _src("static/module_f.js")
    seg = _between(js, "async function api(", "const post =")
    assert "e.data = d;" in seg


# ═══════════════════════ 기준 6 — 백필은 끈다. 지우지는 않는다.
def test_백필이_꺼져_있다():
    s = _src("routes/module_f/api_design.py")
    assert "BACKFILL_DISABLED = True" in s
    seg = _between(s, "if wet and short < k_use:", "got = select_and_expand(")
    assert "if BACKFILL_DISABLED:" in seg
    # 발동하면 **그 자체가 오류** 다 — 조용히 메우지 않는다.
    i = seg.index("if BACKFILL_DISABLED:")
    assert '"ok": False' in seg[i:i + 700]
    assert "[★비정상]" in seg


def test_백필_코드를_지우지_않았다():
    """§6 — 「비활성으로 두되 지우지 말 것」. 왜 껐는지 모르면 다시 켠다."""
    s = _src("routes/module_f/api_design.py")
    seg = _between(s, "if wet and short < k_use:", "got = select_and_expand(")
    assert "zone_confined_pool(" in seg, "백필 블록을 지웠다"
    assert "only = (set(pool) & wet) if pool else None" in seg


# ═══════════════════════ 기준 8 — 되먹임이 없다
def test_되먹임이_없다():
    s = _src("routes/module_f/api_design.py")
    assert not re.findall(r'sess\[\s*["\']worst["\']\s*\]\s*=', s)
    assert "_adopt_final_worst" not in s
    for rel in ("routes/module_f/remote30.py", "routes/module_f/views.py"):
        t = _src(rel)
        for bad in ("_design_corridor", "net_from", "from_design"):
            assert bad not in t, f"{rel}: {bad}"
    js = _src("static/module_f.js")
    assert "net_from" not in js and "from_design" not in js
