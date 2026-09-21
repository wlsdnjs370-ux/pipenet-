# -*- coding: utf-8 -*-
"""손질 상태가 «저장 → 재열기» 왕복을 견디는가 — 사람의 결정이 남는지.

이 저장소에서 되풀이된 결함 유형이 「찍었는데 산출·재열기에서 사라진다」였다
(헤드 종류 · 알람밸브 기기 · 접속점). 그래서 왕복 자체를 한 판에 잰다.
"""
from __future__ import annotations

import os
import sys
import tempfile

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
for _p in (_ROOT, _G):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from services.cad_import.edit.board import EditBoard          # noqa: E402
from services.cad_import.edit.io import load_edits, write_edits  # noqa: E402
from services.cad_import.edit.session import (                # noqa: E402
    EditSession, MODE_VALVE)
from services.cad_import.kinds import (                       # noqa: E402
    disk_kind_list, require_head_kinds)

bad: list[str] = []


def chk(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '★어긋남'}] {name}"
          + (f" · {detail}" if detail else ""))
    if not ok:
        bad.append(f"{name} — {detail}")


def _board(key):
    return EditBoard(key, [(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)],
                     [(0, 1), (1, 2)],
                     [(0.0, 0.0, 150.0), (6000.0, 0.0, 150.0)],
                     head_kinds=[])


def _eng(b, xy):
    rec = require_head_kinds(b.disks, b.head_kinds)
    return {tuple(r["c"]): r["kind"] for r in rec}.get(xy)


with tempfile.TemporaryDirectory() as td:
    print("[1] 헤드 종류 — 찍은 뒤 왕복")
    b1 = _board("왕복")
    b1.set_head_kind(b1.disks[0], "하향식")
    b1.set_head_kind(b1.disks[1], "상하향식")
    write_edits(b1, td)
    b2 = _board("왕복")
    load_edits(b2, td)
    chk("첫 헤드", _eng(b2, (0.0, 0.0)) == "하향식", _eng(b2, (0.0, 0.0)))
    chk("둘째 헤드", _eng(b2, (6000.0, 0.0)) == "상하향식",
        _eng(b2, (6000.0, 0.0)))
    chk("화면이 파생 그대로",
        b2.disk_kinds == disk_kind_list(b2.disks, b2.head_kinds),
        str(b2.disk_kinds))

    print("\n[2] 알람밸브 = 접속점 — 왕복 뒤에도 같은가")
    b3 = _board("왕복2")
    es = EditSession(b3, key="왕복2", out_dir=td)
    es.set_mode(MODE_VALVE)
    es.click(1500.0, 0.0, 500.0)
    before = (list(b3.valves), list(b3.sources))
    write_edits(b3, td)
    b4 = _board("왕복2")
    load_edits(b4, td)
    after = (list(b4.valves), list(b4.sources))
    chk("밸브가 남는다", len(after[0]) == 1, str(after))
    chk("접속점이 남는다", len(after[1]) == 1, str(after))
    chk("둘이 같은 절점", after[0] == after[1], f"{before} → {after}")

    print("\n[3] 되돌리기 — 왕복한 판에서도 한 번에 풀리나")
    b5 = _board("왕복3")
    es5 = EditSession(b5, key="왕복3", out_dir=td)
    es5.set_mode(MODE_VALVE)
    es5.click(1500.0, 0.0, 500.0)
    es5.click(1500.0, 0.0, 500.0)          # 토글로 끄기
    chk("토글로 끄면 둘 다 빈다",
        not b5.valves and not b5.sources,
        f"밸브 {list(b5.valves)} · 접속점 {list(b5.sources)}")

print("\n" + "=" * 56)
print(f"어긋남 {len(bad)}건" if bad else "왕복 불변량 전부 성립")
for x in bad:
    print("  -", x)
