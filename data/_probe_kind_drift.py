# -*- coding: utf-8 -*-
"""헤드 종류 두 원천(disk_kinds/화면 · head_kinds/엔진)이 갈리는가 — 실측.

가설: 분류가 레코드를 못 만든 헤드에 사람이 종류를 찍으면,
  ① 화면(disk_kinds)은 set_head_kind 의 덧댐 블록이 칠해 준다
  ② 엔진(head_kinds)에는 apply_kind_overrides 가 «기존 레코드 갱신만» 이라
     아무것도 안 남는다
  ③ 저장→재열기(io.py)는 apply(갱신만) → require(미지정 삽입) 순서라
     사람의 선택이 통째로 사라진다
"""
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
for p in (_G,):
    if p not in sys.path:
        sys.path.insert(0, p)

from services.cad_import.edit.board import EditBoard
from services.cad_import.kinds import require_head_kinds, disk_kind_list

# 헤드 두 개 — 배관 끝에 하나씩. head_kinds 는 «분류 실패» 로 비어 있다.
pts = [(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)]
edges = [(0, 1), (1, 2)]
disks = [(0.0, 0.0, 150.0), (6000.0, 0.0, 150.0)]
b = EditBoard("탐침", pts, edges, disks, head_kinds=[])

got = b.set_head_kind(disks[0], "하향식")
print("set_head_kind →", got)

# ① 화면이 보는 것
print("화면 disk_kinds[0] =", b.disk_kinds[0])
# ② 엔진이 보는 것 — planar/preflight 는 require_head_kinds 를 거친다
eng = require_head_kinds(b.disks, b.head_kinds)
by = {tuple(r["c"]): r["kind"] for r in eng}
print("엔진 head_kinds[disk0] =", by.get((0.0, 0.0), "레코드 없음"))
# 파생 불변량
print("파생 일치? ", b.disk_kinds == disk_kind_list(b.disks, b.head_kinds))

# ③ 저장→재열기 — io.py 순서 재현
from services.cad_import.pipeline.user_net import apply_kind_overrides
from services.cad_import.kinds import resolve_head_kinds
hk2 = resolve_head_kinds(b.disks, [], b.kind_overrides)   # io.py 새 경로
by2 = {tuple(r["c"]): r["kind"] for r in hk2}
print("재열기 후[disk0] =", by2.get((0.0, 0.0)))
