# -*- coding: utf-8 -*-
"""오늘 만진 자리들이 서로 어긋나지 않는가 — 불변량을 한 판에 모아 잰다.

각 항목은 «두 곳이 같은 답을 내야 한다» 는 형태다. 갈라져 있으면 그 자체가
모순이고, 대개 한쪽만 고쳐진 흔적이다.
"""
from __future__ import annotations

import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_G = os.path.join(_ROOT, "cad_project_editor_g")
for _p in (_ROOT, _G):
    if _p not in sys.path:
        sys.path.insert(0, _p)

bad: list[str] = []


def chk(name, ok, detail=""):
    print(f"  [{'OK  ' if ok else '★어긋남'}] {name}" + (f" · {detail}" if detail else ""))
    if not ok:
        bad.append(f"{name} — {detail}")


print("[1] 헤드 종류 — 화면(disk_kinds)은 엔진(head_kinds)의 파생인가")
from services.cad_import.edit.board import EditBoard
from services.cad_import.kinds import disk_kind_list, resolve_head_kinds

b = EditBoard("일관성", [(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)],
              [(0, 1), (1, 2)], [(0.0, 0.0, 150.0), (6000.0, 0.0, 150.0)],
              head_kinds=[])
states = []
b.set_head_kind(b.disks[0], "하향식")
states.append(("찍은 뒤", b.disk_kinds == disk_kind_list(b.disks, b.head_kinds)))
b.set_head_kind(b.disks[1], "상하향식")
states.append(("둘째까지", b.disk_kinds == disk_kind_list(b.disks, b.head_kinds)))
b.undo()
states.append(("undo 뒤", b.disk_kinds == disk_kind_list(b.disks, b.head_kinds)))
b.undo()
states.append(("undo 둘", b.disk_kinds == disk_kind_list(b.disks, b.head_kinds)))
for label, ok in states:
    chk(f"파생 불변량 — {label}", ok)

print("\n[2] 접속점 — «앵커» 라는 낱말이 한 뜻만 갖는가")
from services.cad_import.design.worst import worst_k_heads
w = worst_k_heads([(0.0, 0.0), (3000.0, 0.0), (6000.0, 0.0)],
                  [(0, 1), (1, 2)], [{2}], [0], k=1)
chk("최불리는 «기준 헤드» 라 부른다",
    "worst_head" in w and "anchor" not in w, str(sorted(w)[:6]))
from routes.module_f.merge import ANCHOR_LABEL, LABEL_OFFSET
chk("merge 의 앵커는 접속점(라벨 10)",
    ANCHOR_LABEL == "10" and LABEL_OFFSET == 9,
    f"{ANCHOR_LABEL} · +{LABEL_OFFSET}")

print("\n[3] 등가길이 — 부속표와 기기표가 같은 함수를 쓰는가")
from services.cad_import.design.fitting import (
    load_equivalent_lengths, resolve_eq_len)
lib = load_equivalent_lengths()
a1 = resolve_eq_len("elbow", 100, lib=lib)
a2 = resolve_eq_len("elbow", 100, lib=lib, ov_eq={("elbow", 100): (99.0, "x")})
chk("라이브러리 값은 사람이 못 덮는다", a1 == a2, f"{a1} vs {a2}")
chk("못 구하면 None (0 이 아니다)",
    resolve_eq_len("elbow", 9999, lib=lib) == (None, None))

print("\n[4] 좌표계 — board → kfp 는 한 규칙인가")
from services.cad_import.convert.main_walk import xf_mm_to_m
from services.cad_import.design.anchor import VALVE_SNAP_M
from services.cad_import.design.tables import WORST_HEAD_SNAP_M
mx, my, ox, oy = 646763.6, 88854.1, 646000.0, 88000.0
by_fn = xf_mm_to_m(mx, my, ox, oy)
by_hand = ((mx - ox) / 1000.0 + 1.0, (my - oy) / 1000.0 + 1.0)
chk("xf_mm_to_m 이 그 규칙 그대로", by_fn == by_hand, f"{by_fn}")
chk("되짚기 허용치 두 개가 같은 근거",
    VALVE_SNAP_M == WORST_HEAD_SNAP_M,
    f"밸브 {VALVE_SNAP_M} · 기준헤드 {WORST_HEAD_SNAP_M}")

print("\n[5] 밑그림 — 정규화 파라미터가 엔진 것과 같은가")
import types
from services.cad_import.design.sdf_post import (
    node_norm_params, normalize_node_coords)
nodes = [{"label": "1", "x": 1000, "y": 2000, "elevation": 0.0},
         {"label": "2", "x": 5000, "y": 4000, "elevation": 0.0}]
t = types.SimpleNamespace(nodes=[dict(n) for n in nodes], pipes=[])
cx, cy, sc = node_norm_params(t, canvas_units=3000.0)
got = normalize_node_coords(t, canvas_units=3000.0)
same = (got == sc and all(
    abs(a["x"] - (o["x"] - cx) * sc) < 1e-9
    and abs(a["y"] - (o["y"] - cy) * sc) < 1e-9
    for o, a in zip(nodes, t.nodes)))
chk("두 함수가 같은 변환", same, f"scale {sc}")

print("\n[6] 접속점 없는 망 — 두 소비처가 같은 태도인가")
from services.cad_import.design.anchor import AnchorMissing, require_anchor
net = {"pipe_data": {"P1": {"start": "N1", "end": "N2", "length_m": 1.0}},
       "nodes_meta_runtime": {"N1": {"coords": [0, 0, 0], "type_id": "base"},
                              "N2": {"coords": [1, 0, 0], "type_id": "base"}}}
try:
    require_anchor(net["nodes_meta_runtime"])
    chk("접속점 없으면 던진다", False, "안 던졌다")
except AnchorMissing:
    chk("접속점 없으면 던진다", True)

print("\n" + "=" * 56)
print(f"어긋남 {len(bad)}건" if bad else "불변량 전부 성립")
for x in bad:
    print("  -", x)
