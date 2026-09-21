# -*- coding: utf-8 -*-
"""상하향식(combo) 헤드가 «어디에» 박히나 — 아래 헤드의 +X 치우침 실측.

`build_combo` 는 아래 헤드를 언제나 **+X 방향** 으로 ③(combo_2, 기본 0.3m)
만큼 옮긴다. 가지관이 어느 쪽으로 뻗든 상관없이 +X 다. 하향 팔이 관을 따라
가지 않고 늘 동쪽으로 뻗는 셈이라, 화면에서 그 헤드만 «엉뚱한 데 박힌» 것으로
보인다.

수리계산 자체는 어긋나지 않는다 — 배관 길이를 좌표로 재지 않고 입력값
(combo_2)을 그대로 쓰기 때문이다. 어긋나는 것은 **그림** 이다.
"""
from __future__ import annotations

import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

KEY = os.environ.get("MF_KEY", "1. 입력도면 대명동 단위세대 평면도")
N_COMBO = 2          # 사용자가 말한 «헤드 두 개»

from routes.module_f.common import _boot                       # noqa: E402
_boot()
from services.cad_import.edit.session import (                 # noqa: E402
    MODE_SOURCE, EditSession)
from services.cad_import.design.restrict import select_and_expand    # noqa: E402

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
if not b.sources:
    es.set_mode(MODE_SOURCE)
    for (u, v) in b.edges:
        if es.click((b.pts[u][0] + b.pts[v][0]) / 2,
                    (b.pts[u][1] + b.pts[v][1]) / 2, 2000.0):
            break

kinds = list(b.disk_kinds)
print("헤드 종류(원래):", {k: kinds.count(k) for k in set(kinds)})


def run(tag):
    payload = es.convert_payload()
    got = select_and_expand(payload, b, k=30, selected_source=None)
    if not got.get("ok"):
        print(f"[{tag}] 실패:", got.get("error"))
        return None
    net = got["kfp"]
    meta = net["nodes_meta_runtime"]
    pump = next(n for n, m in meta.items()
                if str((m or {}).get("type_id", "")) == "pump")
    pc = meta[pump]["coords"]
    si = b.sources[0]
    ox = float(b.pts[si][0]) - float(pc[0]) * 1000.0
    oy = float(b.pts[si][1]) - float(pc[1]) * 1000.0
    heads = [(nid, m) for nid, m in meta.items()
             if str((m or {}).get("type_id", "")) == "head"]
    disks = [(float(d[0]), float(d[1])) for d in b.disks]
    out = []
    for nid, m in heads:
        c = m["coords"]
        hx = float(c[0]) * 1000.0 + ox
        hy = float(c[1]) * 1000.0 + oy
        d2, bi = min(((hx - dx) ** 2 + (hy - dy) ** 2, i)
                     for i, (dx, dy) in enumerate(disks))
        out.append((math.sqrt(d2), nid, bi, hx, hy, float(c[2])))
    globals()["LAST_WORST"] = list((got.get("worst") or {}).get("heads") or [])
    print(f"[{tag}] 헤드 노드 {len(out)}개 · "
          f"찍은 자리에서 300mm 넘게 벗어난 것 "
          f"{sum(1 for r in out if r[0] > 300)}개")
    for r in sorted(out, reverse=True)[:4]:
        print(f"      벗어남 {r[0]:7.0f}mm · z={r[5]:.2f}m · "
              f"board #{r[2]} 기준")
    return out


run("전 — 상향식만")

# ★«아무 헤드나» 가 아니라 **실제로 뽑히는 최불리 헤드** 를 바꿔야 한다.
#   선택 밖의 헤드를 바꾸면 전개 자체에 안 들어와 아무 변화가 없다
#   (처음에 그렇게 재서 «차이 없음» 이라는 헛된 답을 얻었다).
# ★`disk_kinds` 를 직접 바꾸면 안 된다 — 그것은 «화면 색» 이고, 전개가 읽는
#   종류는 `head_kinds` 레코드다(engine._resolved_kinds). 사람이 쓰는 경로인
#   `board.set_head_kind` 로 바꾼다. 처음에 disk_kinds 를 바꿔 재고 «차이
#   없음» 이라는 헛된 답을 얻었다.
idx = list(globals().get("LAST_WORST") or [])[:N_COMBO]
for i in idx:
    b.set_head_kind(b.disks[i], "상하향식")   # 인덱스가 아니라 디스크 자체
print(f"\n헤드 {idx} 를 «상하향식» 으로 바꿈")
print("헤드 종류(바꾼 뒤):",
      {k: list(b.disk_kinds).count(k) for k in set(b.disk_kinds)})

run("후 — 상하향식 2개")
