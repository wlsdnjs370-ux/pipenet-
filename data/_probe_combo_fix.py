# -*- coding: utf-8 -*-
"""상하향식 하향 팔 방향 고침 — «그림만 바뀌고 수리계산은 그대로» 인가.

바뀌어야 하는 것 : 아래 헤드의 x·y (가지를 따라간다)
바뀌면 안 되는 것: 배관 길이·등가길이·관경·표고·부속·노즐 유량 — 전부
                   좌표가 아니라 입력값·규칙에서 나오므로 그대로여야 한다.
"""
from __future__ import annotations

import json
import math
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

KEY = os.environ.get("MF_KEY", "1. 입력도면 대명동 단위세대 평면도")

from routes.module_f.common import _boot                       # noqa: E402
_boot()
from services.cad_import.edit.session import (                 # noqa: E402
    MODE_SOURCE, EditSession)
from services.cad_import.design.restrict import select_and_expand    # noqa: E402
from services.cad_import.design.tables import build_design_tables    # noqa: E402

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
if not b.sources:
    es.set_mode(MODE_SOURCE)
    for (u, v) in b.edges:
        if es.click((b.pts[u][0] + b.pts[v][0]) / 2,
                    (b.pts[u][1] + b.pts[v][1]) / 2, 2000.0):
            break

# 최불리에 실제로 뽑히는 헤드 둘을 상하향식으로.
got0 = select_and_expand(es.convert_payload(), b, k=30, selected_source=None)
worst = list((got0.get("worst") or {}).get("heads") or [])[:2]
for i in worst:
    b.set_head_kind(b.disks[i], "상하향식")
print("상하향식으로 바꾼 헤드:", worst)

got = select_and_expand(es.convert_payload(), b, k=30, selected_source=None)
net = got["kfp"]
meta = net["nodes_meta_runtime"]
tbl = build_design_tables(net, got["worst"], got["edge_ref"], [],
                          board_pts=b.pts,
                          excluded_heads=got.get("excluded_heads", 0),
                          tree_loads=got.get("tree_loads"))

pump = next(n for n, m in meta.items()
            if str((m or {}).get("type_id", "")) == "pump")
pc = meta[pump]["coords"]
si = b.sources[0]
ox = float(b.pts[si][0]) - float(pc[0]) * 1000.0
oy = float(b.pts[si][1]) - float(pc[1]) * 1000.0
disks = [(float(d[0]), float(d[1])) for d in b.disks]

far = []
for nid, m in meta.items():
    if str((m or {}).get("type_id", "")) != "head":
        continue
    c = m["coords"]
    hx = float(c[0]) * 1000.0 + ox
    hy = float(c[1]) * 1000.0 + oy
    d2, bi = min(((hx - dx) ** 2 + (hy - dy) ** 2, i)
                 for i, (dx, dy) in enumerate(disks))
    far.append((math.sqrt(d2), nid, float(c[2])))
far.sort(reverse=True)
print(f"\n헤드 노드 {len(far)}개 · 찍은 자리에서 가장 먼 것 4개")
for d, nid, z in far[:4]:
    print(f"   {d:7.0f}mm · z={z:+.2f}m · {nid}")

# 수리계산에 쓰이는 값 — 좌표를 뺀 나머지를 통째로 지문 찍는다.
def fingerprint(t):
    """★라벨을 빼고 «값» 만 본다.

    표 라벨은 급수원 기점 BFS 로 매겨지므로 좌표가 조금만 달라져도 번호가
    통째로 재배열된다 — 라벨을 넣고 대조하면 무엇을 바꾸든 «전부 다름» 이
    나온다(실제로 그렇게 한 번 헛짚었다). 수리계산이 쓰는 것은 이름이
    아니라 값이므로 값의 multiset 으로 대조한다.
    """
    return {
        "pipes": sorted((r.get("dia"), r.get("length"), r.get("elev"),
                         r.get("c"), str(r.get("type")))
                        for r in t.pipes),
        "nozzles": sorted((r.get("flow_lmin"), r.get("flow_m3s"),
                           str(r.get("lib"))) for r in t.nozzles),
        "fittings": sorted((str(r.get("type")), r.get("count"))
                           for r in t.fittings),
        "elev": sorted(r.get("elevation") for r in t.nodes),
    }

out = os.environ.get("MF_FP", ROOT + "/data/_combo_fp.json")
json.dump(fingerprint(tbl), open(out, "w", encoding="utf-8"),
          ensure_ascii=False)
print("지문 저장 —", out)
