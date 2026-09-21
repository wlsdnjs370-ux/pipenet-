# -*- coding: utf-8 -*-
"""수리계산 화면의 헤드가 «평면도에서 찍은 자리» 와 어긋나는가.

설계 표의 노드 좌표는 kfp 그래프에서 곧장 온다(tables.py 의 `xy()`).
그러므로 화면에서 헤드가 엉뚱한 데 있으면 표시가 아니라 **망 자체** 다.

board 헤드(사람이 평면도에서 찍은 원)와 설계 표의 헤드 노드를 좌표로 맞대
어긋난 것을 찾는다. 좌표계는 kfp 가 board 에서 `origin_mm` 만큼 옮기고
m 로 바꾼 것이라, 되돌려서 비교한다.
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

from routes.module_f.common import _boot                      # noqa: E402
_boot()
from services.cad_import.edit.session import (                # noqa: E402
    MODE_SOURCE, EditSession)
from services.cad_import.design.restrict import select_and_expand   # noqa: E402
from services.cad_import.design.tables import build_design_tables   # noqa: E402

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
print(f"board — 노드 {len(b.pts)} · 간선 {len(b.edges)} · 헤드 {len(b.disks)}")

# 급수원이 없으면 사람이 하듯 «클릭» 으로 하나 찍는다(간선 중점).
if not b.sources:
    es.set_mode(MODE_SOURCE)
    for (u, v) in b.edges:
        mx = (b.pts[u][0] + b.pts[v][0]) / 2
        my = (b.pts[u][1] + b.pts[v][1]) / 2
        if es.click(mx, my, 2000.0):
            break
print("급수원", len(b.sources), "곳")

payload = es.convert_payload()
got = select_and_expand(payload, b, k=30, selected_source=None)
print("select_and_expand ok =", got.get("ok"), got.get("error") or "")
if not got.get("ok"):
    raise SystemExit(1)

net = got["kfp"]
meta = net["nodes_meta_runtime"]
w = got["worst"]
# ★원점을 «추측» 하지 않는다. payload 의 origin_mm 은 이 경로에서 (0,0) 이라
#   그것으로 맞추면 전부 어긋난 것처럼 보인다(실측으로 24/24 가 그랬다).
#   급수원은 board 와 kfp 양쪽에 «같은 한 점» 으로 있으므로 그것으로 잰다.
pump = next((n for n, m in meta.items()
             if str((m or {}).get("type_id", "")) == "pump"), None)
pc = meta[pump]["coords"]
src_i = b.sources[0]
bx, by = float(b.pts[src_i][0]), float(b.pts[src_i][1])
M = 1000.0
ox = bx - float(pc[0]) * M
oy = by - float(pc[1]) * M
print(f"급수원 — board({bx:.0f},{by:.0f}) kfp({pc[0]*M:.0f},{pc[1]*M:.0f})")
print(f"→ 좌표 옮김 = ({ox:.0f}, {oy:.0f}) · 최불리 헤드 {len(w['heads'])}")
print("   got 키:", sorted(k for k in got if not k.startswith('_'))[:12])

tbl = build_design_tables(net, w, got["edge_ref"], [], board_pts=b.pts,
                          excluded_heads=got.get("excluded_heads", 0),
                          tree_loads=got.get("tree_loads"))
node_of = {str(r["label"]): r for r in tbl.nodes}

# board 헤드 좌표(사람이 평면도에서 찍은 것)
disks = [(float(d[0]), float(d[1])) for d in b.disks]

print(f"\n설계 표 — 노드 {len(tbl.nodes)} · 배관 {len(tbl.pipes)} "
      f"· 노즐 {len(tbl.nozzles)}")
print("노즐마다 «가장 가까운 board 헤드» 까지 거리(mm):\n")

rows = []
for nz in tbl.nozzles:
    lab = str(nz.get("in"))
    n = node_of.get(lab)
    if n is None:
        rows.append((None, lab, "표에 노드가 없다"))
        continue
    # 설계 표 노드는 mm(원점 이동됨) — board 좌표로 되돌린다.
    hx, hy = float(n["x"]) + ox, float(n["y"]) + oy
    best, bi = 1e18, -1
    for i, (dx, dy) in enumerate(disks):
        dd = (hx - dx) ** 2 + (hy - dy) ** 2
        if dd < best:
            best, bi = dd, i
    rows.append((math.sqrt(best), lab, f"board #{bi} "
                 f"({disks[bi][0]:.0f},{disks[bi][1]:.0f}) "
                 f"vs 표({hx:.0f},{hy:.0f}) z={n['elevation']}"))

ok = [r for r in rows if r[0] is not None and r[0] <= 300]
bad = [r for r in rows if r[0] is None or r[0] > 300]
print(f"300mm 안에 맞는 것 {len(ok)} · 어긋난 것 {len(bad)}\n")
for d, lab, why in sorted(bad, key=lambda t: -(t[0] or 1e18)):
    print(f"  ★ 노즐 in={lab}  거리 {('%.0f' % d) if d else '—'}mm  {why}")
for d, lab, why in sorted(ok, key=lambda t: -(t[0] or 0))[:5]:
    print(f"     노즐 in={lab}  거리 {d:.0f}mm  {why}")
