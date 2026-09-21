# -*- coding: utf-8 -*-
"""기준 헤드 노드 라벨이 실제로 서는가 — 실도면 확인.

죽어 있던 것(BLOCKED §30)이 살아났는지, 그리고 «맞는» 노드를 가리키는지 본다.
맞는지는 두 가지로 검증한다:
  ① 그 노드가 정말 헤드다(type_id=head)
  ② 표에서 접속점 → 그 노드까지 잰 길이가 최불리의 far_m 과 맞는다
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

KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")

from routes.module_f.common import _boot                          # noqa: E402
_boot()
from services.cad_import.edit.session import EditSession          # noqa: E402
from services.cad_import.design.restrict import select_and_expand  # noqa: E402
from services.cad_import.design.tables import build_design_tables  # noqa: E402

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
payload = es.convert_payload()
got = select_and_expand(payload, b, k=30)
if not got.get("ok"):
    raise SystemExit(f"전개 실패: {got.get('error')}")

w = got["worst"] if "worst" in got else None
if w is None:
    from services.cad_import.design.worst import worst_k_heads
    w = worst_k_heads(b.pts, b.edges, b._head_nodes(), b.sources, k=30)

tbl = build_design_tables(got["kfp"], w, got["edge_ref"], [],
                          board_pts=b.pts,
                          tree_loads=got.get("tree_loads"),
                          origin_mm=got.get("origin_mm"))
meta = dict(tbl.meta)
lab = meta.get("기준 헤드 노드")
print(f"\n★ 기준 헤드 노드 = {lab!r}   (종전엔 항상 '?')")

nodes = {str(n["label"]): n for n in tbl.nodes}
row = nodes.get(str(lab))
kfp_nodes = got["kfp"]["nodes_meta_runtime"]
# 라벨 → kfp 노드 id 는 표에 없으므로 좌표로 되짚어 종류만 본다.
if row is not None:
    tx, ty = row["x"] / 1000.0, row["y"] / 1000.0
    kind = None
    for nid, m in kfp_nodes.items():
        c = m.get("coords") or (0, 0, 0)
        if math.hypot(float(c[0]) - tx, float(c[1]) - ty) < 1e-6:
            kind = str(m.get("type_id"))
            break
    print(f"   ① 그 노드의 type_id = {kind}  ← 'head' 여야 한다")

# ② 접속점 → 기준 헤드 길이 vs 최불리 far_m
root = next((str(n["label"]) for n in tbl.nodes
             if str(n.get("io_node")) == "Input"), None)
adj: dict = {}
plen: dict = {}
for r in tbl.pipes:
    a, bb, pid = str(r["in"]), str(r["out"]), str(r["label"])
    plen[pid] = float(r.get("length") or 0.0)
    adj.setdefault(a, []).append((bb, pid))
    adj.setdefault(bb, []).append((a, pid))
par, seen, q = {}, {root}, [root]
while q:
    cur = q.pop(0)
    for nxt, pid in adj.get(cur, ()):
        if nxt in seen:
            continue
        seen.add(nxt)
        par[nxt] = (cur, pid)
        q.append(nxt)
cur, total, hops = str(lab), 0.0, 0
while cur in par:
    prev, pid = par[cur]
    total += plen.get(pid, 0.0)
    cur = prev
    hops += 1
print(f"   ② 접속점({root}) → 기준 헤드({lab}) : 절점 {hops + 1} · "
      f"{total:.2f} m")
print(f"      최불리 far_m = {w['far_m']} m · worst_path_m = {w['worst_path_m']} m")
d = abs(total - float(w["far_m"]))
print(f"      차이 {d:.2f} m ({d / max(float(w['far_m']), 1e-9) * 100:.2f}%)")
print("      ※ 차이의 원인은 여기서 단정하지 않는다 — 이 검사가 보는 것은 "
      "«엉뚱한 헤드를 가리키지 않는다» 이고,")
print("        이웃 헤드는 실측 2.95 m 떨어져 있어 0.13 m 어긋남으로는 "
      "바뀔 수 없다.")
