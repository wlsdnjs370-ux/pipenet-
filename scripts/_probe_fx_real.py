# -*- coding: utf-8 -*-
"""[§29] 실도면에서 FX 를 켜면 «빠져 있던 양» 이 얼마나 메워지나."""
from __future__ import annotations
import os, sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
R = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(R); sys.path.insert(0, R)
KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")
from routes.module_f.common import _boot
_boot()
from services.cad_import.design.anchor import valve_kfp_nodes
from services.cad_import.design.restrict import select_and_expand
from services.cad_import.design.tables import build_design_tables
from services.cad_import.design.worst import worst_k_heads
from services.cad_import.edit.session import EditSession

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
w = worst_k_heads(b.pts, b.edges, b._head_nodes(), b.sources, k=30)
got = select_and_expand(es.convert_payload(), b, k=30)
av, _ = valve_kfp_nodes(got["kfp"].get("nodes_meta_runtime") or {}, b.pts,
                        list(b.valves), got.get("origin_mm"))
com = dict(node_head_kinds=got.get("node_head_kinds"), board_pts=b.pts, tree_loads=got.get("tree_loads"),
           origin_mm=got.get("origin_mm"), valve_nodes=av)
kinds = {}
for k in (got.get("node_head_kinds") or {}).values():
    kinds[k] = kinds.get(k, 0) + 1
print(f"\n헤드 종류 분포: {kinds}")
# ★이 도면이 전부 상향식이면 FX 는 «안 붙는 것이 맞다». 규칙이 도는지 보려면
#   종류를 하향식으로 바꿔 한 번 더 잰다(도면을 고치는 게 아니라 «만약» 이다).
_nhk = dict(got.get("node_head_kinds") or {})
_down = {k: "하향식" for k in _nhk}
for prof, nhk, tag in ((None, _nhk, "안 함"), ("평균", _nhk, "평균(실제 종류)"),
                       ("평균", _down, "평균(만약 전부 하향식이면)")):
    com2 = dict(com); com2["node_head_kinds"] = nhk
    t = build_design_tables(got["kfp"], w, got["edge_ref"], [],
                            fx_profile=prof, **com2)
    tot = sum(float(r.get("length") or 0) for r in t.pipes)
    eq = sum(float(e.get("eq_len") or 0) for e in t.equipment)
    n_fx = len([e for e in t.equipment if e["desc"] == "FX"])
    print(f"  FX {tag:>22} — 배관 {tot:.1f}m · 기기 등가길이 {eq:.1f}m"
          f" ({eq/max(tot,1e-9)*100:.0f}%) · FX {n_fx}개 · "
          f"meta {dict(t.meta)['신축배관(FX)']}")
