# -*- coding: utf-8 -*-
"""밑그림 변환이 정말 «표의 절점과 같은 자리» 를 내는가 — 화면을 만들기 전에.

board 헤드 부착점을 그 식으로 옮겨, 표의 헤드 절점 좌표와 맞대 본다.
어긋나면 화면에 무엇을 그려도 소용없다.
"""
from __future__ import annotations

import math
import os
import statistics
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
os.chdir(ROOT)
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

KEY = os.environ.get("MF_KEY", "B1F 현장조사 소화설비 평면도")

from routes.module_f.common import _boot                          # noqa: E402
_boot()
from services.cad_import.design.emit import display_tables        # noqa: E402
from services.cad_import.design.restrict import select_and_expand  # noqa: E402
from services.cad_import.design.tables import build_design_tables  # noqa: E402
from services.cad_import.design.worst import worst_k_heads        # noqa: E402
from services.cad_import.edit.session import EditSession          # noqa: E402

es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
b = es.board
w = worst_k_heads(b.pts, b.edges, b._head_nodes(), b.sources, k=30)
got = select_and_expand(es.convert_payload(), b, k=30)
tbl = build_design_tables(got["kfp"], w, got["edge_ref"], [],
                          board_pts=b.pts, tree_loads=got.get("tree_loads"),
                          origin_mm=got.get("origin_mm"))
origin = got["origin_mm"]
kfp_nodes = got["kfp"]["nodes_meta_runtime"]

# 표 절점 ↔ kfp 노드 (표는 kfp 좌표를 mm 로 올려 실었다)
lab_of = {}
for n in tbl.nodes:
    lab_of[(round(n["x"] / 1000.0, 3), round(n["y"] / 1000.0, 3))] = str(n["label"])


def xf(mx, my, u):
    """라우트가 화면에 보내는 그 식 그대로 — 합성된 한 장."""
    nx = u["k"] * mx + u["tx"]
    ny = u["k"] * my + u["ty"]
    if not u["iso"]:
        return nx, ny
    dy = (u["e"] - u["e_ref"]) * u["lift"]
    return (nx - ny) * u["cos30"], (nx + ny) * u["sin30"] + dy


for iso in (False, True):
    view, stood = display_tables(tbl, iso=iso, canvas_units=3000.0)
    norm = view.norm
    kk = float(norm["scale"])
    u = {"k": kk,
         "tx": kk * (1000.0 - float(origin[0]) - float(norm["cx"])),
         "ty": kk * (1000.0 - float(origin[1]) - float(norm["cy"])),
         "cos30": norm["cos30"], "sin30": norm["sin30"], "iso": iso,
         "lift": float((stood or {}).get("lift") or 0.0),
         "e_ref": float((stood or {}).get("e_ref") or 0.0),
         "e": next((float(n.get("elevation", 0.0) or 0.0) for n in view.nodes
                    if str(n.get("io_node")) == "Input"),
                   min((float(n.get("elevation", 0.0) or 0.0)
                        for n in view.nodes), default=0.0))}
    at = {str(n["label"]): n for n in view.nodes}
    head_labs = {str(r.get("in")) for r in (tbl.nozzles or ())}

    errs, errs_head = [], []
    hn = b._head_nodes()
    for hi in w["heads"]:
        for bn in [n for n in hn[hi] if n < len(b.pts)]:
            tx = (b.pts[bn][0] - origin[0]) / 1000.0 + 1.0
            ty = (b.pts[bn][1] - origin[1]) / 1000.0 + 1.0
            near = sorted((math.hypot(float((m.get("coords") or [0, 0])[0]) - tx,
                                      float((m.get("coords") or [0, 0])[1]) - ty),
                           nid)
                          for nid, m in kfp_nodes.items()
                          if str(m.get("type_id")) == "head")
            if not near or near[0][0] > 0.10:
                break
            c = kfp_nodes[near[0][1]]["coords"]
            lab = lab_of.get((round(float(c[0]), 3), round(float(c[1]), 3)))
            if lab is None or lab not in at:
                break
            X, Y = xf(float(b.pts[bn][0]), float(b.pts[bn][1]), u)
            d = math.hypot(X - float(at[lab]["x"]), Y - float(at[lab]["y"]))
            (errs_head if lab in head_labs else errs).append(d)
            break

    allv = errs + errs_head
    span = max(max(float(n["x"]) for n in view.nodes)
               - min(float(n["x"]) for n in view.nodes),
               max(float(n["y"]) for n in view.nodes)
               - min(float(n["y"]) for n in view.nodes))
    tag = "아이소" if iso else "평면"
    print(f"\n■ {tag} 보기 — 밑그림 식으로 옮긴 board 점 vs 표 절점 ({len(allv)}개)")
    if allv:
        print(f"    중앙 {statistics.median(allv):.3f} · 최대 {max(allv):.3f} "
              f"· 한 변({span:,.0f})의 {max(allv) / span * 100:.4f}%")
    if errs_head:
        print(f"    ※ 그중 «세운 헤드» {len(errs_head)}개는 최대 "
              f"{max(errs_head):.1f} — 아이소가 일부러 세운 만큼이다(§G15).")
    if errs:
        print(f"    ※ 세우지 않은 절점 {len(errs)}개: 최대 {max(errs):.3f}")
    if stood:
        print(f"    stub={stood.get('stub'):.2f} · lift={stood.get('lift'):.1f}"
              f" · e_ref={stood.get('e_ref'):.3f}")
    # ── 헤드가 아닌 점 — 접속점(board 급수 노드 ↔ io_node=Input)
    root = next((str(n["label"]) for n in tbl.nodes
                 if str(n.get("io_node")) == "Input"), None)
    if root and root in at and b.sources:
        bn = b.sources[0]
        X, Y = xf(float(b.pts[bn][0]), float(b.pts[bn][1]), u)
        d = math.hypot(X - float(at[root]["x"]), Y - float(at[root]["y"]))
        ez = float(at[root].get("elevation", 0.0) or 0.0)
        print(f"    ★세우지 않는 접속점: 어긋남 {d:.3f} "
              f"(한 변의 {d / span * 100:.4f}%) · 그 절점 표고 {ez:.3f} m"
              f" · e_ref 와 차 {abs(ez - u['e_ref']):.3f} m")
