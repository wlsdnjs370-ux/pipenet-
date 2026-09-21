# -*- coding: utf-8 -*-
"""헤드 배치 — 실제 렌더 결과와 라벨 탈락 실측."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from core.d_display_model import load_display_model
from core.d_iso_renderer import render_iso
from core.d_result_binder import bind_results

SUB = ROOT / "routes" / "제출용[최종]"
OUT = ROOT / "data" / "_dnoz"
OUT.mkdir(exist_ok=True)

for stem in ("2. Pipenet_auto", "2. Pipenet_hand"):
    model = load_display_model(SUB / f"{stem}.sdf")
    bound = bind_results(model, SUB / f"{stem}.xml")
    tag = "auto" if "auto" in stem else "hand"
    for link, node in (("Pipe volumetric flow", "None"),
                       ("Nozzle calculated pressure", "None")):
        name = f"{tag}_{link.split()[0].lower()}"
        rep = render_iso(bound, OUT / f"{name}.png", link_item=link, node_item=node,
                         section="헤드 배치 진단")
        print(f"{tag:<5} {link:<26} 노즐 {rep.nozzles_drawn:>3} 라벨 {rep.labels_drawn:>4}"
              f" 지시선 {rep.labels_with_leader:>4} 탈락 {len(rep.labels_dropped):>3}"
              f" 빈칸관로 {len(rep.blank_link_values):>3}")
        if rep.labels_dropped:
            print("     탈락:", ", ".join(rep.labels_dropped[:15]))
