# -*- coding: utf-8 -*-
"""헤드 배치 수정 전/후를 같은 입력으로 나란히 잰다 (수정 전은 HEAD 판을 임시 적재)."""
from __future__ import annotations

import importlib.util
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "core")):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from core.d_display_model import load_display_model
from core.d_result_binder import bind_results
from core import d_iso_renderer as after

before_src = ROOT / "data" / "_dnoz_before_renderer.py"
before_src.write_bytes(subprocess.run(
    ["git", "show", "HEAD:core/d_iso_renderer.py"], cwd=ROOT, check=True,
    stdout=subprocess.PIPE).stdout)
spec = importlib.util.spec_from_file_location("_dnoz_before_renderer", before_src)
before = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = before   # dataclass 가 자기 모듈을 sys.modules 에서 되찾는다
spec.loader.exec_module(before)

SUB = ROOT / "routes" / "제출용[최종]"
OUT = ROOT / "data" / "_dnoz"
OUT.mkdir(exist_ok=True)

for stem in ("2. Pipenet_auto", "2. Pipenet_hand"):
    model = load_display_model(SUB / f"{stem}.sdf")
    bound = bind_results(model, SUB / f"{stem}.xml")
    tag = "auto" if "auto" in stem else "hand"
    for link, node in (("Pipe volumetric flow", "None"),
                       ("Nozzle calculated pressure", "None")):
        for name, mod in (("전", before), ("후", after)):
            png = OUT / f"{name}_{tag}_{link.split()[0].lower()}.png"
            rep = mod.render_iso(bound, png.with_suffix(".pdf"), link_item=link,
                                 node_item=node, section="헤드 배치", also_png=png)
            print(f"{name} {tag:<5} {link:<26} 노즐 {rep.nozzles_drawn:>3}"
                  f" 라벨 {rep.labels_drawn:>4} 지시선 {rep.labels_with_leader:>4}"
                  f" 탈락 {len(rep.labels_dropped):>3}"
                  f" 유도 {len(getattr(rep, 'nozzles_derived', ())):>3}"
                  f" 근거없음 {len(getattr(rep, 'nozzles_undirected', ())):>3}")
    print()

before_src.unlink()
