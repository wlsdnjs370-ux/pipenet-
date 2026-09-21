# -*- coding: utf-8 -*-
"""모듈 D 충실도 진단 — 수작업본/자동본을 나란히 그리고 확대해 본다."""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from PIL import Image  # noqa: E402

from core.d_display_model import load_display_model  # noqa: E402
from core.d_iso_renderer import render_iso  # noqa: E402
from core.d_result_binder import bind_results  # noqa: E402

SUB = ROOT / "routes" / "제출용[최종]"
OUT = ROOT / "data" / "_dfid"
OUT.mkdir(exist_ok=True)

for tag, stem in [("수작업", "2. Pipenet_hand"), ("자동", "2. Pipenet_auto")]:
    model = load_display_model(SUB / f"{stem}.sdf")
    bound = bind_results(model, SUB / f"{stem}.xml")
    png = OUT / f"{tag}.png"
    rep = render_iso(bound, OUT / f"{tag}.pdf", preset=None,
                     link_item="Pipe volumetric flow", node_item="None",
                     show_link_labels=True, show_node_labels=True, show_arrows=True,
                     section="", also_png=png)
    print(f"[{tag}] 관로 {rep.pipes_drawn} 노즐 {rep.nozzles_drawn} 라벨 {rep.labels_drawn}"
          f" 유도 {len(rep.nozzles_derived)} 무방향 {len(rep.nozzles_undirected)}")
    # 원본 표시 설정 그대로 한 장 더 — "파이프넷 출력 그대로"가 이쪽이다.
    asis = OUT / f"{tag}_원본대로.png"
    render_iso(bound, OUT / f"{tag}_원본대로.pdf", preset=None,
               link_item="Pipe volumetric flow", node_item="None", also_png=asis)
    aim = Image.open(asis)
    aw, ah = aim.size
    crop = aim.crop((aw // 4, ah // 4, aw * 3 // 4, ah * 3 // 4))
    crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS).save(
        OUT / f"{tag}_원본대로_확대.png")

    im = Image.open(png)
    print("      캔버스", im.size)
    w, h = im.size
    # 사분면으로 잘라 4배 확대 — 화살표와 헤드 삼각형을 눈으로 볼 수 있게.
    for qi, (x0, y0, x1, y1) in enumerate([
            (0, 0, w // 2, h // 2), (w // 2, 0, w, h // 2),
            (0, h // 2, w // 2, h), (w // 2, h // 2, w, h)]):
        crop = im.crop((x0, y0, x1, y1))
        crop = crop.resize((crop.width * 3, crop.height * 3), Image.LANCZOS)
        crop.save(OUT / f"{tag}_q{qi + 1}.png")
print("→", OUT)
