# -*- coding: utf-8 -*-
"""모델 단위에 고정한 글자 높이가 우리 파일에서는 몇 pt 로 떨어지는지 본다.

참조 40 장에서는 2.15~4.03 pt 였다. 우리 SDF 가 좌표 눈금을 다르게 쓰면
같은 상수라도 글씨가 안 보이거나 종이를 덮는다 — 옮기기 전에 확인한다.
"""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from core.d_display_model import load_display_model  # noqa: E402

LABEL_UNITS = 24.17          # 참조 실측 (모델 단위, 변동계수 0.029)
PLATE_PT = (0.930 - 0.156) * 842.0      # 그림틀 세로 (pt)

for sdf in sorted((ROOT / "routes" / "제출용[최종]").glob("*.sdf")):
    model = load_display_model(sdf)
    real = [n for n in model.nodes if not n.virtual]
    if len(real) < 2:
        print(f"{sdf.name}: 노드가 모자라 못 잰다")
        continue
    mspan = max(max(n.x for n in real) - min(n.x for n in real),
                max(n.y for n in real) - min(n.y for n in real))
    # 우리는 망을 그림틀에 맞춰 넣으므로 배율은 그림틀/모델 이다.
    scale = PLATE_PT / mspan if mspan else 0.0
    print(f"{sdf.name}: 노드 {len(real):3d}  모델폭 {mspan:10.2f}  "
          f"배율 {scale:8.5f}  글자 {LABEL_UNITS * scale:6.2f} pt")
