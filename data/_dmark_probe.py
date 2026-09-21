# -*- coding: utf-8 -*-
"""우리 기호(종이 pt)가 망(모델 단위)에 견줘 얼마나 커지는지 잰다."""
from __future__ import annotations

import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from core import d_iso_renderer as R  # noqa: E402
from core.d_display_model import load_display_model  # noqa: E402

LIB = ROOT / "data" / "reference_library"
BASE = ROOT / "routes" / "제출용[최종]"

# 참조 40 짝에서 나온 값 글씨 실측 pt 범위(모델 24.17 단위)로 배율 폭을 되짚는다.
for name, label_pt in (("참조 최소배율(가장 넓은 망)", 2.1525),
                       ("참조 중앙", 3.13395),
                       ("참조 최대배율", 4.028)):
    per_unit = label_pt / R._LABEL_UNITS
    print(f"{name}: 배율 {per_unit:.5f} pt/단위")
    print(f"    배관 굵기 {R._PIPE_WIDTH_UNITS * per_unit:.3f} pt  "
          f"노드점 {R._NODE_DOT_UNITS * per_unit:.3f} pt")
    print(f"    기기기호 {R._DEVICE_PT:.2f} pt = 노드점의 "
          f"{R._DEVICE_PT / (R._NODE_DOT_UNITS * per_unit):.2f} 배")
    print(f"    기기연결선 1.40 pt = 배관의 {1.4 / (R._PIPE_WIDTH_UNITS * per_unit):.1f} 배  "
          f"지시선 0.25 pt = 배관의 {0.25 / (R._PIPE_WIDTH_UNITS * per_unit):.1f} 배")

for stem in ("2. Pipenet_hand.sdf", "2. Pipenet_auto.sdf"):
    p = BASE / stem
    if not p.exists():
        continue
    m = load_display_model(p)
    xs = [n.x for n in m.nodes if not n.virtual]
    ys = [n.y for n in m.nodes if not n.virtual]
    span = max(max(xs) - min(xs), max(ys) - min(ys))
    # 세로 A4 폭 210mm 중 망이 쓰는 몫으로 대략 배율을 낸다.
    per_unit = (210.0 * 0.86) / 25.4 * 72.0 / span
    print(f"{stem}: 모델폭 {span:.0f}  배율 대략 {per_unit:.5f} pt/단위")
    print(f"    배관 {R._PIPE_WIDTH_UNITS * per_unit:.3f} pt  "
          f"노드점 {R._NODE_DOT_UNITS * per_unit:.3f} pt  "
          f"기기 {R._DEVICE_PT:.2f} pt = 노드점의 "
          f"{R._DEVICE_PT / (R._NODE_DOT_UNITS * per_unit):.2f} 배")
    print(f"    기기연결선 = 배관의 {1.4 / (R._PIPE_WIDTH_UNITS * per_unit):.1f} 배  "
          f"지시선 = 배관의 {0.25 / (R._PIPE_WIDTH_UNITS * per_unit):.1f} 배")
