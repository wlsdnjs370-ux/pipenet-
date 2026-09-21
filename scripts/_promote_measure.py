# -*- coding: utf-8 -*-
"""헤드 틈 지문 승격 결과 — 캐시 우회 무-부작용 측정."""
from __future__ import annotations

import sys
import time
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as R  # noqa: E402

DXFS = [
    ("B1F_upload", BASE / "data/uploads/B1F_.dxf"),
    ("B1F_최소", BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"),
    ("대명동", BASE / "samples/dxf/대명동201동 단위세대_layer정리.dxf"),
    ("LH306", BASE / "samples/dxf/LH306동_배관망.dxf"),
    ("LH지하", BASE / "samples/dxf/LH 지하층배관도_배관망.dxf"),
]

for tag, path in DXFS:
    if not path.exists():
        print(f"[{tag}] 파일 없음")
        continue
    t0 = time.time()
    bundle = R.parse_dxf_bundle(path)
    print(f"[{tag}] {time.time()-t0:5.1f}s 승격={bundle.promoted_layers}")
