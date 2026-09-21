# -*- coding: utf-8 -*-
"""코퍼스 전체에서 typesize 가 실제로 갈리는 값인지 센다."""
from __future__ import annotations

import collections
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "core")]

from core.d_display_model import load_display_model  # noqa: E402

LIB = ROOT / "data" / "reference_library"

per_text = collections.Counter()
per_file = collections.Counter()
mixed = 0
files = 0
for sdf in sorted(LIB.rglob("*.sdf")):
    try:
        model = load_display_model(sdf)
    except Exception:                                      # noqa: BLE001
        continue
    sizes = [t.typesize for t in model.texts]
    if not sizes:
        continue
    files += 1
    per_text.update(sizes)
    per_file[tuple(sorted(set(sizes)))] += 1
    if len(set(sizes)) > 1:
        mixed += 1

print(f"주기가 있는 SDF {files}개")
print(f"  글자 하나하나의 typesize: {per_text.most_common(8)}")
print(f"  파일 안에서 값이 섞인 파일: {mixed}개")
print(f"  파일별 값 조합: {per_file.most_common(6)}")

for name in ("2. Pipenet_hand.sdf", "2. Pipenet_auto.sdf"):
    p = ROOT / "routes" / "제출용[최종]" / name
    if p.exists():
        m = load_display_model(p)
        print(f"  우리 {name}: {collections.Counter(t.typesize for t in m.texts).most_common()}")
