# -*- coding: utf-8 -*-
"""Stage 2(헤드 인식) 실측 프로파일 — 큰 도면에서 어디에 시간이 가는지."""
from __future__ import annotations

import cProfile
import pstats
import sys
import time
from io import StringIO
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.stdout.reconfigure(encoding="utf-8")

import remote30_prototype as rp  # noqa: E402

TARGET = Path(sys.argv[1]) if len(sys.argv) > 1 else BASE / "samples/dxf/LH 지하층배관도.dxf"

t = time.perf_counter()
bundle = rp.parse_dxf_bundle_cached(TARGET)
print(f"parse            {time.perf_counter()-t:7.2f}s  entity {len(bundle.entities):,}")

layer_cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
t = time.perf_counter()
pipe_ents = rp.filter_pipenet_only(bundle)
print(f"filter_pipenet   {time.perf_counter()-t:7.2f}s  entity {len(pipe_ents):,}")

n_h = sum(1 for e in pipe_ents if e["t"] == "H")
n_head_layer = sum(1 for e in pipe_ents if layer_cat.get(e.get("l", "")) == "HEAD")
print(f"  HATCH {n_h:,} · HEAD 레이어 entity {n_head_layer:,}")

t = time.perf_counter()
heads = rp.detect_heads(pipe_ents, layer_cat)
el = time.perf_counter() - t
print(f"detect_heads     {el:7.2f}s  head {len(heads):,}")

pr = cProfile.Profile()
pr.enable()
rp.detect_heads(pipe_ents, layer_cat)
pr.disable()
s = StringIO()
pstats.Stats(pr, stream=s).sort_stats("tottime").print_stats(14)
print("\n".join(s.getvalue().splitlines()[4:26]))
