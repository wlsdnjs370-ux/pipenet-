# -*- coding: utf-8 -*-
"""D2 코퍼스 회귀 — 262 짝 전량 결합 (일회성)."""
from __future__ import annotations

import re
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.d_result_binder import bind_results  # noqa: E402

LIB = ROOT / "data" / "reference_library"
SUB = ROOT / "routes" / "제출용[최종]"

pairs = [(s, s.with_suffix(".xml")) for s in LIB.rglob("*.sdf") if s.with_suffix(".xml").exists()]
pairs += [(p, p.with_suffix(".xml")) for p in SUB.glob("*.sdf") if p.with_suffix(".xml").exists()]

t0 = time.time()
errors, join, totals = [], Counter(), Counter()
datum, warns, factors = Counter(), Counter(), {}
for sdf, xml in pairs:
    try:
        b = bind_results(sdf, xml)
    except Exception as exc:
        errors.append((sdf.name, repr(exc)))
        continue
    join["일치" if b.report.matched else "불일치"] += 1
    totals["nodes"] += len(b.nodes)
    totals["pipes"] += len(b.pipes)
    totals["nozzles"] += len(b.nozzles)
    d = b.datum_offset_pa
    datum[None if d is None else round(d, 1)] += 1
    for w in b.warnings:
        warns[re.sub(r"\d+", "N", w)] += 1
    for k, s in b.scales.items():
        if s.factor is not None:
            factors.setdefault(k, Counter())[(s.factor, s.source)] += 1

print(f"{len(pairs)} 짝 / {time.time() - t0:.1f}s / 예외 {len(errors)} {errors[:3]}")
print("조인", dict(join), "| 결합량", dict(totals))
print("기준면", datum.most_common())
print("배율:")
for k in sorted(factors):
    print(f"  {k:20s}", [(f"{f:.6g}", src, n) for (f, src), n in factors[k].most_common()])
print("경고:")
for w, n in warns.most_common():
    print(f"  {n:5d}  {w}")
