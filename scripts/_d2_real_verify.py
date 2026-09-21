# -*- coding: utf-8 -*-
"""D2 실물 2 짝 검증 (일회성)."""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from core.d_result_binder import bind_results  # noqa: E402

SUB = ROOT / "routes" / "제출용[최종]"

for stem in ("2. Pipenet_hand", "2. Pipenet_auto"):
    b = bind_results(SUB / f"{stem}.sdf", SUB / f"{stem}.xml")
    print(f"=== {stem} {b.title}")
    print("  배율:", " / ".join(
        f"{k} {s.factor:.6g} {s.source}" for k, s in sorted(b.scales.items()) if s.factor))
    print("  기준면 offset Pa:", b.datum_offset_pa)
    print("  조인:", b.report.summary(),
          f"| 노드 {len(b.nodes)} 관로 {len(b.pipes)} 노즐 {len(b.nozzles)}")
    print("  내용다름:", b.report.divergent_pipes)
    print("  XML 만 노드:", b.report.xml_only_nodes)
    print("  경고:", b.warnings)
    p = next(iter(b.pipes.values()))
    print("  관로:", p)
    if "100" in b.nodes:
        print("  노드 100:", b.nodes["100"])
    z = next(iter(b.nozzles.values()))
    print("  노즐:", z)
