# -*- coding: utf-8 -*-
"""완주 원자료(`data/_f11f_lanes.json`)를 사람이 읽는 표로 — 문서에 옮길 재료."""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KEYS = ["mb", "open_s", "cands", "bands", "rule", "conf_min", "adopt_n",
        "head_applied", "ghost", "board", "anchor_clicks", "worst",
        "build1_s", "n_pipes", "unresolved_kind", "unresolved_pairs",
        "fill", "fill_kind", "build2_s", "meta_fit", "meta_bore",
        "missed", "survived", "emit_ok", "sdf", "total_s", "decisions",
        "fail"]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    rows = json.loads((ROOT / "data" / "_f11f_lanes.json")
                      .read_text(encoding="utf-8"))
    for r in rows:
        print(f"\n■ {r.get('file')}")
        for k in KEYS:
            if k in r:
                print(f"   {k:<18} {r[k]}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
