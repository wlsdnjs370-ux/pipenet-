# -*- coding: utf-8 -*-
"""run_stages_0_2 (VIZ path) 분기영역 검증 — _graph_loop 방출 + 단일 corridor."""
from __future__ import annotations
import sys
from pathlib import Path
BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
import remote30_prototype as R  # noqa: E402

ORIG = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"
SRC = (156776.0, 177970.0)
RECT = (273900.0, 74450.0, 328050.0, 150350.0)


def run(bz):
    summary = None; loops = 0; edges = 0
    for evt in R.run_stages_0_2(ORIG, "smoke", alarm_xy=SRC, branch_zones=bz):
        if evt.get("type") == "entities" and evt.get("stage") == 3:
            summary = evt["summary"]
            for en in evt["entities"]:
                if en.get("l") == "_graph_loop":
                    loops += 1
                elif en.get("l") == "_graph_edge":
                    edges += 1
    return summary, loops, edges


def main():
    s0, l0, e0 = run(None)
    print("NO   branch_zone: loop=%d real_edge=%d junction=%d removed_cycle=%d"
          % (l0, e0, s0["junction_count"], s0["removed_cycle_edges"]), flush=True)
    s1, l1, e1 = run([RECT])
    print("WITH branch_zone: loop=%d real_edge=%d junction=%d removed_cycle=%d loop_summary=%d"
          % (l1, e1, s1["junction_count"], s1["removed_cycle_edges"], s1.get("loop_edge_count", -1)),
          flush=True)
    print("DONE", flush=True)


if __name__ == "__main__":
    main()
