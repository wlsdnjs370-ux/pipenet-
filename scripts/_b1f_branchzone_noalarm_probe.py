# -*- coding: utf-8 -*-
"""분기영역 검증 — alarm_xy 미지정(수동 알람밸브 없음) 케이스.

버그: viz 경로는 alarm_xy 없으면 spt_source=None → 영역제한 no-op.
수정 후엔 자동 source 로 루팅돼 영역 밖 junction 이 급감해야 한다.
"""
from __future__ import annotations
import sys
from pathlib import Path
BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
import remote30_prototype as R  # noqa: E402

ORIG = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"
RECT = (273900.0, 74450.0, 328050.0, 150350.0)


def run(bz, alarm):
    summary = None; loops = 0; edges = 0
    for evt in R.run_stages_0_2(ORIG, "smoke", alarm_xy=alarm, branch_zones=bz):
        if evt.get("type") == "entities" and evt.get("stage") == 3:
            summary = evt["summary"]
            for en in evt["entities"]:
                if en.get("l") == "_graph_loop":
                    loops += 1
                elif en.get("l") == "_graph_edge":
                    edges += 1
    return summary, loops, edges


def main():
    s0, l0, e0 = run(None, None)
    print("NO   bz / NO alarm : junction=%d real_edge=%d removed_cycle=%d loop=%d"
          % (s0["junction_count"], e0, s0["removed_cycle_edges"], l0), flush=True)
    s1, l1, e1 = run([RECT], None)
    print("WITH bz / NO alarm : junction=%d real_edge=%d removed_cycle=%d loop=%d loop_summary=%d source_kind=%s"
          % (s1["junction_count"], e1, s1["removed_cycle_edges"], l1,
             s1.get("loop_edge_count", -1), s1.get("source_kind")), flush=True)
    print("DONE", flush=True)


if __name__ == "__main__":
    main()
