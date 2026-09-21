# -*- coding: utf-8 -*-
"""원본 B1F 전체 추출(worst-K subgraph) vs 참조 최소망 기하 비교.

현재 추출이 참조처럼 국소 클러스터로 수렴하는가, 아니면 플로어 전체로 뻗는가?
worst-K subgraph edge 들의 bbox·총연장·헤드 위치를 참조와 비교한다.
ASCII-only.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
import remote30_prototype as R  # noqa: E402

ORIG = BASE / "samples" / "dxf" / "B1F 현장조사 소화설비 평면도.dxf"
REF = BASE / "B1F 현장조사 소화설비 평면도_도면정리_최소.dxf"
ALARM = (661506.0, 177357.0)


def bbox(pts):
    if not pts:
        return None
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return (min(xs), min(ys), max(xs), max(ys))


def run_select(dxf, k, alarm=None):
    bundle = R.parse_dxf_bundle_cached(dxf)
    lc = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pe = R.filter_pipenet_only(bundle)
    kw = {}
    if alarm is not None:
        kw["manual_source"] = alarm
    return R.select_worst30_heads(pe, lc, k=k, **kw)


def summarize(tag, res):
    heads = res.heads
    edges = res.edges  # set of (a,b) node-pair keys?
    hpts = [h.pos for h in heads]
    hbb = bbox(hpts)
    # edges: iterate
    epts = []
    tot = 0.0
    n_e = 0
    for e in edges:
        try:
            a, b = e
            epts.append(a); epts.append(b)
            tot += math.hypot(a[0] - b[0], a[1] - b[1])
            n_e += 1
        except Exception:
            pass
    ebb = bbox(epts)
    print(f"[{tag}] heads={len(heads)} edges={n_e} tot_len={tot:,.0f}", flush=True)
    if hbb:
        print(f"   head bbox=({hbb[0]:.0f},{hbb[1]:.0f},{hbb[2]:.0f},{hbb[3]:.0f}) "
              f"size {hbb[2]-hbb[0]:.0f}x{hbb[3]-hbb[1]:.0f}", flush=True)
    if ebb:
        print(f"   edge bbox=({ebb[0]:.0f},{ebb[1]:.0f},{ebb[2]:.0f},{ebb[3]:.0f}) "
              f"size {ebb[2]-ebb[0]:.0f}x{ebb[3]-ebb[1]:.0f}", flush=True)
    print(f"   source_kind={res.source_kind} source_pos={res.source_pos} "
          f"fallback={res.source_fallback}", flush=True)


def main():
    print("REF network geometry (target):", flush=True)
    print("   REF 가지관 bbox size ~ 56162x13401 (국소 클러스터)\n", flush=True)
    for k in (30, 115):
        print(f"--- ORIG select k={k}, alarm={ALARM} ---", flush=True)
        res = run_select(ORIG, k, ALARM)
        summarize(f"ORIG k={k}", res)
        print(flush=True)
    print("--- REF select k=115 (auto source) ---", flush=True)
    resr = run_select(REF, 115, None)
    summarize("REF k=115", resr)
    print("\nDONE", flush=True)


if __name__ == "__main__":
    main()
