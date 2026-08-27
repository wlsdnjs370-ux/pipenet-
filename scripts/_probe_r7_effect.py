# -*- coding: utf-8 -*-
"""R7 상한을 올리면 «추출 결과» 가 어떻게 되나 — 여러 도면에서 한 번에.

상한을 올려서 접속이 몇 건 늘었나까지는 `_probe_r7_gate.py` 가 잰다. 여기서는
그 다음을 본다: **헤드가 실제로 더 붙는가, 망이 어떻게 변하는가.**

한 도면이라도 나빠지면 그 값은 못 쓴다 — 고치려는 도면만 보고 정하면 다른
현장에서 터진다.

    python scripts/_probe_r7_effect.py [--limits 400,700,800,1000]
"""
from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

DRAWINGS = [
    ("B1F 원본", ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf", 30),
    ("대명동", ROOT / "routes" / "제출용[최종]"
     / "1. 입력도면 대명동 단위세대 평면도.dxf", 10),
    ("LH306", ROOT / "samples" / "dxf" / "LH306동_평면도.dxf", 30),
    ("죽전 주차장", ROOT / "data" / "_genz_dxf"
     / "S2_죽전_지하주차장소화설비.dxf", 30),
]


def run_one(A, HeadRegion, dxf: Path, k: int, limit: float) -> dict:
    """상한 하나로 한 도면을 뽑고 요약을 낸다."""
    real = A._join_head_gap_endpoints
    box: dict = {}

    def hook(graph, edge_len, head_pts, *args, **kw):
        joins: list = []
        n = real(graph, edge_len, head_pts, limit,
                 A.HEAD_GAP_JOIN_TOL_MM, joins)
        box["joins"] = n
        box["gaps"] = sorted(float(j["gap_mm"]) for j in joins)
        return n

    A._join_head_gap_endpoints = hook
    try:
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = bundle.entities
        cat = {}
        for nm in {str(e.get("l") or "0") for e in ents}:
            try:
                cat[nm] = A._categorize_layer(nm)
            except Exception:  # noqa: BLE001
                cat[nm] = "OTHER"
        heads = A.detect_heads(ents, cat)
        pts = [(h.pos[0], h.pos[1]) for h in heads]
        sheet = A.sheet_frame_at(pts)
        inside = pts
        if sheet is not None:
            x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
            inside = [q for q in pts if x0 <= q[0] <= x1 and y0 <= q[1] <= y1]
        cx = sum(q[0] for q in inside) / len(inside)
        cy = sum(q[1] for q in inside) / len(inside)
        alarm = min(inside, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
        zones = A.head_bbox_for_region(pts, alarm)
        audit: dict = {}
        sel = A.select_worst30_heads_anchored(
            pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
            head_region=HeadRegion.from_rects(zones), zones=zones,
            k=k, audit_out=audit)
    except Exception as exc:  # noqa: BLE001
        return {"error": f"{type(exc).__name__}: {exc}"}
    finally:
        A._join_head_gap_endpoints = real

    ed = list(getattr(sel, "edges", None) or ())
    hs = getattr(sel, "heads", None) or ()
    ds = [float(d) for d in (getattr(sel, "distances", None) or ())]
    xs = [h.pos[0] for h in hs]
    ys = [h.pos[1] for h in hs]
    fr = audit.get("fragments") or {}
    sa = audit.get("source_attach") or {}
    uh = (audit.get("heads") or {}).get("unreachable") or []
    g = box.get("gaps") or []
    return {
        "joins": box.get("joins", 0),
        "gap_med": (g[len(g) // 2] if g else 0.0),
        "gap_max": (g[-1] if g else 0.0),
        "wet": sa.get("comp_head_count"),
        "unreach": len(uh),
        "frag": fr.get("count"),
        "k": len(hs),
        "nodes": len({n for e in ed for n in (e[0], e[1])}),
        "pipes": len(ed),
        "len_m": round(sum(float(e[2]) for e in ed) / 1000.0, 1),
        "far_m": round(max(ds) / 1000.0, 2) if ds else 0.0,
        "spread_m": (round(math.hypot(max(xs) - min(xs),
                                      max(ys) - min(ys)) / 1000.0, 1)
                     if hs else 0.0),
    }


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("--limits", default="400,700,800,1000")
    ap.add_argument("--only", default="")
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A
    from remote30_graph import HeadRegion

    limits = [float(s) for s in a.limits.split(",") if s.strip()]
    for label, dxf, k in DRAWINGS:
        if a.only and a.only not in label:
            continue
        if not dxf.is_file():
            print(f"\n■ {label} — 파일 없음 ({dxf.name})")
            continue
        print(f"\n■ {label} · K={k} · {dxf.name}")
        print(f"  {'상한':>6} {'접속':>6} {'틈중앙':>7} {'물닿음':>7} "
              f"{'미도달':>7} {'조각':>6} {'헤드':>5} {'절점':>6} "
              f"{'연장m':>8} {'최원m':>8} {'퍼짐m':>7}")
        print("  " + "-" * 84)
        base = None
        for lim in limits:
            r = run_one(A, HeadRegion, dxf, k, lim)
            if r.get("error"):
                print(f"  {lim:>6.0f}  실패 — {r['error'][:60]}")
                continue
            if base is None:
                base = r
            mark = ""
            if base is not None and r is not base:
                if (r["wet"] or 0) < (base["wet"] or 0):
                    mark = "  ★물닿음 감소"
                elif r["k"] < base["k"]:
                    mark = "  ★선정 헤드 감소"
            print(f"  {lim:>6.0f} {r['joins']:>6,} {r['gap_med']:>7.0f} "
                  f"{(r['wet'] or 0):>7,} {r['unreach']:>7,} "
                  f"{(r['frag'] or 0):>6,} {r['k']:>5} {r['nodes']:>6,} "
                  f"{r['len_m']:>8,.1f} {r['far_m']:>8.2f} "
                  f"{r['spread_m']:>7.1f}{mark}")
    print("\n  물닿음 = 급수원이 닿는 컴포넌트의 헤드 수 (많을수록 좋다)")
    print("  ★ 표시가 하나라도 있으면 그 상한은 못 쓴다 — 다른 도면을 망친다")
    return 0


if __name__ == "__main__":
    sys.exit(main())
