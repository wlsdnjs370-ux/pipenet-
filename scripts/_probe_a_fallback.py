# -*- coding: utf-8 -*-
"""A 화면에서 «영역을 안 그렸을 때» 무엇이 달라지는지 잰다.

A 는 영역이 없으면 조용히 옛 알고리즘(select_worst30_heads)으로 떨어졌다.
그 경로가 실제로 무엇을 내놓는지, 그리고 영역을 검출에서 만들어 앵커 경로로
보내면 무엇이 되는지 나란히 놓는다.

    python scripts/_probe_a_fallback.py [도면.dxf]
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def _stat(name, sel):
    hs = getattr(sel, "heads", None) or ()
    ed = getattr(sel, "edges", None) or ()
    ds = [float(d) for d in (getattr(sel, "distances", None) or ())]
    xs = [h.pos[0] for h in hs]
    ys = [h.pos[1] for h in hs]
    diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) / 1000.0 if hs else 0.0
    bd = float(getattr(sel, "source_bridge_dist_mm", 0.0) or 0.0)
    print(f"  [{name}]")
    print(f"    헤드 {len(hs)} · 배관 {len(ed)} · "
          f"연장 {sum(float(e[2]) for e in ed) / 1000:.2f} m")
    if ds:
        print(f"    유하거리 최원 {max(ds) / 1000:.2f} m · 최근 {min(ds) / 1000:.2f} m")
    print(f"    급수원 결합 {bd:,.0f} mm"
          + ("   ★멀다(억지로 붙임)" if bd > 1000 else ""))
    print(f"    헤드 퍼짐 대각 {diag:.1f} m"
          + ("   ★흩어졌다(설계면적 아님)" if diag > 60 else "   (뭉쳤다)"))


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    import remote30_prototype as A
    from remote30_graph import HeadRegion

    ents = A.parse_dxf_bundle_cached(dxf).entities
    cat = {}
    for n in {str(e.get("l") or "0") for e in ents}:
        try:
            cat[n] = A._categorize_layer(n)
        except Exception:  # noqa: BLE001
            cat[n] = "OTHER"

    heads = A.detect_heads(ents, cat)
    pts = [(h.pos[0], h.pos[1]) for h in heads]
    print(f"{dxf.name} · 헤드 {len(heads):,}개\n")

    # 알람밸브 = 실제 도면 한가운데 헤드 (사람이 찍었을 법한 자리)
    sheet = A.sheet_frame_at(pts)
    inside = pts
    if sheet is not None:
        x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
        inside = [p for p in pts if x0 <= p[0] <= x1 and y0 <= p[1] <= y1]
    cx = sum(p[0] for p in inside) / len(inside)
    cy = sum(p[1] for p in inside) / len(inside)
    alarm = min(inside, key=lambda p: (p[0] - cx) ** 2 + (p[1] - cy) ** 2)
    print(f"  알람밸브 {alarm[0]:.0f}, {alarm[1]:.0f}\n")

    print("  ── 고친 전: 영역이 없으면 옛 경로 ──")
    try:
        _stat("select_worst30_heads (비-anchored)",
              A.select_worst30_heads(pipe_entities=ents, layer_categories=cat,
                                     manual_source=alarm, k=30))
    except Exception as exc:  # noqa: BLE001
        print(f"    실패 — {type(exc).__name__}: {exc}")

    print("\n  ── 고친 후: 검출에서 범위를 만들어 앵커 경로 ──")
    zones = A.head_bbox_for_region(pts, alarm)
    w = (zones[0][2] - zones[0][0]) / 1000
    h = (zones[0][3] - zones[0][1]) / 1000
    print(f"    만든 범위 {w:,.0f} x {h:,.0f} m")
    try:
        _stat("select_worst30_heads_anchored",
              A.select_worst30_heads_anchored(
                  pipe_entities=ents, layer_categories=cat, alarm_xy=alarm,
                  head_region=HeadRegion.from_rects(zones), zones=zones, k=30))
    except Exception as exc:  # noqa: BLE001
        print(f"    실패 — {type(exc).__name__}: {exc}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
