# -*- coding: utf-8 -*-
"""자동의 «기본 범위» — 장을 가르기 전과 후를 나란히 잰다.

한 파일에 도면이 여러 장이면, 검출한 헤드 전부를 범위로 삼는 것이 틀린다.
도면 밖 이상점이 범위를 부풀려 최불리가 엉뚱한 곳을 후보로 삼는다.

    python scripts/_probe_auto_region.py [도면.dxf]
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def _run(A, ents, cat, alarm, rects, name):
    from remote30_graph import HeadRegion
    try:
        sel = A.select_worst30_heads_anchored(
            ents, cat, alarm,
            HeadRegion.from_rects([tuple(float(v) for v in r) for r in rects]),
            k=30)
    except Exception as exc:  # noqa: BLE001
        print(f"  [{name}] 실패 — {type(exc).__name__}: {exc}")
        return
    hs = getattr(sel, "heads", None) or ()
    ed = getattr(sel, "edges", None) or ()
    ds = [float(d) for d in (getattr(sel, "distances", None) or ())]
    xs = [h.pos[0] for h in hs]
    ys = [h.pos[1] for h in hs]
    diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) / 1000.0 if hs else 0
    print(f"  [{name}]")
    print(f"    헤드 {len(hs)} · 배관 {len(ed)} · "
          f"연장 {sum(float(e[2]) for e in ed) / 1000:.2f} m")
    print(f"    유하거리 최원 {max(ds) / 1000:.2f} m · 최근 {min(ds) / 1000:.2f} m"
          if ds else "    유하거리 —")
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
    from routes.module_f.auto import region_around, sheet_of

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

    # 알람밸브 = 실제 도면(헤드가 가장 많은 장) 한가운데 헤드
    sheet = sheet_of(pts)
    inside = pts
    if sheet is not None:
        x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
        inside = [p for p in pts if x0 <= p[0] <= x1 and y0 <= p[1] <= y1]
    cx = sum(p[0] for p in inside) / len(inside)
    cy = sum(p[1] for p in inside) / len(inside)
    alarm = min(inside, key=lambda p: (p[0] - cx) ** 2 + (p[1] - cy) ** 2)
    print(f"  알람밸브(모의) {alarm[0]:.0f}, {alarm[1]:.0f}\n")

    axs = [p[0] for p in pts]
    ays = [p[1] for p in pts]
    pad = 1000.0
    old = [[min(axs) - pad, min(ays) - pad, max(axs) + pad, max(ays) + pad]]
    print(f"  고친 전 범위 {(old[0][2] - old[0][0]) / 1000:,.0f} x "
          f"{(old[0][3] - old[0][1]) / 1000:,.0f} m")
    _run(A, ents, cat, alarm, old, "고친 전 — 헤드 전부의 bbox")

    new = region_around([{"x": p[0], "y": p[1]} for p in pts], alarm)
    print(f"\n  고친 후 범위 {(new[0][2] - new[0][0]) / 1000:,.0f} x "
          f"{(new[0][3] - new[0][1]) / 1000:,.0f} m")
    _run(A, ents, cat, alarm, new, "고친 후 — 알람밸브가 놓인 장")
    return 0


if __name__ == "__main__":
    sys.exit(main())
