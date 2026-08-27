# -*- coding: utf-8 -*-
"""모듈 A 의 평면도 추출을 그대로 돌려 산출을 들여다본다.

「A 가 이상하게 작동한다」를 코드 변경 탓으로 몰기 전에, 지금 실제로 무엇이
나오는지부터 본다. A 는 이번 작업에서 한 줄도 안 건드렸다(diff 로 확인) —
그러니 여기서 이상한 것이 나오면 그것은 종전부터 그랬거나, 도면·입력 쪽이다.

두 경로를 나란히 돌린다:
    select_worst30_heads           비-anchored (A 의 옛 경로)
    select_worst30_heads_anchored  2앵커 (F 의 자동이 쓰는 그 경로)

    python scripts/_probe_a_plan.py [도면.dxf]
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = ROOT / "samples" / "dxf" / "LH306동_평면도.dxf"


def _stat(name, sel, heads_all):
    hs = getattr(sel, "heads", None) or ()
    ed = getattr(sel, "edges", None) or ()
    nodes = getattr(sel, "nodes_in_subgraph", None) or ()
    dists = [float(d) for d in (getattr(sel, "distances", None) or ())]
    total_m = sum(float(e[2]) for e in ed) / 1000.0 if ed else 0.0
    print(f"\n  [{name}]")
    print(f"    선정 헤드   {len(hs)} / 검출 {heads_all}")
    print(f"    절점·배관   {len(nodes)} · {len(ed)}")
    print(f"    망 연장     {total_m:.2f} m")
    if dists:
        print(f"    유하거리    최원 {max(dists)/1000:.2f} m · "
              f"최근 {min(dists)/1000:.2f} m")
    src = getattr(sel, "source_pos", None)
    print(f"    급수원      {src}")
    bd = float(getattr(sel, "source_bridge_dist_mm", 0.0) or 0.0)
    fb = bool(getattr(sel, "source_fallback", False))
    print(f"    급수원 결합 {bd:.0f} mm" + ("  ★최근접으로 대체됨" if fb else ""))
    # 헤드가 한 구역에 뭉쳤나 — 흩어지면 설계면적이 성립하지 않는다.
    if hs:
        xs = [h.pos[0] for h in hs]
        ys = [h.pos[1] for h in hs]
        import math
        diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) / 1000.0
        print(f"    헤드 퍼짐   대각 {diag:.1f} m"
              + ("   ★흩어졌다(설계면적 아님)" if diag > 60 else ""))


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

    print(f"모듈 A 평면 추출 — {dxf.name} "
          f"({dxf.stat().st_size / 1024 / 1024:.1f} MB)")

    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    names = {str(e.get("l") or "0") for e in ents}
    cat = {}
    for n in names:
        try:
            cat[n] = A._categorize_layer(n)
        except Exception:  # noqa: BLE001
            cat[n] = "OTHER"
    counts: dict = {}
    for c in cat.values():
        counts[c] = counts.get(c, 0) + 1
    print(f"  도형 {len(ents):,} · 레이어 {len(names)} · "
          + " · ".join(f"{k} {v}" for k, v in sorted(counts.items())))

    heads = A.detect_heads(ents, cat)
    print(f"  헤드 검출 {len(heads):,}개")
    if not heads:
        print("  ★헤드 0개 — 여기서 끝난다. 레이어 분류나 XREF 를 보라.")
        print(f"    XREF 진단: {bundle.xref_diagnostics}")
        return 1
    xs = [h.pos[0] for h in heads]
    ys = [h.pos[1] for h in heads]
    print(f"  헤드 범위 x {min(xs):.0f}~{max(xs):.0f} · y {min(ys):.0f}~{max(ys):.0f}")

    # 알람밸브 = 헤드 무리 한가운데 헤드(배관 위라 결합이 된다)
    cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
    mid = min(heads, key=lambda h: (h.pos[0] - cx) ** 2 + (h.pos[1] - cy) ** 2)
    alarm = (mid.pos[0], mid.pos[1])
    pad = 1000.0
    region = HeadRegion.from_rects(
        [(min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad)])

    try:
        sel_a = A.select_worst30_heads_anchored(
            ents, cat, alarm, region, k=30)
        _stat("anchored (F 자동이 쓰는 경로)", sel_a, len(heads))
    except Exception as exc:  # noqa: BLE001
        print(f"\n  [anchored] 실패 — {type(exc).__name__}: {exc}")

    try:
        sel_b = A.select_worst30_heads(
            pipe_entities=ents, layer_categories=cat, k=30)
        _stat("비-anchored (A 옛 경로)", sel_b, len(heads))
    except Exception as exc:  # noqa: BLE001
        print(f"\n  [비-anchored] 실패 — {type(exc).__name__}: {exc}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
