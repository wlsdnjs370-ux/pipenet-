# -*- coding: utf-8 -*-
"""평행 ladder 합성이 주차장 도면에서 제대로 도는가.

`collapse_parallel_ladders` 는 배관을 두 평행선으로 그린 관례를 중심선 하나로
합성한다. 안 하면 «시각적으로 꼬임/겹침» 이 생긴다(그 함수 도크스트링의 말).

상수에 「단위세대 도면 기준」이라는 주석이 붙어 있다:

    LADDER_MAX_RUNG_MM = 300.0

주차장은 구획 폭이 2,300~2,500mm 라 스케일이 다르다. 그래서 실제로 재 본다 —
합성이 몇 건이나 걸리고, 문턱을 올리면 얼마나 더 걸리는가.

    python scripts/_probe_ladder.py [도면.dxf ...]
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
GRAB: dict = {}


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A
    from remote30_graph import HeadRegion

    args = sys.argv[1:] or [
        str(ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"),
        str(ROOT / "data" / "_genz_dxf" / "S2_죽전_지하주차장소화설비.dxf"),
        str(ROOT / "routes" / "제출용[최종]"
            / "1. 입력도면 대명동 단위세대 평면도.dxf"),
    ]
    real = A.collapse_parallel_ladders

    print(f"현재 상수 — rung ≤ {A.LADDER_MAX_RUNG_MM:.0f}mm · "
          f"rail/rung ≥ {A.LADDER_MIN_RAIL_RATIO} · "
          f"평행 cos ≥ {A.LADDER_PARALLEL_COS}\n")

    for f in args:
        dxf = Path(f)
        if not dxf.is_file():
            print(f"■ {dxf.name} — 파일 없음\n")
            continue

        for rung in (A.LADDER_MAX_RUNG_MM, 600.0, 1200.0, 2500.0):
            GRAB.clear()

            def hook(graph, edge_len, max_rung_mm=None, *a, _r=rung, **kw):
                # 문턱만 갈아끼우고 나머지는 그대로 — 계측이 규칙을 바꾸면 안 된다.
                n = real(graph, edge_len, _r, *a, **kw)
                GRAB["n"] = GRAB.get("n", 0) + n
                GRAB["nodes"] = len(graph)
                GRAB["edges"] = len(edge_len)
                return n

            A.collapse_parallel_ladders = hook
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
                if not pts:
                    print(f"■ {dxf.name} — 헤드 0\n")
                    break
                sheet = A.sheet_frame_at(pts)
                ins = pts
                if sheet is not None:
                    x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
                    ins = [q for q in pts
                           if x0 <= q[0] <= x1 and y0 <= q[1] <= y1]
                cx = sum(q[0] for q in ins) / len(ins)
                cy = sum(q[1] for q in ins) / len(ins)
                al = min(ins, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
                zs = A.head_bbox_for_region(pts, al)
                au: dict = {}
                sel = A.select_worst30_heads_anchored(
                    pipe_entities=ents, layer_categories=cat, alarm_xy=al,
                    head_region=HeadRegion.from_rects(zs), zones=zs, k=30,
                    audit_out=au)
            except Exception as exc:  # noqa: BLE001
                print(f"  rung {rung:>6.0f} — 실패 {type(exc).__name__}: {exc}")
                continue
            finally:
                A.collapse_parallel_ladders = real

            ed = list(getattr(sel, "edges", None) or ())
            hs = getattr(sel, "heads", None) or ()
            wet = (au.get("source_attach") or {}).get("comp_head_count")
            if rung == A.LADDER_MAX_RUNG_MM:
                print(f"■ {dxf.name}")
            tag = " ← 현재값" if rung == A.LADDER_MAX_RUNG_MM else ""
            print(f"  rung ≤{rung:>6.0f}mm  합성 {GRAB.get('n', 0):>5}건 · "
                  f"그래프 절점 {GRAB.get('nodes', 0):>6,} 간선 "
                  f"{GRAB.get('edges', 0):>6,} · 뽑힌 배관 {len(ed):>4} · "
                  f"헤드 {len(hs):>3} · 물닿음 {wet}{tag}")
        print()
    print("  합성 건수가 문턱을 올려도 안 늘면 이 도면엔 이중선 배관이 없다 —")
    print("  그러면 «꼬임» 의 원인은 ladder 가 아니다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
