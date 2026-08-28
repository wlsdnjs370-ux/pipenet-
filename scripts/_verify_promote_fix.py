# -*- coding: utf-8 -*-
"""승격 분류 수정의 효과 — 헤드 검출과 최불리 추출에서 무엇이 달라지나.

`routes/module_f/auto.py::open_plan` 과 `routes/module_f/recon.py::run_recon`
은 번들이 이미 매긴 `auto_category` 를 버리고 레이어 «이름» 으로 다시 분류하고
있었다. 승격은 entity 를 파생 레이어로 옮기는 방식이라, 파생 이름이 이름
사전에서 PIPE 로 안 떨어지면 통째로 증발한다. 실측:

    "<원이름> (배관 승격)"   → PIPE    (우연히 살아 있었다)
    "<원이름> (연결관 승격)" → OTHER   ★증발

여기서는 두 분류를 나란히 놓고 **결과가 실제로 달라지는지** 잰다. 분류만 다르고
헤드·배관이 그대로면 고칠 값어치가 없는 수정이다.

    python scripts/_verify_promote_fix.py [도면.dxf ...]
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf",
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf",
    ROOT / "samples" / "dxf" / "계통도_LH_306.dxf",
]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A
    from remote30_graph import HeadRegion

    args = [Path(a) for a in sys.argv[1:]] or DEFAULT
    for dxf in args:
        if not dxf.is_file():
            print(f"■ {dxf.name} — 파일 없음\n")
            continue
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = list(bundle.entities or ())

        # 고친 뒤 — 번들 분류 그대로
        after = {str(ly.get("name")): str(ly.get("auto_category") or "OTHER")
                 for ly in (bundle.layers or ())}
        # 고치기 전 — 이름으로 재분류
        before = {}
        for n in after:
            try:
                before[n] = A._categorize_layer(n)
            except Exception:  # noqa: BLE001
                before[n] = "OTHER"

        print(f"■ {dxf.name}")
        if before == after:
            print("    분류가 같다 — 이 도면은 이 수정과 무관\n")
            continue

        rows = []
        for tag, cat in (("고치기 전", before), ("고친 뒤", after)):
            heads = A.detect_heads(ents, cat)
            pts = [(h.pos[0], h.pos[1]) for h in heads]
            n_pipe = sum(1 for e in ents
                         if cat.get(str(e.get("l") or "0"), "OTHER") == "PIPE")
            sel = None
            if pts:
                sheet = A.sheet_frame_at(pts)
                ins = pts
                if sheet is not None:
                    x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
                    ins = [q for q in pts
                           if x0 <= q[0] <= x1 and y0 <= q[1] <= y1] or pts
                cx = sum(q[0] for q in ins) / len(ins)
                cy = sum(q[1] for q in ins) / len(ins)
                al = min(ins, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
                zs = A.head_bbox_for_region(pts, al)
                try:
                    sel = A.select_worst30_heads_anchored(
                        pipe_entities=ents, layer_categories=cat, alarm_xy=al,
                        head_region=HeadRegion.from_rects(zs), zones=zs, k=30)
                except Exception as exc:  # noqa: BLE001
                    print(f"    {tag} 추출 실패 — {type(exc).__name__}: {exc}")
            n_sel = len(getattr(sel, "heads", None) or ()) if sel else 0
            n_edge = len(getattr(sel, "edges", None) or ()) if sel else 0
            rows.append((tag, n_pipe, len(heads), n_sel, n_edge))

        print(f"    {'':<10} {'PIPE ent':>9} {'헤드 검출':>9} "
              f"{'뽑힌 헤드':>9} {'뽑힌 배관':>9}")
        print("    " + "-" * 50)
        for tag, a_, b_, c_, d_ in rows:
            print(f"    {tag:<10} {a_:>9,} {b_:>9,} {c_:>9,} {d_:>9,}")
        (_, p0, h0, s0, e0), (_, p1, h1, s1, e1) = rows
        print(f"    {'차이':<10} {p1 - p0:>+9,} {h1 - h0:>+9,} "
              f"{s1 - s0:>+9,} {e1 - e0:>+9,}")
        print()

    print("  전부 0 이면 분류만 달라지고 결과는 같다 — 그래도 옳은 분류지만,")
    print("  «대대적 조치» 는 아니다. 그 점을 숨기지 말 것.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
