# -*- coding: utf-8 -*-
"""「이 레이어를 배관으로 취급」이 실제로 먹히는가 — B1F 셔터선으로 확인.

수동 차선은 색으로 찍어 재료를 확정하는 길이 있는데, 자동에는 없었다. 이름
사전이 OTHER 로 떨어뜨리면 그것으로 끝이라 사람이 손댈 수가 없었다.

여기서 보는 것:
  ① 지정 전 — `현장조사#셔터` 는 OTHER 이고 그래프에 없다
  ② 지정 후 — PIPE 로 올라가 그래프 연장이 는다
  ③ 파스 캐시가 오염되지 않는다 (지정은 사본에만 적용된다)
  ④ 지정을 지우면 원래대로 돌아온다

    python scripts/_verify_pipe_override.py [도면.dxf]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"
FAILS: list[str] = []


def check(label, cond, detail=""):
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return bool(cond)


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--layer", default="현장조사#셔터")
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import remote30_prototype as A
    from routes.module_f.auto import (FORCED_PIPE_SUFFIX, apply_pipe_overrides)

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    cat = {}
    for nm in {str(e.get("l") or "0") for e in ents}:
        try:
            cat[nm] = A._categorize_layer(nm)
        except Exception:  # noqa: BLE001
            cat[nm] = "OTHER"

    print(f"{dxf.name}\n")
    tgt = [e for e in ents if str(e.get("l") or "0") == a.layer]
    check(f"«{a.layer}» 가 도면에 있다", bool(tgt), f"entity {len(tgt)}")
    if not tgt:
        return 1
    check("지정 전에는 배관이 아니다", cat.get(a.layer) != "PIPE",
          str(cat.get(a.layer)))

    colors = {(tuple(e["c"]) if isinstance(e.get("c"), list) else e.get("c"))
              for e in tgt}
    picks = [{"layer": a.layer, "color": c} for c in colors]
    print(f"  지정 묶음 {len(picks)}개 (색 {sorted(map(str, colors))})\n")

    ents2, cat2 = apply_pipe_overrides(ents, cat, picks)
    new = a.layer + FORCED_PIPE_SUFFIX
    check("지정하면 PIPE 로 올라간다", cat2.get(new) == "PIPE",
          f"{new} → {cat2.get(new)}")
    moved = sum(1 for e in ents2 if str(e.get("l") or "0") == new)
    check("entity 가 파생 레이어로 옮겨진다", moved == len(tgt),
          f"{moved}/{len(tgt)}")

    # ★캐시 오염 — 원본 목록은 손대지 않아야 한다.
    still = sum(1 for e in ents if str(e.get("l") or "0") == a.layer)
    check("원본(파스 캐시)은 그대로다", still == len(tgt),
          f"원본에 {still}개 남음")
    check("원본 분류표도 그대로다", new not in cat)

    # 그래프가 실제로 커지는가.
    def net(es, ct):
        got: dict = {}
        real = A._join_head_gap_endpoints

        def grab(graph, edge_len, head_pts, *ar, **kw):
            n = real(graph, edge_len, head_pts, *ar, **kw)
            got["nodes"] = len(graph)
            got["edges"] = len(edge_len)
            got["len_m"] = sum(edge_len.values()) / 1000.0
            return n

        A._join_head_gap_endpoints = grab
        try:
            from remote30_graph import HeadRegion
            hs = A.detect_heads(es, ct)
            pts = [(h.pos[0], h.pos[1]) for h in hs]
            sheet = A.sheet_frame_at(pts)
            ins = pts
            if sheet is not None:
                x0, y0, x1, y1 = [float(v) for v in sheet["bbox"]]
                ins = [q for q in pts if x0 <= q[0] <= x1 and y0 <= q[1] <= y1]
            cx = sum(q[0] for q in ins) / len(ins)
            cy = sum(q[1] for q in ins) / len(ins)
            al = min(ins, key=lambda q: (q[0] - cx) ** 2 + (q[1] - cy) ** 2)
            zs = A.head_bbox_for_region(pts, al)
            au: dict = {}
            A.select_worst30_heads_anchored(
                pipe_entities=es, layer_categories=ct, alarm_xy=al,
                head_region=HeadRegion.from_rects(zs), zones=zs, k=30,
                audit_out=au)
            got["wet"] = (au.get("source_attach") or {}).get("comp_head_count")
        finally:
            A._join_head_gap_endpoints = real
        return got

    print()
    before = net(ents, cat)
    after = net(ents2, cat2)
    print(f"  지정 전 — 절점 {before['nodes']:,} · 간선 {before['edges']:,} · "
          f"연장 {before['len_m']:,.0f} m · 물닿음 {before.get('wet')}")
    print(f"  지정 후 — 절점 {after['nodes']:,} · 간선 {after['edges']:,} · "
          f"연장 {after['len_m']:,.0f} m · 물닿음 {after.get('wet')}")
    check("지정한 선이 그래프에 들어온다",
          after["edges"] > before["edges"],
          f"간선 {before['edges']:,} → {after['edges']:,}")

    # 지정을 지우면 원래대로.
    ents3, cat3 = apply_pipe_overrides(ents, cat, [])
    check("지정을 비우면 원래 그대로다",
          ents3 is ents and cat3 is cat)

    print("\n" + "=" * 52)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        return 1
    print("배관 지정 — 전 항목 PASS")
    return 0


if __name__ == "__main__":
    sys.exit(main())
