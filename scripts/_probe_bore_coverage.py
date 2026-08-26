# -*- coding: utf-8 -*-
"""관경 근거 커버리지 실측 — 왜 별표1 폴백이 대부분인가.

특허 S520 은 «담당 헤드 수 기준 법정 최소» 와 «도면 표기 판독값» 중 안전측을
채택하라고 한다. 구현(design/bore.py)은 그 규칙 그대로다. 그런데 B1F 기준선은
텍스트 19 / 폴백 79 로 81%가 폴백이다 — 규칙이 아니라 **커버리지** 문제다.

폴백이 나는 길은 둘뿐이다:
  ㉮ 역참조(edge_ref)가 없다 — 헤드 접속관·가지 상승처럼 도면에 그린 선이 아닌 배관
  ㉯ 역참조는 있는데 1,500mm 안에 치수 텍스트가 없다

둘의 비율과, ㉯ 의 실제 거리 분포를 재서 «한계를 늘리면 얼마나 건지는가» 를 본다.

    python scripts/_probe_bore_coverage.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
EDITOR = ROOT / "cad_project_editor_g"
KEY = "B1F 현장조사 소화설비 평면도"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    # 엔진은 cwd 기준 상대경로("docs/import")로 작업폴더를 찾는다. 모듈 F 의
    # 부팅이 그것을 절대경로로 못박으므로 같은 것을 태운다 — 여기서 직접
    # sys.path 를 만지면 E/G 중 어느 트리가 올라오는지가 순서 우연에 걸린다.
    from routes.module_f.common import _boot
    _boot()

    from services.cad_import.design.bore import (
        DIA_RANGE_LIMIT_MM, decide_bores, extract_dia_text_points,
        match_diameter_for_segment, nfpc_min_bore_mm, source_counts)
    from services.cad_import.design.restrict import select_and_expand
    from services.cad_import.edit.session import EditSession
    from services.cad_import.pipeline import handoff, stage1 as s1

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None
    got = select_and_expand(payload, es.board, k=30, selected_source=sel)
    if not got.get("ok"):
        print("!! 제한 전개 실패:", got.get("error"))
        return 1

    spec = EDITOR / "docs" / "import" / "0단계_새찍기" / f"{KEY}_찍은스펙.json"
    src = json.loads(spec.read_text(encoding="utf-8")).get("source_dxf")
    world = handoff.load_world(KEY, src, s1.World) if src else None
    texts = extract_dia_text_points(world.texts) if world else []

    print(f"도면: {KEY}")
    print(f"  world.texts 총 {len(world.texts) if world else 0:,} 개")
    print(f"  치수 텍스트로 읽힌 것 {len(texts):,} 개  (한계 {DIA_RANGE_LIMIT_MM:.0f}mm)")
    if texts:
        from collections import Counter
        c = Counter(d for _x, _y, d in texts)
        print("  값 분포:", " · ".join(f"{k}A×{v}" for k, v in sorted(c.items())))

    net, edge_ref = got["kfp"], got["edge_ref"]
    loads = (got.get("worst") or {}).get("loads") or {}
    pts = es.board.pts
    bores = decide_bores(net, edge_ref, loads, texts, pts=pts)
    cnt = source_counts(bores)
    total = sum(cnt.values())
    print(f"\n관경 근거 — 배관 {total}개")
    for k, label in (("text", "도면 텍스트"), ("nfpc_min", "별표1 보강(안전측)"),
                     ("nfpc_fallback", "별표1 폴백")):
        pct = (cnt[k] / total * 100.0) if total else 0.0
        print(f"  {label:18s} {cnt[k]:4d}  ({pct:5.1f}%)")

    # ── 폴백의 원인을 가른다 ────────────────────────────────────────
    no_ref = with_ref_no_text = 0
    dists: list[float] = []
    for pid in (net.get("pipe_data") or {}):
        if bores.get(pid, (0, ""))[1] != "nfpc_fallback":
            continue
        ref = edge_ref.get(pid)
        if ref is None:
            no_ref += 1
            continue
        i, j = ref
        if not (0 <= i < len(pts) and 0 <= j < len(pts)):
            no_ref += 1
            continue
        with_ref_no_text += 1
        # 한계를 풀고 가장 가까운 텍스트까지의 거리를 잰다.
        near = match_diameter_for_segment(pts[i], pts[j], texts,
                                          limit_mm=float("inf"))
        if near is not None:
            best = min(
                _pt_seg(tx, ty, pts[i], pts[j]) for tx, ty, _d in texts) \
                if texts else float("inf")
            dists.append(best)

    print(f"\n별표1 폴백 {cnt['nfpc_fallback']}개의 원인")
    print(f"  ㉮ 역참조 없음(헤드 접속관·가지 상승 등) {no_ref:4d}")
    print(f"  ㉯ 역참조 있으나 텍스트 없음            {with_ref_no_text:4d}")

    if dists:
        dists.sort()
        print(f"\n  ㉯ 의 «가장 가까운 치수 텍스트까지 거리» 분포 (n={len(dists)})")
        for q, name in ((0, "최소"), (len(dists) // 4, "25%"),
                        (len(dists) // 2, "중앙"), (3 * len(dists) // 4, "75%"),
                        (len(dists) - 1, "최대")):
            print(f"     {name:>4s} {dists[q]:10,.0f} mm")
        print("\n  한계를 늘리면 건지는 수 (누적)")
        for lim in (1500, 2000, 3000, 5000, 8000, 12000, 20000):
            n = sum(1 for d in dists if d <= lim)
            print(f"     ≤{lim:>6,}mm : {n:4d} / {len(dists)}")

    # ── 안전측 규칙이 실제로 몇 번 이겼나 ───────────────────────────
    overridden = []
    for pid, (dia, srcname) in bores.items():
        if srcname != "nfpc_min":
            continue
        ref = edge_ref.get(pid)
        if ref is None:
            continue
        i, j = ref
        t = match_diameter_for_segment(pts[i], pts[j], texts)
        if t is not None:
            n_head = int(loads.get((min(i, j), max(i, j)), 0))
            overridden.append((t, nfpc_min_bore_mm(n_head), n_head))
    print(f"\n안전측(별표1이 도면 표기를 이김) {len(overridden)}건")
    for t, m, n in overridden[:10]:
        print(f"  도면 {t}A → 별표1 {m}A (담당 헤드 {n})")
    return 0


def _pt_seg(px, py, a, b) -> float:
    import math
    ax, ay, bx, by = a[0], a[1], b[0], b[1]
    dx, dy = bx - ax, by - ay
    L2 = dx * dx + dy * dy
    if L2 < 1e-9:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / L2))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


if __name__ == "__main__":
    sys.exit(main())
