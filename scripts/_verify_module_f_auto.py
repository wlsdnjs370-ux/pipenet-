# -*- coding: utf-8 -*-
"""[A 방식] 평면도 자동 추출을 실도면으로 끝까지 태운다.

단위 테스트는 어댑터와 라벨 규약만 본다. 여기서는 실제 DXF 로

    파싱 → 헤드 후보 → 영역·알람밸브 → anchored 선정 → 5종 입력표

를 돌리고, 그 표가 **수동 경로와 같은 자리에** 들어가는지(통합 S740 이 받는지)
까지 본다. 라벨 규약이 어긋나면 여기서 드러난다.

    python scripts/_verify_module_f_auto.py [도면.dxf]
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FAILS: list[str] = []
CANDIDATES = [
    ROOT / "samples" / "dxf" / "LH306동_평면도.dxf",
    ROOT / "data" / "sample_problem" / "대명동201동 단위세대_layer정리.dxf",
    ROOT / "samples" / "dxf" / "LH306동.dxf",
]


def check(label: str, cond: bool, detail: str = "") -> bool:
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{'OK  ' if cond else 'FAIL'}] {label}"
          + (f" · {detail}" if detail else ""))
    return cond


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    from routes.module_f.auto import (detect_head_candidates, parse_plan,
                                      preview_view, run_auto)
    from routes.module_f.merge import label_offset_for, to_head_tables

    dxf = (Path(sys.argv[1]) if len(sys.argv) > 1
           else next((p for p in CANDIDATES if p.is_file()), None))
    if dxf is None or not dxf.is_file():
        check("실도면", False, "후보를 못 찾음")
        return 1
    print(f"[A 방식] 자동 추출 실측 — {dxf.name}")

    # ① 파싱
    ents, layer_cat, diag = parse_plan(dxf)
    check("도면 파싱", diag["entities"] > 0,
          f"도형 {diag['entities']:,} · 레이어 {diag['layers']}")
    print("       용도: " + " · ".join(f"{k} {v}"
                                       for k, v in sorted(diag["cats"].items())))
    if (diag.get("xref") or {}).get("is_xref_shell"):
        print("       ★외부참조 껍데기 — 헤드가 안 나올 수 있다")

    # ② 헤드 후보
    heads = detect_head_candidates(ents, layer_cat)
    check("헤드 후보 검출", len(heads) > 0, f"{len(heads):,}개")
    if not heads:
        return 1
    xs = [h["x"] for h in heads]
    ys = [h["y"] for h in heads]
    print(f"       범위 x {min(xs):.0f}~{max(xs):.0f} · y {min(ys):.0f}~{max(ys):.0f}")

    # ③ 사람이 찍는 두 값을 재현 가능하게 고정한다.
    #    영역   = 헤드 전체를 감싸는 사각형 하나
    #    알람밸브 = 헤드 무리 **한가운데에 가장 가까운 헤드** 의 좌표
    #
    #    ★bbox 모서리를 쓰면 안 된다 — 거기엔 배관이 없다. A 는 알람밸브가
    #      배관망 컴포넌트에서 25m 안에 있어야 결합한다(실측으로 여기서 한 번
    #      막혔다). 헤드는 배관에 붙어 있으므로 헤드 좌표가 안전하다.
    pad = 1000.0
    rects = [[min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad]]
    cx = sum(xs) / len(xs)
    cy = sum(ys) / len(ys)
    mid = min(heads, key=lambda h: (h["x"] - cx) ** 2 + (h["y"] - cy) ** 2)
    alarm = (mid["x"], mid["y"])
    print(f"       알람밸브(시험용) = 한가운데 헤드 ({alarm[0]:.0f}, {alarm[1]:.0f})")

    # ④ 선정 + 표
    try:
        got = run_auto(ents, layer_cat, alarm_xy=alarm, rects=rects, k=30)
    except Exception as exc:  # noqa: BLE001
        check("anchored 선정", False, f"{type(exc).__name__}: {exc}")
        return 1
    s = got["summary"]
    check("anchored 선정", s["k"] > 0,
          f"헤드 {s['k']} · 절점 {s['nodes']} · 배관 {s['pipes']}"
          f" · 최원 {s['far_m']} m")
    if s["source_fallback"]:
        print(f"       ★급수원이 그래프에서 {s['source_bridge_mm']:.0f}mm 떨어져 "
              f"최근접으로 대체됨")

    tbl = got["tables"]
    check("5종 입력표", bool(tbl.nodes and tbl.pipes and tbl.nozzles),
          f"절점 {len(tbl.nodes)} · 배관 {len(tbl.pipes)} · "
          f"노즐 {len(tbl.nozzles)} · 부속 {len(tbl.fittings)}")

    # ⑤ A 의 표는 기준점이 이미 10 이다
    labels = [str(n.get("label")) for n in tbl.nodes]
    check("기준점이 10 이다", "10" in labels, f"첫 라벨 {labels[:5]}")
    root = next((n for n in tbl.nodes if str(n.get("io_node")) == "Input"), None)
    check("급수원이 10 이다",
          root is not None and str(root.get("label")) == "10",
          str(root and root.get("label")))

    # ⑥ 통합(S740)이 그대로 받는가 — 오프셋을 자동으로 골라야 한다
    ht = to_head_tables(tbl, offset=label_offset_for("auto"))
    check("통합이 자동 표를 그대로 받는다",
          any(n["label"] == "10" for n in ht.nodes),
          f"절점 {len(ht.nodes)}")
    from routes.module_f.merge import MergeError
    try:
        to_head_tables(tbl, offset=9)
        check("잘못된 오프셋은 막힌다", False, "통과해 버렸다")
    except MergeError:
        check("잘못된 오프셋은 막힌다", True, "+9 를 먹이면 기준점이 사라진다")

    # ⑦ 미리보기
    v = preview_view(tbl)
    n_head = sum(1 for n in v["nodes"] if n.get("head"))
    check("미리보기에 헤드가 표시된다", n_head == s["k"],
          f"{n_head} / {s['k']}")

    print()
    if FAILS:
        print(f"FAIL {len(FAILS)}건")
        for f in FAILS:
            print("  - " + f)
        return 1
    print("PASS — 자동 추출 전 항목 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
