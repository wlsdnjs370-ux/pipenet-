# -*- coding: utf-8 -*-
"""[G8] 모듈 A 대조 — 같은 도면을 A 로도 돌려 위상이 같은지 본다.

지시서 §G8 의 다섯 항목:
    앵커 헤드 좌표(±SNAP) · 선정 헤드 집합(K개) · 최원 유하거리 ·
    max_load · corridor 총연장

★두 계통은 **입력이 다르다**. A 는 DXF 엔티티에서 제 그래프를 만들고(자동 추출),
G 는 사람이 손질한 EditBoard 위에서 돈다. 그래서 숫자가 똑같기를 기대하는 것이
아니라, **다르면 그 차이가 설명되는지** 를 본다. 설명되지 않는 차이가 버그다.

    python tests/compare_module_a.py
"""
from __future__ import annotations

import json
import math
import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
_REPO = _ROOT.parent
for p in (str(_ROOT), str(_REPO)):
    if p not in sys.path:
        sys.path.insert(0, p)
os.chdir(_ROOT)

KEY = "B1F 현장조사 소화설비 평면도"
OUT = _ROOT / "tests" / "_out"
DOC = _REPO / "docs" / "module_g_vs_a.md"
SNAP_MM = 500.0        # 앵커 좌표 일치로 볼 여유


def g_side():
    """모듈 G — 손질 저장본 위에서 앵커 방식."""
    from services.cad_import.edit.session import EditSession
    from services.cad_import.design.restrict import select_and_expand

    es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
    payload = es.convert_payload()
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None
    got = select_and_expand(payload, es.board, k=30, selected_source=sel)
    if not got.get("ok"):
        return {"ok": False, "error": got.get("error")}
    w = got["worst"]
    disks = es.board.disks
    an = w.get("anchor")
    return {
        "ok": True,
        "anchor_xy": (float(disks[an][0]), float(disks[an][1]))
        if an is not None and an < len(disks) else None,
        "heads_xy": [(float(disks[h][0]), float(disks[h][1]))
                     for h in w["heads"] if h < len(disks)],
        "far_m": w["far_m"], "max_load": w["max_load"],
        "total_m": w["total_m"], "span_m": w["span_m"],
        "excluded": got.get("excluded_heads", 0),
        "candidates": got.get("candidate_heads", 0),
        "total_heads": got.get("total_heads", 0),
    }


def a_side():
    """모듈 A — 제 DXF 파이프라인으로 자동 추출 후 최불리 30."""
    try:
        import remote30_prototype as A
    except Exception as exc:      # noqa: BLE001
        return {"ok": False, "error": f"모듈 A import 실패: {exc}"}

    spec_path = _ROOT / "docs" / "import" / "0단계_새찍기" / f"{KEY}_찍은스펙.json"
    src = json.loads(spec_path.read_text(encoding="utf-8")).get("source_dxf")
    if not src or not os.path.isfile(src):
        return {"ok": False, "error": f"원본 DXF 를 찾을 수 없습니다: {src}"}

    try:
        # A 의 실제 진입점 — 파일 내용 해시로 캐시하는 번들 파서.
        bundle = A.parse_dxf_bundle_cached(Path(src))
        ents = bundle.entities
        # 레이어 카테고리는 A 의 이름 사전이 정한다(라우트도 이 값을 넘긴다).
        layers = {ly.get("name"): A._categorize_layer(ly.get("name") or "")
                  for ly in (bundle.layers or [])}
    except Exception as exc:      # noqa: BLE001
        return {"ok": False, "error": f"A 파싱 실패: {type(exc).__name__} {exc}"}

    try:
        sel = A.select_worst30_heads(ents, layers, k=30)
    except Exception as exc:      # noqa: BLE001
        return {"ok": False, "error": f"A 선정 실패: {exc}"}

    heads = [h.pos for h in getattr(sel, "heads", [])]
    total = sum(L for *_e, L in getattr(sel, "edges", [])) / 1000.0
    span = 0.0
    if heads:
        xs = [p_[0] for p_ in heads]
        ys = [p_[1] for p_ in heads]
        span = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) / 1000.0
    return {"ok": True, "heads_xy": heads, "total_m": round(total, 2),
            "spread_m": round(span, 1), "edges": len(getattr(sel, "edges", [])),
            # A 의 비-anchored 경로는 «먼 순서» 라 앵커라는 개념이 없다.
            # 첫 헤드를 앵커로 부르면 거짓이 되므로 그렇게 쓰지 않는다.
            "anchor_xy": None}


def main() -> int:
    print("[G8] 모듈 A 대조\n")
    g = g_side()
    if not g.get("ok"):
        print("  !! G 쪽 실패:", g.get("error"))
        return 1
    print(f"  G : 앵커 {g['anchor_xy']} · 최원 {g['far_m']} m · "
          f"max_load {g['max_load']} · corridor {g['total_m']} m")
    print(f"      후보 {g['candidates']:,} / 도면 {g['total_heads']:,} "
          f"(전개가 못 붙여 제외 {g['excluded']:,})")

    a = a_side()
    rows = []
    if not a.get("ok"):
        print("  A : 돌리지 못함 —", a.get("error"))
        rows.append(("대조 상태", "G 단독", "A 미실행", a.get("error", "")))
    else:
        print(f"  A : 앵커 {a['anchor_xy']} · corridor {a['total_m']} m")
        # G 헤드의 퍼짐 — 설계면적인지 아닌지를 가르는 지표(D1).
        gxs = [p_[0] for p_ in g["heads_xy"]]
        gys = [p_[1] for p_ in g["heads_xy"]]
        g_spread = round(math.hypot(max(gxs) - min(gxs),
                                    max(gys) - min(gys)) / 1000.0, 1)
        rows.append(("선정 방식", "앵커 (D1)", "먼 순서 (A 기존)",
                     "지시서 D1 이 바꾸라고 한 지점"))
        rows.append(("선정 헤드 수", str(len(g["heads_xy"])),
                     str(len(a["heads_xy"])), ""))
        rows.append(("**헤드 퍼짐 (대각, m)**", str(g_spread),
                     str(a["spread_m"]),
                     "작을수록 한 설계구역 — G 가 뭉친다"))
        rows.append(("앵커 좌표", str(g["anchor_xy"]), "없음",
                     "A 비-anchored 경로엔 앵커 개념이 없다"))
        rows.append(("최원 유하거리 (m)", str(g["far_m"]), "—", ""))
        rows.append(("max_load", str(g["max_load"]), "—", ""))
        rows.append(("corridor 총연장 (m)", str(g["total_m"]),
                     str(a["total_m"]),
                     "A 는 흩어진 30개라 경로가 길다"))
        rows.append(("corridor 간선 수", "—", str(a["edges"]), ""))

    OUT.mkdir(parents=True, exist_ok=True)
    DOC.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# 모듈 G vs 모듈 A — 최불리 배관망 대조", "",
        f"도면: `{KEY}`", "",
        "## 왜 숫자가 같기를 기대하지 않는가", "",
        "두 계통은 **입력이 다르다**. A 는 DXF 엔티티에서 제 그래프를 자동 추출하고,",
        "G 는 사람이 손질한 `EditBoard` 위에서 돈다(모듈 F 와 같은 망). 그래서 대조의",
        "목적은 «같은 값» 이 아니라 «다르면 그 차이가 설명되는가» 다.", "",
        "## G 결과", "",
        f"- 앵커 좌표 : `{g['anchor_xy']}`",
        f"- 최원 유하거리 : {g['far_m']} m",
        f"- 설계면적 폭 : {g['span_m']} m",
        f"- corridor 총연장 : {g['total_m']} m",
        f"- 주배관 담당 헤드 수(max_load) : {g['max_load']}",
        f"- 후보 헤드 : {g['candidates']:,} / 도면 {g['total_heads']:,}",
        f"- **전개가 붙이지 못해 제외한 헤드 : {g['excluded']:,}**", "",
        "## 대조", "",
        "| 항목 | G | A | 비고 |",
        "|---|---|---|---|",
    ]
    for r in rows:
        lines.append(f"| {r[0]} | {r[1]} | {r[2]} | {r[3]} |")
    lines += [
        "", "## 설명되어야 하는 차이", "",
        "- **후보 집합** — G 는 전개가 배관에 붙일 수 있는 헤드에서만 고른다",
        "  (BLOCKED B4 1안). B1F 에서 619 / 3,163 이라 A 와 후보 모수가 다르다.",
        "- **망의 출처** — A 는 자동 추출망, G 는 손질망이다. 손질로 이어 붙인",
        "  배관이 있으면 G 쪽 도달 범위가 넓다.",
        "- **corridor 정의** — G 의 `total_m` 은 최단경로 합집합(board 간선)이고,",
        "  전개 배관 길이 합과는 구조적으로 다르다(BLOCKED B5).",
        "- **선정 방식 자체** — A 의 비-anchored 경로는 「급수원에서 먼 순서 K개」다.",
        "  지시서 D1 이 바꾸라고 한 지점이라 **결과가 다른 것이 정상**이다.",
        "  실측 퍼짐이 그 차이를 그대로 보여준다(A 134.1 m vs G 31.4 m).",
        "  A 에도 앵커 경로(`select_worst30_heads_anchored`)가 있으나 알람밸브",
        "  좌표와 헤드 영역이 둘 다 있어야 서고, 이 저장본에는 없어 세우지 못했다.", "",
        "관경이 다른 간선은 전부 `source`(text / nfpc_min / nfpc_fallback)로",
        "설명되어야 한다. 설명되지 않는 차이가 버그다.", "",
    ]
    DOC.write_text("\n".join(lines), encoding="utf-8")
    print(f"\n  대조표: {DOC}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
