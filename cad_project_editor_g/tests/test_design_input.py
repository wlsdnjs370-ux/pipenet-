# -*- coding: utf-8 -*-
"""모듈 G 수리계산 입력 — 항목별 수용 기준 검증(지시서 §4).

    python tests/test_design_input.py [G1 G2 ...]      항목 지정(기본 전부)

Qt 없이 돈다(헤드리스). 실도면 통합은 tests/_out/ 아래에만 쓴다 — G 트리 밖으로
중간 산출물을 내보내지 않는다(§3).
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))
os.chdir(_ROOT)

KEY = "B1F 현장조사 소화설비 평면도"
OUT_DIR = _ROOT / "tests" / "_out"
FAILS: list[str] = []


def check(label, cond, detail=""):
    mark = "OK  " if cond else "FAIL"
    if not cond:
        FAILS.append(f"{label} — {detail}")
    print(f"  [{mark}] {label}" + (f" · {detail}" if detail else ""))
    return cond


def _board():
    """손질 저장본을 연다. 여러 테스트가 공유하므로 한 번만 만든다."""
    if not hasattr(_board, "_b"):
        from services.cad_import.edit.session import EditSession
        es = EditSession.open(KEY, out_dir=None, load_saved=True, use_cache=True)
        _board._b = es
    return _board._b


# ─────────────────────────────────────────────────────────── G1
def g1():
    print("\n[G1] 최불리 K 선정 이식")
    from services.cad_import.design.worst import worst_k_heads

    es = _board()
    b = es.board
    check("급수 시작 위치 있음", bool(b.sources), f"{len(b.sources)}곳")

    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30)
    need = ("heads", "anchor", "edges", "loads", "nodes",
            "far_m", "near_m", "span_m", "total_m", "max_load")
    missing = [k for k in need if k not in w]
    check("반환 키 10종", not missing, f"빠짐 {missing}" if missing else "전부 있음")
    check("K개 선정", len(w["heads"]) == 30, f"{len(w['heads'])}개")

    # ★핵심: 모듈 F 와 완전히 일치해야 이식이 성공한 것이다.
    sys.path.append(str(_ROOT.parent))
    from routes.module_f.remote30 import _worst_k_heads as f_worst
    wf = f_worst(b.pts, b.edges, b.hnodes, b.sources, k=30)
    check("모듈 F 와 앵커 일치", w["anchor"] == wf["anchor"],
          f"G {w['anchor']} / F {wf['anchor']}")
    check("모듈 F 와 헤드 집합 일치", set(w["heads"]) == set(wf["heads"]),
          f"차집합 {len(set(w['heads']) ^ set(wf['heads']))}개")
    check("모듈 F 와 far_m 일치", w["far_m"] == wf["far_m"],
          f"G {w['far_m']} / F {wf['far_m']}")
    check("모듈 F 와 max_load 일치", w["max_load"] == wf["max_load"],
          f"G {w['max_load']} / F {wf['max_load']}")

    # 앵커 방식이면 설계면적이 뭉친다 — 「먼 순서」와 갈리는 지점.
    check("설계면적 폭이 corridor 총연장보다 작다",
          w["span_m"] < w["total_m"],
          f"폭 {w['span_m']} m / 총연장 {w['total_m']} m")
    print(f"      앵커 {w['far_m']} m · 폭 {w['span_m']} m · "
          f"연장 {w['total_m']} m · max_load {w['max_load']}")
    return w


# ─────────────────────────────────────────────────────────── G2
def g2():
    print("\n[G2] corridor 제한 전개 + 역참조")
    from services.cad_import.design.restrict import expand_worst
    from services.cad_import.design.worst import worst_k_heads

    es = _board()
    b = es.board
    w = worst_k_heads(b.pts, b.edges, b.hnodes, b.sources, k=30)
    payload = es.convert_payload()
    # 이 저장본은 급수원이 둘이다 — 기준선과 같은 것(Z1)을 쓴다(BLOCKED B2).
    srcs = payload.get("sources") or ()
    sel = srcs[0].get("tag") if len(srcs) > 1 else None

    got = expand_worst(payload, b, w, selected_source=sel)
    if not check("제한 전개 성공", got.get("ok"), str(got.get("error"))[:90]):
        return None

    kfp = got["kfp"]
    pipes = kfp.get("pipe_data") or {}
    nodes = kfp.get("nodes_meta_runtime") or {}
    # ★수용 기준은 «입력에 30개를 넣었나» 가 아니라 «전개가 30개를 살렸나» 다.
    #   hcov 는 입력을 되돌려줄 뿐이라 그걸 세면 빈 망도 초록불이 된다(실제로 그랬다).
    n_built_heads = sum(1 for m in (kfp.get("nodes_meta_runtime") or {}).values()
                        if str((m or {}).get("type_id", "")) == "head")
    check("제한 전개가 살린 헤드 수 == 선정 K",
          n_built_heads == len(w["heads"]),
          f"전개 {n_built_heads} / 선정 {len(w['heads'])}  "
          f"(BLOCKED B4 — 선정과 전개가 «물 닿음» 을 다르게 본다)")
    check("역참조가 모든 배관을 덮는다", not got["uncovered_pipes"],
          f"미포함 {len(got['uncovered_pipes'])}개")
    check("역참조가 board 간선을 가리킨다",
          all(isinstance(v, tuple) and len(v) == 2
              and 0 <= v[0] < len(b.pts) and 0 <= v[1] < len(b.pts)
              for v in got["edge_ref"].values()),
          f"{len(got['edge_ref'])}건")
    check("node_ref 존재", bool(got["node_ref"]), f"{len(got['node_ref'])}건")
    # 제한 전개는 전체망보다 작아야 한다 — 그게 «제한» 의 뜻이다.
    check("제한망이 corridor 규모다", len(pipes) >= len(w["heads"]),
          f"배관 {len(pipes)} · 선정 헤드 {len(w['heads'])}개를 먹이려면 그 이상")
    print(f"      제한망 노드 {len(nodes)} · 배관 {len(pipes)}")
    return got


ITEMS = {"G1": g1, "G2": g2}


def main() -> int:
    want = sys.argv[1:] or list(ITEMS)
    for name in want:
        fn = ITEMS.get(name)
        if fn is None:
            print(f"  (모르는 항목: {name})")
            continue
        fn()
    print("\n" + "=" * 56)
    if FAILS:
        for f in FAILS:
            print("  !!", f)
        print(f"\n실패 {len(FAILS)}건")
        return 1
    print("수용 기준 통과")
    return 0


if __name__ == "__main__":
    sys.exit(main())
