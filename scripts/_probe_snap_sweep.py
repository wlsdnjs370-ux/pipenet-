# -*- coding: utf-8 -*-
"""이음매 간격(auto_snap_eps) 스윕을 열어 본다 — 왜 조각이 306개인가.

`select_worst30_heads_anchored` 는 이미 이음매 간격을 도면에서 «재서» 쓴다
(R1). 그런데 B1F 는 조각이 306개이고 최대 조각이 전체의 8% 다. 그러면 둘 중
하나다:

  ① 스윕이 고른 값이 이 도면에 안 맞다
  ② 후보 상한이 낮아 애초에 닫을 수 있는 값이 목록에 없다

감사(audit.snap_eps)에 시도 기록이 남으므로 그것을 그대로 펼친다. 그리고
후보를 «더 넓게» 줬을 때 최대 조각과 총연장이 어떻게 움직이는지 같이 잰다 —
R1 이 쓴 그 기준(최대 조각 ↑, 총연장 90% 하한)으로.

    python scripts/_probe_snap_sweep.py [도면.dxf] [--wide]
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--wide", action="store_true",
                    help="후보를 1200mm 까지 넓혀 다시 스윕")
    a = ap.parse_args()
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    print(f"{dxf.name}")
    print(f"  현재 후보 {list(A.SNAP_EPS_CANDIDATES_MM)}")
    print(f"  총연장 하한 비율 {A.SNAP_EPS_MIN_LEN_RATIO} · "
          f"최소간선 가드 {A.SNAP_EPS_GUARD_MIN_EDGE_MM}\n")

    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    cat = {}
    for nm in {str(e.get("l") or "0") for e in ents}:
        try:
            cat[nm] = A._categorize_layer(nm)
        except Exception:  # noqa: BLE001
            cat[nm] = "OTHER"

    audit: dict = {}
    eps = A.auto_snap_eps(ents, cat, audit_out=audit)
    rec = audit.get("snap_eps") or {}
    print(f"■ 스윕이 고른 값 — {rec.get('chosen_mm', eps)} mm\n")
    tr = rec.get("trials") or []
    if tr:
        keys = [k for k in ("eps_mm", "largest_component", "components",
                            "total_len_mm", "edges", "nodes")
                if k in tr[0]]
        print("  " + " ".join(f"{k:>18}" for k in keys))
        print("  " + "-" * (19 * len(keys)))
        for t in tr:
            print("  " + " ".join(f"{str(t.get(k)):>18}" for k in keys))
    else:
        print("  (시도 기록 없음)")

    if not a.wide:
        print("\n  --wide 로 후보를 넓혀 다시 재려면:")
        print("    python scripts/_probe_snap_sweep.py --wide")
        return 0

    # ── 후보를 넓혀 다시 — R1 의 기준 그대로(최대 조각 ↑ · 총연장 하한).
    wide = (30., 50., 75., 90., 130., 160., 200., 300., 450., 600.,
            800., 1000., 1200.)
    print(f"\n■ 후보를 넓혀 다시 스윕 {list(wide)}\n")
    old = A.SNAP_EPS_CANDIDATES_MM
    try:
        A.SNAP_EPS_CANDIDATES_MM = wide
        a2: dict = {}
        eps2 = A.auto_snap_eps(ents, cat, audit_out=a2)
        r2 = a2.get("snap_eps") or {}
        tr2 = r2.get("trials") or []
        if tr2:
            keys = [k for k in ("eps_mm", "largest_component", "components",
                                "total_len_mm", "edges", "nodes")
                    if k in tr2[0]]
            print("  " + " ".join(f"{k:>18}" for k in keys))
            print("  " + "-" * (19 * len(keys)))
            base = None
            for t in tr2:
                if base is None:
                    base = float(t.get("total_len_mm") or 0) or 1.0
                keep = float(t.get("total_len_mm") or 0) / base
                mark = ""
                if keep < A.SNAP_EPS_MIN_LEN_RATIO:
                    mark = "  ★총연장 하한 미달 — 채택 불가"
                print("  " + " ".join(f"{str(t.get(k)):>18}" for k in keys)
                      + mark)
        print(f"\n  넓힌 뒤 고른 값 — {r2.get('chosen_mm', eps2)} mm")
    finally:
        A.SNAP_EPS_CANDIDATES_MM = old
    print("\n  최대 조각이 크게 오르는 «무릎» 이 있으면 그 값이 이 도면의 이음매다.")
    print("  안 오르면 틈이 끝점끼리가 아니라는 뜻 — 클러스터로는 못 닫는다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
