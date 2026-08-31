# -*- coding: utf-8 -*-
"""레이어의 «역할» 을 기하로 판정한다 — 이름이 거짓말할 때.

대명동 계통도에서 그림으로 본 것을 수치로 못 박는다:

    SP  [PIPE 로 분류됨]   실제로는 **헤드 기호** (삼각형 918개)
    FF  [OTHER 로 분류됨]  실제로는 **배관** (가로 분기선 7,333)

이름 사전이 «SP = 스프링클러 배관» 으로 읽었는데, 이 도면에서 SP 는 스프링클러
**헤드** 레이어다. 그래서 그래프는 헤드 삼각형으로 세워지고(조각 63개) 진짜
배관은 통째로 빠진다. 증상 두 개가 한 원인에서 나온다.

기하로 가르는 법 — 배관은 «길고 곧은 선», 기호는 «작고 닫힌 도형»:

    긴 선분 비율      길이 ≥ 500mm 인 선분이 얼마나 되나
    도형 지름         entity 하나가 차지하는 크기의 중앙값
    닫힘              첫점=끝점(닫힌 폴리곤) 비율 — 기호는 닫혀 있다

    python scripts/_probe_riser_layer_role.py [도면.dxf ...]
"""
from __future__ import annotations

import math
import statistics
import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = [
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 계통도.dxf",
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 기계실.dxf",
]
LONG_MM = 500.0          # 이보다 길면 «배관다운» 선분


def ent_geom(en):
    """entity → (선분 길이 목록, 경계상자 지름, 닫혔나)."""
    t = str(en.get("t") or "")
    pts = []
    closed = False
    if t == "L":
        x1, y1, x2, y2 = en["p"]
        pts = [(x1, y1), (x2, y2)]
    elif t == "PL":
        pts = [(p[0], p[1]) for p in en["p"]]
        if len(pts) >= 3 and math.hypot(pts[0][0] - pts[-1][0],
                                        pts[0][1] - pts[-1][1]) <= 1.0:
            closed = True
    elif t in ("C", "A"):
        cx, cy = en["c"]
        r = float(en.get("r", 0.0) or 0.0)
        return [], 2 * r, (t == "C")
    else:
        return [], 0.0, False
    lens = [math.hypot(b[0] - a[0], b[1] - a[1]) for a, b in zip(pts, pts[1:])]
    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    diag = math.hypot(max(xs) - min(xs), max(ys) - min(ys)) if pts else 0.0
    return lens, diag, closed


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    for dxf in [Path(x) for x in sys.argv[1:]] or DEF:
        if not dxf.is_file():
            print(f"\n■ {dxf.name} — 파일 없음")
            continue
        bundle = A.parse_dxf_bundle_cached(dxf)
        layers = {ly.get("name"): (ly.get("auto_category") or "OTHER")
                  for ly in (bundle.layers or [])}
        by = defaultdict(list)
        for en in bundle.entities:
            by[str(en.get("l") or "0")].append(en)

        print(f"\n{'=' * 92}")
        print(f"■ {dxf.name}")
        print("=" * 92)
        print(f"   {'레이어':<16}{'분류':<7}{'ent':>7}{'종류':>16}"
              f"{'긴선분%':>9}{'지름중앙':>10}{'닫힘%':>8}   판정")
        rows = []
        for nm, ee in sorted(by.items(), key=lambda kv: -len(kv[1])):
            if len(ee) < 30:
                continue
            allen, diags, nclosed, types = [], [], 0, Counter()
            for en in ee:
                lens, diag, cl = ent_geom(en)
                allen += lens
                if diag > 0:
                    diags.append(diag)
                nclosed += 1 if cl else 0
                types[str(en.get("t") or "?")] += 1
            if not diags:
                continue
            longp = (100.0 * sum(1 for v in allen if v >= LONG_MM)
                     / max(1, len(allen)))
            med = statistics.median(diags)
            clp = 100.0 * nclosed / len(ee)
            # 판정 — 배관은 «길고 열린 선», 기호는 «작고 닫힌 도형».
            if longp >= 25 and med >= LONG_MM:
                verdict = "배관다움"
            elif med <= 400 and (clp >= 30 or types.get("H") or types.get("C")):
                verdict = "기호다움"
            else:
                verdict = "—"
            mark = ""
            cat = layers.get(nm, "OTHER")
            if verdict == "배관다움" and cat != "PIPE":
                mark = "  ★배관인데 PIPE 가 아니다"
            if verdict == "기호다움" and cat == "PIPE":
                mark = "  ★기호인데 PIPE 로 잡혔다"
            rows.append((nm, cat, len(ee), types.most_common(2),
                         longp, med, clp, verdict, mark))
        for nm, cat, k, tt, longp, med, clp, v, mark in rows[:14]:
            ts = " ".join(f"{a}{b}" for a, b in tt)
            print(f"   {nm:<16}{cat:<7}{k:>7,}{ts:>16}"
                  f"{longp:>8.0f}%{med:>10,.0f}{clp:>7.0f}%   {v}{mark}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
