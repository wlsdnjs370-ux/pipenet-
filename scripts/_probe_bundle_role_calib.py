# -*- coding: utf-8 -*-
"""찍기 묶음의 «기하로 본 역할» — 문턱을 **재서** 정한다.

찍기 화면의 재료 묶음은 (레이어 × 색) 단위이고, 그 안에 든 것은 **선분**이다.
앞선 진단(`_probe_riser_layer_role.py`)은 entity 단위(닫힘 여부·도형 지름)로
갈랐는데 그 정보는 묶음에 없다. 그래서 선분 단위 지표만으로도 갈리는지를 먼저
재고, 갈린다면 그 수치로 문턱을 정한다. **문턱을 먼저 정하고 맞는 도면을
찾으면 그것은 캘리브레이션이 아니라 끼워 맞추기다.**

대조군을 같이 잰다:
    계통도·기계실   이름 사전이 틀린 도면 (SP=헤드, FF·0=배관)
    대명동·LH306    이름 사전이 맞는 도면 — 여기서도 배관이 «배관다움» 이어야
                    한다. 안 그러면 지표가 도면 종류를 가르는 것일 뿐이다.

    python scripts/_probe_bundle_role_calib.py [도면.dxf ...]
"""
from __future__ import annotations

import math
import statistics
import sys
from collections import defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = [
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 계통도.dxf",
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 기계실.dxf",
    ROOT / "samples" / "dxf" / "대명동201동 단위세대_layer정리.dxf",
    ROOT / "samples" / "dxf" / "LH306동_평면도.dxf",
]
LONG_MM = 500.0


def seg_lengths(en) -> list:
    """entity → 선분 길이 목록. 찍기 world 가 담는 것과 같은 단위다."""
    t = str(en.get("t") or "")
    if t == "L":
        x1, y1, x2, y2 = en["p"]
        return [math.hypot(x2 - x1, y2 - y1)]
    if t == "PL":
        pts = en["p"]
        return [math.hypot(b[0] - a[0], b[1] - a[1])
                for a, b in zip(pts, pts[1:])]
    return []


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
        seg = defaultdict(list)
        rnd = defaultdict(int)          # 원·호 개수 — 기호의 신호
        for en in bundle.entities:
            # 색이 목록으로 오는 entity 가 있다 — 키로 쓰려면 문자열로 굳힌다.
            key = (str(en.get("l") or "0"), str(en.get("c")))
            ls = seg_lengths(en)
            if ls:
                seg[key] += ls
            elif str(en.get("t")) in ("C", "A"):
                rnd[key] += 1

        print(f"\n{'=' * 96}")
        print(f"■ {dxf.name}")
        print("=" * 96)
        print(f"   {'레이어':<16}{'색':>5}{'분류':>7}{'선분':>8}{'긴선분%':>9}"
              f"{'중앙값':>9}{'상위10%':>10}{'원호':>7}")
        rows = sorted(seg.items(), key=lambda kv: -len(kv[1]))[:12]
        for (ly, c), ls in rows:
            ls = [v for v in ls if v > 0]
            if len(ls) < 20:
                continue
            longp = 100.0 * sum(1 for v in ls if v >= LONG_MM) / len(ls)
            med = statistics.median(ls)
            p90 = sorted(ls)[int(len(ls) * 0.9)]
            print(f"   {ly[:15]:<16}{str(c):>5}{layers.get(ly, 'OTHER'):>7}"
                  f"{len(ls):>8,}{longp:>8.0f}%{med:>9,.0f}{p90:>10,.0f}"
                  f"{rnd.get((ly, c), 0):>7,}")
    print("\n  읽는 법 — 배관은 «긴 선분이 많고 중앙값이 크다», 기호는 그 반대다.")
    print("  이름 사전이 맞는 도면(대명동·LH306)의 PIPE 줄이 기준선이 된다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
