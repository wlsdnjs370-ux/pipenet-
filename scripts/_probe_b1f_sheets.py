# -*- coding: utf-8 -*-
"""B1F 는 한 파일에 도면이 몇 장인가 — 자동 경로의 «기본 범위» 가 옳은지 본다.

자동 추출은 영역을 안 그리면 «검출한 헤드 전부» 를 범위로 삼는다. 도면이 한
장이면 그것이 맞다. 그런데 한 파일에 여러 장이 들어 있으면 그 범위가 장을
가로질러, 최불리 30이 서로 다른 도면의 헤드를 섞어 뽑는다 — A 가 실측으로
겪은 바로 그 문제다(그래서 A 에 detect_sheet_frames 가 있다).

    python scripts/_probe_b1f_sheets.py [도면.dxf]
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    dxf = Path(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    import remote30_prototype as A

    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    names = {str(e.get("l") or "0") for e in ents}
    cat = {}
    for n in names:
        try:
            cat[n] = A._categorize_layer(n)
        except Exception:  # noqa: BLE001
            cat[n] = "OTHER"

    heads = A.detect_heads(ents, cat)
    xs = [h.pos[0] for h in heads]
    ys = [h.pos[1] for h in heads]
    print(f"{dxf.name}")
    print(f"  헤드 {len(heads):,}개")
    print(f"  전체 범위 {(max(xs) - min(xs)) / 1000:,.0f} x "
          f"{(max(ys) - min(ys)) / 1000:,.0f} m   ← 자동의 «기본 범위»")

    frames = A.detect_sheet_frames(heads)
    print(f"\n  A 의 도면 장 나누기: {len(frames)}장")
    for f in frames:
        x0, y0, x1, y1 = [float(v) for v in f["bbox"]]
        n = sum(1 for h in heads
                if x0 <= h.pos[0] <= x1 and y0 <= h.pos[1] <= y1)
        print(f"    장 {f.get('index')} · 헤드 {n:5,d} · "
              f"{(x1 - x0) / 1000:7.1f} x {(y1 - y0) / 1000:6.1f} m")

    if len(frames) > 1:
        print("\n  ★한 파일에 여러 장이다 — 영역을 안 그리면 자동의 기본 범위가")
        print("    장을 가로질러, 최불리가 서로 다른 도면의 헤드를 섞어 뽑는다.")
    else:
        print("\n  한 장짜리 — 헤드 전부를 범위로 삼아도 된다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
