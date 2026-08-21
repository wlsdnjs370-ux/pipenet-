# -*- coding: utf-8 -*-
"""모듈 A 헤드 인식 일반화 진단 — 타현장 도면에서 왜 못 잡는지 실측한다.

지금 구조는 `_categorize_layer` 가 레이어 이름을 HEAD 로 찍어 줘야만
`detect_heads` 의 R1~R3 이 돈다. 현장마다 레이어 이름 규칙이 달라서 이 문턱을
못 넘으면 헤드가 0 이 된다 — 그것이 실제로 일어나는지, 대신 무엇을 근거로
삼을 수 있는지(블록 이름·원 반지름 분포·삽입 반복성)를 도면별로 뽑는다.
"""
from __future__ import annotations

import collections
import math
import os
import sys
import time
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "core"))

DXF_DIR = ROOT / "data" / "_genz_dxf"


def main():
    from remote30_prototype import (
        _categorize_layer, detect_heads, filter_pipenet_only,
        parse_dxf_bundle)

    files = sorted(DXF_DIR.glob("*.dxf"))
    if len(sys.argv) > 1:
        files = [f for f in files if sys.argv[1] in f.name]

    for path in files:
        print("\n" + "=" * 78)
        print(f"■ {path.name}  ({path.stat().st_size/1024/1024:.1f} MB)")
        t0 = time.perf_counter()
        try:
            bundle = parse_dxf_bundle(path)
        except Exception as exc:  # noqa: BLE001
            print(f"  !! 파싱 실패: {type(exc).__name__}: {exc}")
            continue
        ents = filter_pipenet_only(bundle)
        cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
        _ = _categorize_layer  # 분류는 파서가 이미 붙여 준다
        print(f"  파싱 {time.perf_counter()-t0:.1f}s · 레이어 {len(cats)}"
              f" · pipenet 엔티티 {len(ents):,}")

        by_cat = collections.Counter(cats.values())
        print("  레이어 분류:", dict(by_cat))
        heads_ly = [ly for ly, c in cats.items() if c == "HEAD"]
        print(f"  HEAD 레이어 {len(heads_ly)}개: {heads_ly[:6]}")

        det = detect_heads(ents, cats)
        print(f"  ▶ detect_heads = {len(det)} 개")
        if det:
            print("     kind 분포:", dict(collections.Counter(
                d.kind.split("(")[0] for d in det)))

        # ── 무엇을 근거로 삼을 수 있나 ─────────────────────────────
        # (1) 블록 이름 — 이름에 head/헤드/SP 가 든 INSERT 가 몇 번 쓰였나
        ins = collections.Counter()
        ins_layer = collections.defaultdict(collections.Counter)
        for en in ents:
            if en["t"] == "I":
                bn = str(en.get("n", ""))
                ins[bn] += 1
                ins_layer[bn][en["l"]] += 1
        cand_bn = [(n, c) for n, c in ins.most_common(40)
                   if any(k in n.upper() for k in
                          ("HEAD", "헤드", "SP", "SPR", "K-", "K1", "K2",
                           "펜던트", "PENDENT", "UPRIGHT", "측벽"))]
        print(f"  INSERT 블록 {len(ins)}종 / {sum(ins.values()):,}회")
        print("     헤드 의심 블록:",
              [(n, c, list(ins_layer[n])[:1]) for n, c in cand_bn[:8]] or "없음")
        print("     최다 블록:", ins.most_common(6))

        # (2) 원 — 반지름 분포 (레이어 무관). 헤드 마커는 같은 반지름이 대량 반복.
        circ = collections.Counter()
        circ_layer = collections.defaultdict(collections.Counter)
        for en in ents:
            if en["t"] == "C":
                r = round(float(en.get("r", 0)), 1)
                circ[r] += 1
                circ_layer[r][en["l"]] += 1
        small = [(r, c) for r, c in circ.most_common(30) if 5.0 <= r <= 400.0]
        print(f"  CIRCLE {sum(circ.values()):,}개 · 반지름 {len(circ)}종")
        print("     반복 많은 소형 반지름:",
              [(r, c, list(circ_layer[r])[:1]) for r, c in small[:8]] or "없음")

        # (3) HATCH
        hat = sum(1 for en in ents if en["t"] == "H")
        print(f"  HATCH {hat:,}개")

        # (4) 이름에 head 가 든 레이어인데 HEAD 로 분류 안 된 것
        missed = [(ly, cats[ly]) for ly in cats
                  if cats[ly] != "HEAD" and any(
                      k in ly.upper() for k in ("HEAD", "헤드", "SP-H", "SPH"))]
        if missed:
            print("  ※ 이름엔 헤드가 있는데 HEAD 로 안 잡힌 레이어:", missed[:8])


if __name__ == "__main__":
    main()
