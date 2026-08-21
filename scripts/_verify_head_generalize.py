# -*- coding: utf-8 -*-
"""헤드 인식 일반화 검증 — 타현장 6종 + 기준 도면에서 전후를 비교한다.

기대치는 도면별 «실제 헤드 수» 의 대리값으로 **HEAD 레이어의 INSERT 개수** 를
쓴다. 헤드 하나에 블록 하나가 원칙이므로, 인식 수가 그 값에 가까워야 한다.
기준 도면(대명동)은 회귀가 없어야 한다.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "core"))

CASES = [
    ("대명동(기준)", "routes/제출용[최종]/1. 입력도면 대명동 단위세대 평면도.dxf"),
    ("대명동 계통도", "routes/제출용[최종]/1. 입력도면 대명동 단위세대 계통도.dxf"),
    ("S1 대구오페라", "data/_genz_dxf/S1_대구오페라_지하층소화설비.dxf"),
    ("S2 죽전", "data/_genz_dxf/S2_죽전_지하주차장소화설비.dxf"),
    ("S3 청라포레스트", "data/_genz_dxf/S3_청라포레스트_지하주차장소화설비.dxf"),
    ("S4 대우이안", "data/_genz_dxf/S4_대우이안_지하4층소방시설.dxf"),
    ("S5 청라스타필드(시트)", "data/_genz_dxf/S5_청라스타필드_지하1층스프링클러.dxf"),
    ("S6 대구오페라 단위세대", "data/_genz_dxf/S6_대구오페라_단위세대.dxf"),
    ("S7 청라스타필드(내용)", "data/_genz_dxf/S7_청라스타필드_내용도면.dxf"),
]

FAILS: list[str] = []


def main():
    from remote30_prototype import (
        HEAD_CLUSTER_R_MIN, auto_head_cluster_r, detect_heads,
        filter_pipenet_only, parse_dxf_bundle)

    print(f"{'도면':26s} {'INSERT':>7s} {'인식':>7s} {'배율':>6s} "
          f"{'반경':>6s}  판정")
    print("-" * 78)
    for tag, rel in CASES:
        path = ROOT / rel
        if not path.exists():
            print(f"{tag:26s} {'—':>7s} {'—':>7s}  파일 없음 — 건너뜀")
            continue
        bundle = parse_dxf_bundle(path)
        cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
        ents = filter_pipenet_only(bundle)
        n_ins = sum(1 for e in ents
                    if e["t"] == "I" and cats.get(e["l"]) == "HEAD")
        det = detect_heads(ents, cats)

        # 이 도면에서 실제로 쓰인 반경을 다시 계산해 표에 보인다.
        from remote30_prototype import HeadDetection  # noqa: F401
        cand = []
        for e in ents:
            c = cats.get(e["l"], "OTHER")
            if c == "HEAD" and e["t"] == "I":
                cand.append(("block_match", (float(e["p"][0]), float(e["p"][1]))))
            elif c == "HEAD" and e["t"] == "C":
                r = float(e.get("r", 0))
                if 10.0 <= r <= 250.0:
                    cand.append(("circle_signature",
                                 (float(e["c"][0]), float(e["c"][1]))))

        class _S:
            __slots__ = ("kind", "pos")

            def __init__(self, k, p):
                self.kind, self.pos = k, p

        radius = auto_head_cluster_r([_S(k, p) for k, p in cand])

        xd = bundle.xref_diagnostics or {}
        n_det = len(det)
        ratio = (n_det / n_ins) if n_ins else float("nan")

        if xd.get("is_sheet"):
            verdict = "시트 감지 ✓" if n_det == 0 else "시트인데 헤드가 나옴?"
            if n_det:
                FAILS.append(f"{tag}: XREF 시트인데 헤드 {n_det}개")
        elif n_ins == 0:
            verdict = "HEAD INSERT 없음(원/해치 기반)"
        elif 0.85 <= ratio <= 1.25:
            verdict = "정합 ✓"
        elif ratio > 1.25:
            verdict = "과다 ✗"
            FAILS.append(f"{tag}: 인식 {n_det} vs INSERT {n_ins} ({ratio:.2f}배)")
        else:
            verdict = "과소 ✗"
            FAILS.append(f"{tag}: 인식 {n_det} vs INSERT {n_ins} ({ratio:.2f}배)")

        print(f"{tag:26s} {n_ins:7,d} {n_det:7,d} "
              f"{('%.2f' % ratio) if n_ins else '—':>6s} "
              f"{radius:6.0f}  {verdict}")
        if xd.get("is_sheet"):
            print(f"{'':26s}   → {xd.get('message','')[:100]}")
            got = [os.path.basename(p) for p in (xd.get("resolved") or [])]
            print(f"{'':26s}   → 옆에서 찾은 참조 파일: {got or '없음'}")


if __name__ == "__main__":
    main()
    print()
    if FAILS:
        print(f"실패 {len(FAILS)}건")
        for f in FAILS:
            print("  -", f)
        raise SystemExit(1)
    print("전 도면 통과")
