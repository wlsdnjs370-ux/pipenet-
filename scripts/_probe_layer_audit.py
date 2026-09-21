# -*- coding: utf-8 -*-
"""레이어 분류 감사 — 「주차 구획선이 배관으로 들어오나」를 5분에 가린다.

세 가지를 한 번에 본다:
  ① 주차장에서 흔한 이름들이 실제로 어느 분류로 떨어지나 (합성 이름 시험)
  ② 이 도면에서 PIPE 로 분류된 레이어에 «주차 냄새» 나는 이름이 있나
  ③ 이름 사전을 못 믿어 기하 폴백이 도는가 (그러면 바닥이 통째로 배관 후보)

    python scripts/_probe_layer_audit.py [도면.dxf ...]
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

# 주차장 도면에서 흔한 이름 — 배관으로 잡히면 안 되는 것들.
SUSPECT = [
    "A-PARKING-SPACE", "SPACE-EV", "A-SPACE-LINE", "PARKING", "STALL",
    "CAR", "EV-CHARGER", "A-SPACE", "TRANSPARENT", "SPEC", "SPOT",
    "주차구획", "주차선", "전기차주차",
    # 이것들은 «배관으로 잡혀야» 정상 — 대조군
    "6-소화-SP-메인", "SP1", "SP-2", "HSP", "LSP", "A-SP", "SP배관",
]
# 주차·바닥 냄새 — 이 낱말이 든 레이어가 PIPE 면 의심한다.
SMELL = ("SPACE", "PARK", "STALL", "CAR", "주차", "구획", "바닥", "FLOOR",
         "SLAB", "EV")


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    # ── ① 이름 사전이 실제로 무엇을 하나
    from remote30_prototype import Remote30Settings, layer_match
    live = Remote30Settings is not None and layer_match is not None
    print(f"■ 이름 사전 경로 — {'정상(토큰 경계 매칭)' if live else '★폴백(부분일치)'}")
    if not live:
        print("  ★Remote30Settings/layer_match 를 못 불러왔다. 이때만 «SP» 가")
        print("    부분일치로 돌아 SPACE 를 먹는다. 이것부터 고쳐야 한다.")
    print()
    print(f"  {'이름':<24} {'분류':<8}")
    print("  " + "-" * 34)
    bad = 0
    for nm in SUSPECT:
        c = A._categorize_layer(nm)
        smell = any(s in nm.upper() for s in SMELL)
        mark = ""
        if smell and c == "PIPE":
            mark = "  ★주차 이름이 배관으로"
            bad += 1
        print(f"  {nm:<24} {c:<8}{mark}")
    print(f"\n  주차 이름이 배관으로 분류된 것 {bad}건")

    # ── ②③ 실제 도면
    args = sys.argv[1:] or [
        str(ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"),
        str(ROOT / "data" / "_genz_dxf" / "S2_죽전_지하주차장소화설비.dxf"),
    ]
    for f in args:
        dxf = Path(f)
        if not dxf.is_file():
            print(f"\n■ {dxf.name} — 파일 없음")
            continue
        print(f"\n■ {dxf.name}")
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = bundle.entities
        from collections import Counter
        cnt = Counter()
        cat = {}
        for e in ents:
            ly = str(e.get("l") or "0")
            cnt[ly] += 1
            if ly not in cat:
                try:
                    cat[ly] = A._categorize_layer(ly)
                except Exception:  # noqa: BLE001
                    cat[ly] = "OTHER"
        pipes = [(n, cnt[n]) for n, c in cat.items() if c == "PIPE"]
        pipes.sort(key=lambda t: -t[1])
        print(f"  PIPE 로 분류된 레이어 {len(pipes)}개")
        for n, k in pipes:
            smell = any(s in n.upper() for s in SMELL)
            print(f"    {k:>7,}  {n}" + ("   ★주차/바닥 냄새" if smell else ""))

        # ③ 기하 폴백이 도는가 — PIPE 만으로 간선이 서는지 본다.
        try:
            from remote30_prototype import build_graph, filter_pipenet_only
            strict = filter_pipenet_only(ents, cat)
            print(f"  PIPE·HEAD·ALARM 만 남긴 entity {len(strict):,} / {len(ents):,}")
        except Exception as exc:  # noqa: BLE001
            print(f"  (엄격 필터 확인 건너뜀: {type(exc).__name__}: {exc})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
