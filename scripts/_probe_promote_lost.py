# -*- coding: utf-8 -*-
"""모듈 F 자동 차선이 «레이어 승격» 을 버리고 있는가.

`routes/module_f/auto.py::open_plan` 의 도크스트링은 이렇게 경고한다:

    시각화용 파서와 섞으면 레이어 승격(헤드 틈 지문으로 PIPE 로 올리는 것)
    같은 판정이 통째로 빠진다.

그래 놓고 바로 아래에서 `_categorize_layer(name)` 로 분류를 **다시 만든다**.
승격은 entity 를 ``"<원이름> (배관 승격)"`` 파생 레이어로 옮기는 방식이라,
그 파생 이름이 이름 사전에서 PIPE 로 안 떨어지면 승격은 통째로 증발한다.

사용자 신고와 정확히 맞는다 —
    「현장조사#셔터 x 분홍 레이어는 배관으로써 인식을 못하는 것 같다만?」
그 레이어가 바로 `_promote_headgap_pipe_layers` 도크스트링이 실측 사례로 든
것이다(가지관 85.6m · 헤드 15개).

세 가지를 잰다:
  ① 파생 이름이 `_categorize_layer` 에서 뭘로 떨어지나
  ② 이 도면에서 승격이 몇 건 일어났나 (A 의 번들 기준)
  ③ ★F 방식(이름 재분류)과 A 방식(번들 분류)의 PIPE entity 수 차이

    python scripts/_probe_promote_lost.py [도면.dxf ...]
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEFAULT = [
    ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf",
    ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf",
    ROOT / "samples" / "dxf" / "계통도_LH_306.dxf",
]
PROBE_NAMES = [
    "현장조사#셔터",
    "현장조사#셔터 (배관 승격)",
    "6-소화-가지관",
    "0 (배관 승격)",
    "A-WALL (배관 승격)",
]


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    # ── ① 파생 이름이 이름 사전에서 뭘로 떨어지나
    print("■ ① 파생 레이어 이름을 이름 사전에 그냥 물어보면")
    print(f"    {'분류':<9} 이름")
    print("    " + "-" * 44)
    lost = 0
    for n in PROBE_NAMES:
        c = A._categorize_layer(n)
        mark = ""
        if "(배관 승격)" in n and c != "PIPE":
            mark = "   ★승격이 증발한다"
            lost += 1
        print(f"    {c:<9} {n}{mark}")
    print(f"\n    승격 이름인데 PIPE 가 아닌 것 {lost}건\n")

    args = [Path(a) for a in sys.argv[1:]] or DEFAULT
    for dxf in args:
        if not dxf.is_file():
            print(f"■ {dxf.name} — 파일 없음\n")
            continue
        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = list(bundle.entities or ())
        promoted = list(bundle.promoted_layers or ())

        # A 방식 — 번들이 이미 정한 분류(승격 반영됨)
        cat_a = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
        # F 방식 — entity 레이어 이름으로 다시 분류(auto.py::open_plan)
        names = {str(e.get("l") or "0") for e in ents}
        cat_f = {}
        for n in names:
            try:
                cat_f[n] = A._categorize_layer(n)
            except Exception:  # noqa: BLE001
                cat_f[n] = "OTHER"

        def pipe_ents(cat):
            return sum(1 for e in ents
                       if cat.get(str(e.get("l") or "0"), "OTHER") == "PIPE")

        pa, pf = pipe_ents(cat_a), pipe_ents(cat_f)
        print(f"■ {dxf.name}")
        print(f"    entity {len(ents):,} · 레이어 {len(names):,}")
        print(f"    ② 승격 기록 {len(promoted)}건")
        for r in promoted[:8]:
            print(f"        {r.get('prev_category','?'):<8} "
                  f"{str(r.get('layer')):<28} → {r.get('pipe_layer')}  "
                  f"(지문 {r.get('headgap_count','?')} · "
                  f"entity {r.get('entity_count','?')})")
        if len(promoted) > 8:
            print(f"        … 그 외 {len(promoted) - 8}건")

        # 두 분류가 갈리는 레이어
        diff = sorted(n for n in names
                      if cat_a.get(n, "OTHER") != cat_f.get(n, "OTHER"))
        print(f"    ③ A 방식 PIPE entity {pa:,}  ·  F 방식 PIPE entity {pf:,}"
              f"   차이 {pa - pf:+,}")
        if diff:
            print(f"       분류가 갈리는 레이어 {len(diff)}개")
            for n in diff[:10]:
                cnt = sum(1 for e in ents if str(e.get("l") or "0") == n)
                print(f"        {cnt:>7,}  {cat_a.get(n,'?'):<8}(A) vs "
                      f"{cat_f.get(n,'?'):<8}(F)   {n}")
            if len(diff) > 10:
                print(f"        … 그 외 {len(diff) - 10}개")
        else:
            print("       분류가 갈리는 레이어 없음 — 이 도면은 손해가 없다")
        print()

    print("  A 방식 PIPE 가 F 보다 많으면, 자동 차선은 그만큼의 배관을 못 보고 있다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
