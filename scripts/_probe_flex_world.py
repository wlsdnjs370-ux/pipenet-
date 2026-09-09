# -*- coding: utf-8 -*-
"""[신축배관] 그 레이어에 «무엇이» 있는가 — 선분만인가, 호도 있는가.

계측이 지시서와 어긋났다. 지시서는 원도면을 직접 재서 **85가닥 · 가닥당
4~9선분 · 663~723 mm · 총연장 57,271 mm** 라고 했는데, 찍기 판에서 재니
**320선분 · 200토막 · 4~375 mm · 총연장 44,554 mm** 다.

가닥 복원 방식을 두 번 고쳐도(격자 → eps 클러스터) 안 붙었으므로, 남은 가설은
「`w.segs` 에 그 가닥의 일부가 **없다**」 — 즉 폴리라인이 bulge(호)를 갖고 있어
`w.arcs` 로 갔다는 것이다. 유연 호스라면 그렇게 그리는 것이 자연스럽다.

    python scripts/_probe_flex_world.py
"""
from __future__ import annotations

import math
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
for _p in (str(ROOT), str(ROOT / "cad_project_editor_g")):
    if _p not in sys.path:
        sys.path.insert(0, _p)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

PLAN = ROOT / "routes" / "제출용[최종]" / "1. 입력도면 대명동 단위세대 평면도.dxf"
FLEX_WORDS = ("후렉시블", "후렉", "플렉", "flex", "fx")


def is_flex(layer) -> bool:
    s = str(layer or "").lower()
    return any(w.lower() in s for w in FLEX_WORDS)


def main() -> int:
    if not PLAN.is_file():
        print(f"표본 없음: {PLAN}")
        return 0
    from services.cad_import.pick.io import open_dxf
    w, key, _kn = open_dxf(str(PLAN))
    print(f"\n■ {key}")
    print(f"  세계 — 선분 {len(w.segs)} · 원본선분 {len(w.raw_segs)}"
          f" · 원 {len(w.circles)} · 호 {len(w.arcs)} · 문자 {len(w.texts)}")

    def tally(rows, name, length):
        by: dict = {}
        for r in rows:
            ly = str(r[0])
            if not is_flex(ly):
                continue
            d = by.setdefault((ly, r[1]), [0, 0.0])
            d[0] += 1
            d[1] += length(r)
        print(f"  [{name}] 신축배관 묶음 {len(by)}")
        for k, (n, L) in sorted(by.items()):
            print(f"      {k}  {n}개 · 연장 {L:,.0f} mm")
        return sum(v[0] for v in by.values()), sum(v[1] for v in by.values())

    n_seg, l_seg = tally(w.segs, "선분", lambda r: math.dist(r[2], r[3]))
    n_raw, l_raw = tally(w.raw_segs, "원본선분", lambda r: math.dist(r[2], r[3]))

    # 호 — (layer, color, x, y, r) + arc_ang[i] = (start, sweep) 도
    angs = list(getattr(w, "arc_ang", ()) or ())
    by: dict = {}
    for i, (ly, col, x, y, rad) in enumerate(w.arcs):
        if not is_flex(ly):
            continue
        sweep = abs(float(angs[i][1])) if i < len(angs) and angs[i] else 0.0
        d = by.setdefault((str(ly), col), [0, 0.0])
        d[0] += 1
        d[1] += math.radians(sweep) * float(rad)
    print(f"  [호] 신축배관 묶음 {len(by)}")
    for k, (n, L) in sorted(by.items()):
        print(f"      {k}  {n}개 · 호길이 {L:,.0f} mm")
    n_arc = sum(v[0] for v in by.values())
    l_arc = sum(v[1] for v in by.values())

    print(f"\n  ★신축배관 총연장 — 선분 {l_seg:,.0f} + 호 {l_arc:,.0f}"
          f" = {l_seg + l_arc:,.0f} mm   (지시서 57,271)")
    print(f"    개수 — 선분 {n_seg} · 호 {n_arc} · 원본선분 {n_raw}"
          f" (원본 연장 {l_raw:,.0f})")

    # 선분 길이 분포 — 「4~9 선분이 한 가닥」 인지 보려면 낱개 길이가 필요하다
    lens = sorted(math.dist(r[2], r[3]) for r in w.segs if is_flex(r[0]))
    if lens:
        q = Counter(int(v // 50) * 50 for v in lens)
        print(f"    선분 길이 분포(50 mm 단위) {dict(sorted(q.items()))}")
        print(f"    최소 {lens[0]:.1f} · 중앙 {lens[len(lens) // 2]:.1f}"
              f" · 최대 {lens[-1]:.1f}")

    # ── 호를 «현(chord)» 으로 바꿔 넣고 가닥을 다시 세어 본다.
    #    호가 빠진 채로는 가닥이 호마다 끊긴다 — 그것이 200토막의 정체인지 본다.
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "_fold", str(ROOT / "scripts" / "_probe_flex_fold.py"))
    fold = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(fold)

    seg_only = [(a, b) for ly, _c, a, b in w.segs if is_flex(ly)]
    arc_seg, arc_len = [], 0.0
    for i, (ly, _col, x, y, rad) in enumerate(w.arcs):
        if not is_flex(ly):
            continue
        ang = angs[i] if i < len(angs) else None
        if not ang:
            continue
        sa, sw = float(ang[0]), float(ang[1])
        p0 = (x + rad * math.cos(math.radians(sa)),
              y + rad * math.sin(math.radians(sa)))
        p1 = (x + rad * math.cos(math.radians(sa + sw)),
              y + rad * math.sin(math.radians(sa + sw)))
        arc_seg.append((p0, p1))
        arc_len += abs(math.radians(sw)) * rad

    for name, segs in (("선분만", seg_only),
                       ("선분+호(현)", seg_only + arc_seg)):
        ch, tg, rest = fold.strands(segs)
        n = Counter(len(p) - 1 for p in ch)
        ls = sorted(sum(math.dist(p[i], p[i + 1]) for i in range(len(p) - 1))
                    for p in ch)
        print(f"  [가닥 · {name}] {len(ch)}가닥 · 사슬아님 {len(tg)}"
              f" · 미포함 {len(rest)}")
        print(f"      가닥당 조각 {dict(sorted(n.items()))}")
        if ls:
            print(f"      가닥 길이 {ls[0]:.0f} ~ {ls[-1]:.0f}"
                  f" (중앙값 {ls[len(ls) // 2]:.0f}) · 합 {sum(ls):,.0f} mm")
        # ★표 length 는 좌표의 **L1(맨해튼)** 에서 온다(`planar.py:875`).
        #   가닥이 x·y 로 단조로우면 꺾인 경로의 L1 합 = 현의 L1 이다.
        #   그러면 접어도 표 length 가 안 변한다 — 새 통로가 필요 없다.
        l1 = sum(abs(p[i + 1][0] - p[i][0]) + abs(p[i + 1][1] - p[i][1])
                 for p in ch for i in range(len(p) - 1))
        l1c = sum(abs(p[-1][0] - p[0][0]) + abs(p[-1][1] - p[0][1]) for p in ch)
        mono = sum(1 for p in ch
                   if abs(sum(p[i + 1][0] - p[i][0]
                              for i in range(len(p) - 1)))
                   + abs(sum(p[i + 1][1] - p[i][1] for i in range(len(p) - 1)))
                   >= (sum(abs(p[i + 1][0] - p[i][0])
                           + abs(p[i + 1][1] - p[i][1])
                           for i in range(len(p) - 1)) - 1e-6))
        print(f"      ★L1 합 {l1:,.0f} · 현의 L1 {l1c:,.0f}"
              f" · 차 {l1c - l1:+,.0f} mm · 단조로운 가닥 {mono}/{len(ch)}")
    print(f"  (호를 진짜 호길이로 세면 총연장 {l_seg + arc_len:,.0f} mm)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
