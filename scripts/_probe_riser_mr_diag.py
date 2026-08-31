# -*- coding: utf-8 -*-
"""계통도·기계실 진단 — 「인식이 안 된다 · 경로 추출이 안 된다」를 가른다.

증상 두 개는 원인이 다를 수 있고, 그러면 고칠 곳도 다르다. 한 번에 재서 어느
쪽인지 먼저 정한다:

    ① 파서가 무엇을 봤나        entity 수 · 레이어 분류 (PIPE/HEAD/ALARM/OTHER)
    ② 헤드 인식                 detect_heads 후보 · 신뢰도 띠 · 어느 레이어에서
    ③ 배관 그래프               노드·간선 · **연결성분** · 가장 큰 덩이의 비중
    ④ 경로가 설 수 있나          가장 큰 덩이 안에서 지름(가장 먼 두 점) 거리

★③이 요점이다. 계통도는 «도면» 이 아니라 «그림» 이라 선이 끊겨 그려지는 일이
  잦다(LH306 계통도에서 이미 겪었다 — PIPE 레이어가 조각나 단일망 추출 불가).
  연결성분이 수백 개면 인식을 아무리 고쳐도 경로는 안 선다. 반대로 성분이 하나면
  경로가 안 서는 이유는 다른 데 있다.

    python scripts/_probe_riser_mr_diag.py [도면.dxf ...]
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
DEF = [
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 계통도.dxf",
    ROOT / "data" / "uploads" / "1. 입력도면 대명동 단위세대 기계실.dxf",
]


def _components(adj: dict) -> list:
    """연결성분을 큰 순서로 — `_build_graph` 가 주는 좌표 인접리스트 위에서.

    ★이 한 수가 진단의 요점이다. 계통도는 «도면» 이 아니라 «그림» 이라 선이
      끊겨 그려지곤 한다(LH306 계통도에서 이미 겪었다). 조각이 수백이면 인식을
      아무리 고쳐도 경로는 안 선다.
    """
    seen, out = set(), []
    for start in adj:
        if start in seen:
            continue
        stack, n = [start], 0
        seen.add(start)
        while stack:
            cur = stack.pop()
            n += 1
            for nxt in adj.get(cur, ()):
                if nxt not in seen:
                    seen.add(nxt)
                    stack.append(nxt)
        out.append(n)
    return sorted(out, reverse=True)


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A

    files = [Path(x) for x in sys.argv[1:]] or DEF
    for dxf in files:
        if not dxf.is_file():
            print(f"\n■ {dxf.name} — 파일 없음")
            continue
        print(f"\n{'=' * 74}")
        print(f"■ {dxf.name}  ({dxf.stat().st_size / 1048576:.1f} MB)")
        print("=" * 74)

        bundle = A.parse_dxf_bundle_cached(dxf)
        ents = bundle.entities
        # ★번들이 매긴 분류를 그대로 쓴다 — 이름으로 다시 매기면 파서 끝의
        #   레이어 승격이 사라진다(recon 이 쓰는 그 방식과 같게).
        layers = {ly.get("name"): (ly.get("auto_category") or "OTHER")
                  for ly in (bundle.layers or [])}

        # ── ① 파서가 본 것
        by_cat = Counter()
        by_layer = Counter()
        by_type = Counter()
        for e in ents:
            ly = str(e.get("l") or "0")
            by_layer[ly] += 1
            by_cat[layers.get(ly, "OTHER")] += 1
            by_type[str(e.get("t") or "?")] += 1
        print(f"\n① 파서 — entity {len(ents):,} · 레이어 {len(by_layer)}")
        print(f"   분류  {dict(by_cat)}")
        print(f"   종류  {dict(by_type.most_common(8))}")
        print("   레이어 상위 12")
        for nm, k in by_layer.most_common(12):
            print(f"     {k:>7,}  [{layers.get(nm, 'OTHER'):<5}] {nm}")

        # ── ② 헤드 인식
        heads = A.detect_heads(ents, layers)
        band = Counter()
        head_layer = Counter()
        for h in heads:
            c = float(h.confidence)
            band["높음(≥0.9)" if c >= 0.9
                 else "중간(≥0.75)" if c >= 0.75 else "낮음"] += 1
            head_layer[str(h.layer or "")] += 1
        print(f"\n② 헤드 인식 — 후보 {len(heads):,}")
        print(f"   띠  {dict(band)}")
        for nm, k in head_layer.most_common(6):
            print(f"     {k:>5}  {nm}")
        if not heads:
            print("   ★후보 0 — 이 도면에서는 헤드 기호를 못 찾는다.")

        # ── ③ 배관 그래프
        adj, lens = A._build_graph(ents, layer_categories=layers)
        n_edge = sum(len(v) for v in adj.values()) // 2
        total_m = sum(lens.values()) / 1000.0 if lens else 0.0
        print(f"\n③ 배관 그래프 — 노드 {len(adj):,} · 간선 {n_edge:,} · "
              f"총연장 {total_m:,.1f} m")
        if adj:
            comps = _components(adj)
            top = comps[0] if comps else 0
            iso = sum(1 for c in comps if c <= 2)
            print(f"   연결성분 {len(comps):,}개 · 가장 큰 덩이 {top:,} 노드 "
                  f"({top / max(1, len(adj)) * 100:.1f}%) · 조각(≤2노드) {iso:,}")
            print(f"   상위 덩이 크기 {comps[:12]}")
            if len(comps) > 20 and top / max(1, len(adj)) < 0.6:
                print("   ★조각이 많고 큰 덩이가 절반도 안 된다 — 경로 추출은")
                print("     여기서 막힌다. 인식을 고쳐도 «이어져 있지 않으면»")
                print("     경로는 안 선다.")
        else:
            print("   ★간선 0 — 배관으로 볼 선 자체가 없다.")

        # ── ④ 헤드가 그 선에 붙어 있나
        if heads and adj:
            import math
            xy = list(adj.keys())
            near = defaultdict(int)
            sample = heads[:300]             # 표본 — 전수는 비싸다
            for h in sample:
                hx, hy = float(h.pos[0]), float(h.pos[1])
                d = min(math.hypot(x - hx, y - hy) for x, y in xy)
                for lim in (100, 300, 1000, 3000):
                    if d <= lim:
                        near[lim] += 1
                        break
                else:
                    near["멀다"] += 1
            n = len(sample)
            print(f"\n④ 헤드↔배관 최근접 (표본 {n}개)")
            for lim in (100, 300, 1000, 3000):
                if near[lim]:
                    print(f"     ≤{lim:>5}mm  {near[lim]:>4}개 "
                          f"({near[lim] / n * 100:.0f}%)")
            if near["멀다"]:
                print(f"     >3000mm  {near['멀다']:>4}개 "
                      f"({near['멀다'] / n * 100:.0f}%)")
            if near[100] + near[300] == 0:
                print("   ★헤드가 배관선에 안 붙어 있다 — 물길 판정이 전부 «마름»")
                print("     이 된다. 계통도는 기호와 선이 떨어져 그려지곤 한다.")

        # ── ⑤ «브로드» 로 하면 이어지나
        #
        # `_build_graph` 에는 이미 기하 폴백이 있다 — 이름이 안 붙은 레이어에
        # 작도된 도면(LH306)을 살리려고 넣은 것이다. 그런데 조건이
        # **「간선이 0이면」** 이라, 이 두 도면처럼 «조금은 잡히는데 쓸모없는»
        # 경우에는 안 돈다. 그 폴백이 돌면 어떻게 되는지 여기서 흉내 내 본다.
        nonpipe = getattr(A, "NON_PIPE_GEOMETRY_CATS", set())
        broad = {nm: ("PIPE" if c not in nonpipe else c)
                 for nm, c in layers.items()}
        badj, blens = A._build_graph(ents, layer_categories=broad)
        if badj:
            bc = _components(badj)
            btop = bc[0] if bc else 0
            print(f"\n⑤ 브로드(비-배관 카테고리만 제외) — 노드 {len(badj):,} · "
                  f"간선 {sum(len(v) for v in badj.values()) // 2:,} · "
                  f"총연장 {sum(blens.values()) / 1000.0:,.1f} m")
            print(f"   연결성분 {len(bc):,}개 · 가장 큰 덩이 {btop:,} "
                  f"({btop / max(1, len(badj)) * 100:.1f}%)")
            print(f"   상위 덩이 {bc[:8]}")

        # ── ⑥ 어느 레이어가 «진짜 배관» 인가 — 한 장씩 얹어 본다.
        #
        # 이름 사전이 못 알아본 레이어 중 어느 것을 배관으로 치면 망이 서는지
        # 직접 잰다. 사람이 도면을 열어 보고 정하는 그 판단을 수치로 돕는다.
        cands = [(nm, k) for nm, k in by_layer.most_common(30)
                 if layers.get(nm, "OTHER") == "OTHER" and k >= 40]
        if cands:
            print(f"\n⑥ OTHER 레이어를 배관에 «한 장씩» 더해 보면 "
                  f"(현재 기준 {top:,}노드)")
            print(f"   {'레이어':<18}{'entity':>8}{'노드':>8}{'큰덩이':>8}"
                  f"{'성분':>7}{'연장(m)':>10}")
            rows = []
            for nm, k in cands:
                try_layers = dict(layers)
                try_layers[nm] = "PIPE"
                a2, l2 = A._build_graph(ents, layer_categories=try_layers)
                if not a2:
                    continue
                c2 = _components(a2)
                rows.append((c2[0], nm, k, len(a2), len(c2),
                             sum(l2.values()) / 1000.0))
            for t2, nm, k, n2, nc, m2 in sorted(rows, reverse=True):
                gain = t2 - top
                mark = "  ★크게 이어진다" if gain >= top else ""
                print(f"   {nm:<18}{k:>8,}{n2:>8,}{t2:>8,}{nc:>7,}"
                      f"{m2:>10,.1f}{mark}")

            # ── ⑦ 그 레이어가 «기존 배관을 잇는가» — 제 노이즈만 들고 오나
            #
            # 큰 덩이가 커졌다고 배관인 것은 아니다. 벽 윤곽선도 길게 이어진다.
            # 진짜 배관 레이어라면 **원래 배관 조각들을 한 덩이로 모아야** 한다.
            # 그것을 직접 센다 — 새 큰 덩이 안에 원래 PIPE 노드가 몇이나 들어왔나.
            if rows:
                best = max(rows)[1]
                try_layers = dict(layers)
                try_layers[best] = "PIPE"
                a3, _l3 = A._build_graph(ents, layer_categories=try_layers)
                # 새 그래프의 가장 큰 덩이를 실제로 뽑는다.
                seen, big = set(), []
                for st in a3:
                    if st in seen:
                        continue
                    stack, grp = [st], []
                    seen.add(st)
                    while stack:
                        cur = stack.pop()
                        grp.append(cur)
                        for nx in a3.get(cur, ()):
                            if nx not in seen:
                                seen.add(nx)
                                stack.append(nx)
                    if len(grp) > len(big):
                        big = grp
                bigset = set(big)
                inside = sum(1 for p in adj if p in bigset)
                print(f"\n⑦ «{best}» 를 배관으로 치면 — 원래 배관 노드 "
                      f"{len(adj):,}개 중 {inside:,}개가 새 큰 덩이 안에 든다 "
                      f"({inside / max(1, len(adj)) * 100:.0f}%)")
                if inside >= len(adj) * 0.5:
                    print("   → 흩어져 있던 배관 조각을 실제로 «잇는다». 이 레이어는")
                    print("     배관으로 봐야 한다.")
                else:
                    print("   → 원래 배관은 거의 안 들어왔다. 큰 덩이가 커진 것은")
                    print("     이 레이어가 «제 노이즈» 를 들고 온 것뿐이다.")

                # ── ⑧ 그 그래프에서 헤드가 붙나 — ④를 다시 잰다.
                #
                # ★④ 는 «거의 빈 그래프» 를 상대로 잰 값이라 그 자체로는 결론이
                #   못 된다. 배관이 제대로 선 뒤에 다시 재야 «헤드가 떨어져 있다» 를
                #   말할 수 있다. 안 그러면 없는 선까지의 거리를 재는 셈이다.
                if heads:
                    import math
                    xy3 = list(a3.keys())
                    n3 = defaultdict(int)
                    sample = heads[:300]
                    for h in sample:
                        hx, hy = float(h.pos[0]), float(h.pos[1])
                        d = min(math.hypot(x - hx, y - hy) for x, y in xy3)
                        for lim in (100, 300, 1000, 3000):
                            if d <= lim:
                                n3[lim] += 1
                                break
                        else:
                            n3["멀다"] += 1
                    tot = len(sample)
                    near_now = n3[100] + n3[300]
                    print(f"\n⑧ 그 그래프에서 헤드↔배관 (표본 {tot}개) — "
                          f"300mm 안 {near_now}개 "
                          f"({near_now / tot * 100:.0f}%) · "
                          f"3m 밖 {n3['멀다']}개")
    return 0


if __name__ == "__main__":
    sys.exit(main())
