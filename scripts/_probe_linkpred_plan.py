# -*- coding: utf-8 -*-
"""연결복원(link-pred)을 «평면도» 에 걸어 본다 — 조각난 B1F 가 붙는가.

찾은 사실: `calibration/linkpred_integrate.py` 의 합의 엔진은 이미 있고
`/api/remote30/system/connection_review` 로 열려 있으나 **계통도 전용**이다
(폼 키가 `system_dxf_file`). 평면도 차선은 이 엔진을 안 쓴다.

B1F 의 병이 바로 이 엔진이 노리는 것이다:
  · 끝점끼리는 안 닿는다 → auto_snap_eps 를 1,200mm 까지 올려도 최대 조각 8%
  · 틈이 «끝점 ↔ 배관 중간» 이다 → 휴리스틱 bridge 가 구조적으로 못 만든다
  · 그게 이 엔진의 kind="te" (T분기) — 등급 C 「ML 단독」

그래서 세 가지를 잰다:
  ① 이 엔진이 평면도 entity 로 돌기는 하는가 (계통도용이라 안 돌 수도)
  ② 제안이 몇 건이고 tt(끝점↔끝점) / te(T분기) 비율이 어떤가
  ③ ★그 제안을 그래프에 넣으면 조각 306개가 실제로 줄어드는가

③ 이 이 조사의 값이다. 안 줄면 이 엔진도 답이 아니다.

    python scripts/_probe_linkpred_plan.py [도면.dxf] [--cut 0.45] [--mode allt]
"""
from __future__ import annotations

import argparse
import sys
import time
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
PLAN = ROOT / "data" / "uploads" / "B1F 현장조사 소화설비 평면도.dxf"


def _components(graph, extra=()):
    """graph(dict[node]->set) + 추가 간선으로 조각을 센다. (조각수, 최대조각)"""
    par: dict = {}

    def find(x):
        r = x
        while par.get(r, r) != r:
            r = par[r]
        while par.get(x, x) != x:
            par[x], x = r, par[x]
        return r

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            par[ra] = rb

    for u, vs in graph.items():
        par.setdefault(u, u)
        for v in vs:
            par.setdefault(v, v)
            union(u, v)
    for a, b in extra:
        par.setdefault(a, a)
        par.setdefault(b, b)
        union(a, b)
    sz = Counter(find(n) for n in par)
    return len(sz), (sz.most_common(1)[0][1] if sz else 0), len(par)


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    ap = argparse.ArgumentParser()
    ap.add_argument("dxf", nargs="?", default=str(PLAN))
    ap.add_argument("--cut", type=float, default=0.45)
    ap.add_argument("--mode", default="allt")
    a = ap.parse_args()

    for p in (str(ROOT), str(ROOT / "core"), str(ROOT / "calibration")):
        if p not in sys.path:
            sys.path.insert(0, p)
    import remote30_prototype as A          # noqa: F401  (경로 준비)
    import linkpred_integrate as li
    from clean_candidate_survey import _raw_graph

    dxf = Path(a.dxf)
    if not dxf.is_file():
        print("도면 없음:", dxf)
        return 1

    pair = li.load_model(a.mode)
    if pair is None:
        print(f"모델 없음 (mode={a.mode})")
        return 1
    model, feats = pair
    print(f"{dxf.name}  ·  모델 {a.mode}  ·  ml_cut {a.cut}\n")

    t0 = time.perf_counter()
    bundle = A.parse_dxf_bundle_cached(dxf)
    ents = bundle.entities
    t_parse = time.perf_counter() - t0

    # ── raw 그래프 (엔진이 안에서 쓰는 것과 같은 것) — 개선 전 기준선
    t0 = time.perf_counter()
    graph, edge_len, scale = _raw_graph(ents)
    t_graph = time.perf_counter() - t0
    n0, big0, tot0 = _components(graph)
    print(f"■ 개선 전 raw 그래프  (파싱 {t_parse:.1f}s · 그래프 {t_graph:.1f}s)")
    print(f"    절점 {tot0:,} · 간선 {len(edge_len):,} · scale {scale:.4f}")
    print(f"    조각 {n0:,}개 · 최대 조각 {big0:,} ({big0 / max(1, tot0) * 100:.1f}%)\n")

    # ── 엔진
    t0 = time.perf_counter()
    try:
        res = li.reconcile_entities(ents, model, feats, ml_cut=a.cut)
    except Exception as exc:  # noqa: BLE001
        print(f"★엔진이 평면도에서 실패 — {type(exc).__name__}: {exc}")
        return 2
    t_ml = time.perf_counter() - t0
    if res is None:
        print("★후보 없음 (tips/segs 0) — 평면도에선 못 쓴다")
        return 2

    A_, CONF, B_, C_ = res["A"], res["CONF"], res["B"], res["C"]

    def kc(items):
        c = Counter(it[2]["kind"] for it in items)
        return c.get("tt", 0), c.get("te", 0)

    a_tt, a_te = kc(A_)
    f_tt, f_te = kc(CONF)
    c_tt, c_te = kc(C_)
    print(f"■ 연결복원 엔진  ({t_ml:.1f}s)")
    print(f"    본관 edge {res['n_seg']:,} · 끝단 {res['n_tip']:,} · "
          f"후보쌍 {res['n_cand']:,} · med_edge {res['med_edge']:.1f}")
    print(f"    휴리스틱 bridge {res['H']:,}개 (끝단 엮임 {res['n_htip']:,})")
    print(f"    A 합의       {len(A_):>5,}   tt {a_tt:>5,} · te {a_te:>5,}")
    print(f"    CONFLICT     {len(CONF):>5,}   tt {f_tt:>5,} · te {f_te:>5,}")
    print(f"    B 휴리단독   {len(B_):>5,}")
    print(f"    C ML단독     {len(C_):>5,}   tt {c_tt:>5,} · te {c_te:>5,}"
          f"   ★te = 휴리스틱이 구조적으로 못 만드는 T분기\n")

    # ── ③ 제안을 넣으면 조각이 줄어드나 — 등급별로 누적해서 본다
    def pairs(items):
        return [(it[2]["a"], it[2]["b"]) for it in items]

    steps = [
        ("A 합의만",              pairs(A_)),
        ("A + C(ML단독)",         pairs(A_) + pairs(C_)),
        ("A + C + CONFLICT",      pairs(A_) + pairs(C_) + pairs(CONF)),
        ("전부 + B(휴리단독)",     pairs(A_) + pairs(C_) + pairs(CONF)
                                   + [k for k, _g in B_]),
    ]
    print("■ ★제안을 그래프에 넣으면 조각이 줄어드나")
    print(f"    {'적용':<22} {'조각':>9} {'최대 조각':>11} {'비율':>8}")
    print("    " + "-" * 54)
    print(f"    {'(넣기 전)':<22} {n0:>9,} {big0:>11,} "
          f"{big0 / max(1, tot0) * 100:>7.1f}%")
    for name, ex in steps:
        n1, big1, tot1 = _components(graph, ex)
        print(f"    {name:<22} {n1:>9,} {big1:>11,} "
              f"{big1 / max(1, tot1) * 100:>7.1f}%")

    # ── te 제안의 길이 분포 — 실제 틈 크기와 맞는지
    if C_:
        import statistics
        te = [it for it in C_ if it[2]["kind"] == "te"]
        if te:
            ds = sorted(((it[2]["a"][0] - it[2]["b"][0]) ** 2
                         + (it[2]["a"][1] - it[2]["b"][1]) ** 2) ** 0.5
                        for it in te)
            print(f"\n    C·te 제안 거리  중앙값 {statistics.median(ds):,.0f}mm · "
                  f"p90 {ds[int(len(ds) * .9)]:,.0f}mm · 최대 {ds[-1]:,.0f}mm")
            pr = sorted(float(it[0]) for it in te)
            print(f"    C·te 확률       중앙값 {statistics.median(pr):.2f} · "
                  f"최대 {pr[-1]:.2f}")

    print("\n  최대 조각 비율이 크게 오르면 이 엔진이 B1F 의 답이다.")
    print("  안 오르면 «틈» 이 이 엔진이 보는 범위(search_factor·med_edge) 밖이다.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
