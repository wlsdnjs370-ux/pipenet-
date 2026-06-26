# -*- coding: utf-8 -*-
"""연결복원 통합 — 휴리스틱(거리 bridge) × ML(link-prediction) 합의 등급화.

기존 휴리스틱(`_bridge_components`: 거리 기반 끝점↔끝점 단일연결)과 v2 ML ranker를
**한 raw 그래프 위에서** 같이 돌려, 복원 연결 후보를 신뢰등급으로 나눈다.

  · A 합의      : 휴리스틱 ∧ ML 둘 다 같은 끝점쌍 제안       → 고신뢰(실선/자동 후보)
  · B 휴리스틱단독: 휴리스틱만 제안(ML 낮음/불일치)          → 점선 검수
  · C ML단독    : ML만 제안 — 특히 T분기(끝점↔edge)는        → 점선 검수
                  휴리스틱이 구조적으로 못 만드는 연결

설계 원칙(1단계, advisory overlay): **추출 계산망은 건드리지 않는다.** 이 엔진은
복원 연결의 신뢰등급과 ML 전용 제안을 산출하는 정성 오버레이일 뿐, 그래프를 변형해
수리계산에 자동 주입하지 않는다(0.65 정직 CV 단계 — 사람 검수 전제).

DXF 엔 ground truth 가 없으므로 기하-only. 배포 모델 기본값 = v2_allt.

실행:
    python calibration/linkpred_integrate.py [allt|all|remote] [dxf ...]
"""
from __future__ import annotations

import sys
from collections import defaultdict
from pathlib import Path

import numpy as np

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

_ROOT = Path(__file__).resolve().parent.parent
_HERE = Path(__file__).resolve().parent
for _p in (str(_ROOT), str(_HERE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

import linkpred_data_v2 as lpd  # noqa: E402
from clean_candidate_survey import _raw_graph  # noqa: E402
from linkpred_transfer_v2 import graph_to_segs_tips  # noqa: E402
from remote30_prototype import parse_dxf_for_view, _bridge_components  # noqa: E402

import joblib  # noqa: E402

DEFAULT_DXF = [
    "data/sample_problem/대명동201동 계통도_최소.dxf",
    "data/sample_problem/대명동201동 계통도.dxf",
    "계통도_LH_306.dxf",
]
BRIDGE_TOLERANCES_MM = (200.0, 500.0, 1000.0, 2000.0, 5000.0, 10000.0)


def _model_path(mode):
    fname = "linkpred_rf_v2.joblib" if mode == "remote" else f"linkpred_rf_v2_{mode}.joblib"
    return _HERE / "models" / fname


def heuristic_bridges(graph, edge_len, scale, tolerances=BRIDGE_TOLERANCES_MM):
    """휴리스틱이 추가하는 복원 연결(bridge) 집합 — {끝점쌍 key: gap}.

    실제 추출(build_system_graph)과 같은 단계별 거리 bridge + force_connect(무제한).
    원본 그래프를 보존하기 위해 복사본 위에서 돈다.
    """
    g = {u: set(v) for u, v in graph.items()}
    el = dict(edge_len)
    H: dict = {}
    for tol in list(tolerances) + [float("inf")]:
        out: set = set()
        _bridge_components(g, el, max_bridge_mm=tol * scale if tol != float("inf") else tol,
                           bridge_edges_out=out)
        for k in out:
            H[k] = el.get(k)
    return H


def _dist(p, q):
    return ((p[0] - q[0]) ** 2 + (p[1] - q[1]) ** 2) ** 0.5


_MODEL_CACHE: dict = {}


def load_model(mode="allt"):
    """모델+피처 번들을 모드별 1회 로드해 캐시. 반환 (model, feats) 또는 None."""
    if mode in _MODEL_CACHE:
        return _MODEL_CACHE[mode]
    mpath = _model_path(mode)
    if not mpath.is_file():
        _MODEL_CACHE[mode] = None
        return None
    bundle = joblib.load(mpath)
    pair = (bundle["model"], bundle["features"])
    _MODEL_CACHE[mode] = pair
    return pair


def _xy(p):
    return [int(round(p[0])), int(round(p[1]))]


def serialize_result(res):
    """reconcile_entities 결과 → 프론트 오버레이용 JSON 직렬화(좌표 정수화).

    A/CONFLICT 항목 6-tuple (pr,kind,info,f,gap,d) → {a,b,prob,kind,gap_norm,h_gap,target_diff}
    C 항목 4-tuple (pr,kind,info,f)              → {a,b,prob,kind,gap_norm}
    B 항목 (key,gap) — key=(p,q) 끝점쌍           → {a,b,gap}
    """
    if res is None:
        return {"ok": True, "empty": True}

    def _ml6(row):
        pr, kind, info, f, gap, d = row
        return {"a": _xy(info["a"]), "b": _xy(info["b"]), "prob": round(float(pr), 4),
                "kind": kind, "gap_norm": round(float(f["gap_norm"]), 4),
                "h_gap": (round(float(gap), 3) if gap is not None else None),
                "target_diff": round(float(d), 3)}

    def _ml4(row):
        pr, kind, info, f = row
        return {"a": _xy(info["a"]), "b": _xy(info["b"]), "prob": round(float(pr), 4),
                "kind": kind, "gap_norm": round(float(f["gap_norm"]), 4)}

    def _b(item):
        (a, b), gap = item
        return {"a": _xy(a), "b": _xy(b),
                "gap": (round(float(gap), 3) if gap is not None else None)}

    return {
        "ok": True, "empty": False,
        "med_edge": round(float(res["med_edge"]), 4),
        "scale": round(float(res["scale"]), 6),
        "counts": {"n_seg": res["n_seg"], "n_tip": res["n_tip"], "H": res["H"],
                   "n_cand": res["n_cand"], "n_htip": res["n_htip"],
                   "A": len(res["A"]), "CONFLICT": len(res["CONF"]),
                   "B": len(res["B"]), "C": len(res["C"])},
        "A": [_ml6(r) for r in res["A"]],
        "CONFLICT": [_ml6(r) for r in res["CONF"]],
        "B": [_b(it) for it in res["B"]],
        "C": [_ml4(r) for r in res["C"]],
    }


def reconcile(dxf_path: Path, model, feats, *, ml_cut=0.45, search_factor=2.5,
              match_frac=0.75):
    """DXF 파일 경로를 받아 파싱 후 reconcile_entities 로 위임."""
    parsed = parse_dxf_for_view(dxf_path, include_hidden_layers=True)
    return reconcile_entities(parsed["entities"], model, feats, ml_cut=ml_cut,
                              search_factor=search_factor, match_frac=match_frac)


def reconcile_entities(entities, model, feats, *, ml_cut=0.45, search_factor=2.5,
                       match_frac=0.75):
    """raw 그래프에서 휴리스틱 H 와 ML 제안 M 을 끝단 단위로 대조.

    합의 판정은 "같은 끝단을 비슷한 위치로 연결하는가" — 휴리스틱 bridge 는 끝단↔본관
    중간노드를 잇기도 해 끝점쌍 정확일치는 드물다. 그래서 끝단을 단위로,
    목표점 거리 ≤ match_frac·med_edge 면 합의(A), 멀면 충돌(CONFLICT).

    등급:
      A        : 둘 다 같은 끝단을 비슷한 위치로 → 고신뢰
      CONFLICT : 둘 다 그 끝단을 잇지만 목표가 다름 → 최우선 검수
      B        : 휴리스틱만 (ML 침묵 / 끝단 아닌 component 병합)
      C        : ML 만 (휴리스틱이 그 끝단을 안 이음 — T분기 포함)
    반환 dict; tips·segs 없으면 None.
    """
    graph, edge_len, scale = _raw_graph(entities)
    segs, tips, med_edge = graph_to_segs_tips(graph, edge_len)
    if not tips or not segs:
        return None
    tip_coords = {t.coord for t in tips}

    # ── 휴리스틱 bridge → 끝단별 목표점(끝단 아닌 쪽). 끝단당 최근접 1개. ──
    H = heuristic_bridges(graph, edge_len, scale)
    h_by_tip: dict = {}     # tip 좌표 -> (dest, gap, key)
    h_engaged = set()       # 끝단과 엮인 bridge key
    h_no_tip = []           # 양끝 모두 끝단 아님(순수 component 병합)
    for key, gap in H.items():
        a, b = key
        ends = [(a, b), (b, a)]
        touched = False
        for tip_c, dest in ends:
            if tip_c in tip_coords:
                touched = True
                if tip_c not in h_by_tip or (gap or 0) < (h_by_tip[tip_c][1] or 0):
                    h_by_tip[tip_c] = (dest, gap, key)
        if not touched:
            h_no_tip.append((key, gap))

    # ── ML 후보 점수화 → 끝단별 top-1 ranker ──
    rows = lpd.candidate_rows(segs, tips, med_edge, search_factor=search_factor)
    by_tip: dict = {}       # tip 좌표 -> (prob, kind, info, feat)
    if rows:
        for (f, _l, info), pr in zip(rows, prob_iter(model, feats, rows)):
            key = info["a"]
            if key not in by_tip or pr > by_tip[key][0]:
                by_tip[key] = (pr, info["kind"], info, f)
    m_by_tip = {k: v for k, v in by_tip.items() if v[0] >= ml_cut}

    # ── 끝단 단위 대조 ──
    tol = match_frac * med_edge
    A, CONF, C = [], [], []
    for tip_c, (pr, kind, info, f) in m_by_tip.items():
        if tip_c in h_by_tip:
            dest, gap, key = h_by_tip[tip_c]
            h_engaged.add(key)
            d = _dist(info["b"], dest)
            row = (pr, kind, info, f, gap, d)
            (A if d <= tol else CONF).append(row)
        else:
            C.append((pr, kind, info, f))
    # B = 끝단과 엮였으나 ML 침묵한 bridge + 순수 component 병합
    B = [(k, g) for k, g in H.items() if k not in h_engaged]

    A.sort(key=lambda t: -t[0])
    CONF.sort(key=lambda t: -t[0])
    C.sort(key=lambda t: -t[0])
    B.sort(key=lambda t: (t[1] if t[1] is not None else 0.0))
    return {"med_edge": med_edge, "scale": scale, "n_seg": len(segs),
            "n_tip": len(tips), "H": len(H), "n_cand": len(rows),
            "n_htip": len(h_by_tip), "A": A, "CONF": CONF, "B": B, "C": C}


def prob_iter(model, feats, rows):
    X = np.array([[f[k] for k in feats] for f, _l, _i in rows], float)
    return model.predict_proba(X)[:, 1]


def _kindcount(items):
    tt = sum(1 for it in items if it[2]["kind"] == "tt")
    te = sum(1 for it in items if it[2]["kind"] == "te")
    return tt, te


def report(model, feats, dxf: Path, *, ml_cut=0.45):
    res = reconcile(dxf, model, feats, ml_cut=ml_cut)
    print("\n" + "=" * 96)
    print(f"통합 대조 — {dxf.name}")
    print("=" * 96)
    if res is None:
        print("  후보 없음 (tips/segs 0).")
        return
    print(f"med_edge={res['med_edge']:.3f} · scale={res['scale']:.5f} · "
          f"본관 edge {res['n_seg']} · 끝단 {res['n_tip']} · 후보쌍 {res['n_cand']}")
    A, CONF, B, C = res["A"], res["CONF"], res["B"], res["C"]
    a_tt, a_te = _kindcount(A)
    f_tt, f_te = _kindcount(CONF)
    c_tt, c_te = _kindcount(C)
    print(f"휴리스틱 bridge {res['H']}개(끝단 엮임 {res['n_htip']}) · "
          f"ML 제안(끝단 top-1 ≥{ml_cut}) {len(A) + len(CONF) + len(C)}개")
    print("-" * 96)
    print(f"  A 합의(같은 끝단·같은 위치) : {len(A):>3}  (tt {a_tt} · te {a_te})  → 고신뢰·실선 후보")
    print(f"  CONFLICT(같은 끝단·다른 위치): {len(CONF):>3}  (tt {f_tt} · te {f_te})  → 최우선 검수")
    print(f"  B 휴리스틱-단독            : {len(B):>3}                  → 점선 검수")
    print(f"  C ML-단독                  : {len(C):>3}  (tt {c_tt} · te {c_te})  → 점선 검수(T분기 포함)")

    if A:
        print("\n  [A] 합의 표본 (prob · kind · gap_norm · 목표거리):")
        for pr, kind, info, f, gap, d in A[:6]:
            print(f"    prob={pr:.2f}  {kind}  gap_norm={f['gap_norm']:.2f}  목표차={d:.2f}")
    if CONF:
        print("\n  [CONFLICT] 둘 다 잇지만 목표가 다른 끝단 (prob · kind · 목표거리):")
        for pr, kind, info, f, gap, d in CONF[:6]:
            print(f"    prob={pr:.2f}  {kind}  gap_norm={f['gap_norm']:.2f}  "
                  f"휴리거리={(gap or 0):.2f}  목표차={d:.2f}")
    if C:
        print("\n  [C] ML-단독 표본 (휴리스틱이 안 이은 끝단):")
        for pr, kind, info, f in C[:8]:
            extra = (f"entry={f['te_entry_angle']:.0f}° interior={f['te_proj_interior']:.2f}"
                     if kind == "te"
                     else f"cos_dirs={f['tt_cos_dirs']:+.2f} align={f['tip_align']:.0f}°")
            print(f"    prob={pr:.2f}  {kind}  gap_norm={f['gap_norm']:.2f}  {extra}")
    if B:
        gaps = [g for _k, g in B if g is not None]
        if gaps:
            gaps = np.array(gaps)
            print(f"\n  [B] 휴리스틱-단독 gap 분포: p50={np.percentile(gaps,50):.2f} "
                  f"p90={np.percentile(gaps,90):.2f} max={gaps.max():.2f} "
                  f"(ML 이 동의 안 한 거리 연결 — 검수 필요)")


def main(argv):
    args = argv[1:]
    mode = "allt"
    if args and args[0] in ("remote", "all", "allt"):
        mode, args = args[0], args[1:]
    targets = args if args else DEFAULT_DXF
    mpath = _model_path(mode)
    if not mpath.is_file():
        print(f"모델 없음 — 먼저 linkpred_train_v2.py {mode} 실행: {mpath}")
        return 1
    bundle = joblib.load(mpath)
    model, feats = bundle["model"], bundle["features"]
    print("=" * 96)
    print(f"연결복원 통합 (휴리스틱 × ML) [corpus={mode}] — 기하-only ({mpath.name})")
    print("=" * 96)
    for t in targets:
        p = Path(t)
        if not p.is_absolute():
            p = _ROOT / t
        if not p.is_file():
            print(f"  (없음) {t}")
            continue
        report(model, feats, p)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
