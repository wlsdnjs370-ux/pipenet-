# -*- coding: utf-8 -*-
"""link-prediction 연결복원 v2 — 학습 + 망단위 CV (유형별 분리 리포트).

정답 SDF(스프링클러 REMOTE)를 v2 합성손상(un-snap break + T-tap)으로 끊어
tip-tip ∪ tip-edge 통합 후보를 만들고 단일 RandomForest(`is_edge` 플래그 포함)를
학습한다. 망(network) 단위 GroupKFold 로 미학습 도면 전이성을 보되, **tip-tip /
tip-edge 유형별로 PR-AUC 를 분리 리포트** 한다(v1 은 tip-edge 가 아예 없었음).

피처는 기하-only (DXF 엔 구경 없음).

실행:
    python calibration/linkpred_train_v2.py
"""
from __future__ import annotations

import sys
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

from kfp_sdf_converter import parse_sdf  # noqa: E402
import linkpred_data_v2 as lpd  # noqa: E402

from sklearn.ensemble import RandomForestClassifier  # noqa: E402
from sklearn.model_selection import GroupKFold  # noqa: E402
from sklearn.metrics import average_precision_score, precision_recall_fscore_support  # noqa: E402

SEEDS = (1, 2, 3)
MODEL_DIR = _HERE / "models"
EDGE_IDX = lpd.FEATURE_NAMES.index("is_edge")
GAPN_IDX = lpd.FEATURE_NAMES.index("gap_norm")


def structural_fingerprint(net):
    """좌표 무관 구조 지문 — (파이프수, 정렬된 차수열).

    청주아파트 101~108동×다층처럼 near-duplicate 설계는 노드좌표만 다르고 위상은
    동일하다. 이 지문이 같으면 같은 설계로 보고 GroupKFold 한 fold 에 묶어
    train/test 누출(=CV 부풀림)을 막는다. 334망 → 92 클러스터(3.6x 중복).
    """
    deg = {}
    for p in net.pipes.values():
        deg[p.start] = deg.get(p.start, 0) + 1
        deg[p.end] = deg.get(p.end, 0) + 1
    return (len(net.pipes), tuple(sorted(deg.values())))


def build_matrix(files, seeds=SEEDS, **dskw):
    """반환: X, y, file_groups(파일별), clust_groups(구조지문별), used.

    file_groups 는 망(파일)마다 고유 id (per-file CV — near-dup 누출 포함).
    clust_groups 는 구조지문이 같은 망을 한 그룹으로 묶음 (honest CV).
    """
    X, y, file_groups, clust_groups = [], [], [], []
    used = 0
    fp_id = {}
    for f in files:
        try:
            net = parse_sdf(f)
        except Exception:
            continue
        if len(net.pipes) < 5 or len(net.nodes) < 5:
            continue
        fp = structural_fingerprint(net)
        cid = fp_id.setdefault(fp, len(fp_id))
        rows0 = len(X)
        for s in seeds:
            for feat, lab in lpd.dataset_from_net(net, seed=s, **dskw):
                X.append([feat[k] for k in lpd.FEATURE_NAMES])
                y.append(lab)
                file_groups.append(used)
                clust_groups.append(cid)
        if len(X) > rows0:
            used += 1
    return (np.array(X, float), np.array(y, int),
            np.array(file_groups, int), np.array(clust_groups, int), used)


def _ap(y_true, score):
    return average_precision_score(y_true, score) if y_true.sum() else float("nan")


def _prf(y_true, score, thr=0.5):
    pred = (score >= thr).astype(int)
    p, r, f1, _ = precision_recall_fscore_support(
        y_true, pred, average="binary", zero_division=0)
    return p, r, f1


# 모드 → (코퍼스, 손상 kwargs). allt = 전 코퍼스 + T-tap 합성률 상향(tip-edge 희석 보정).
MODE_CFG = {
    "remote": ("remote", {}),
    "all":    ("all", {}),
    "allt":   ("all", {"ttap_prob": 0.9}),
}


def _m(a):
    a = [v for v in a if v == v]   # drop nan
    return float(np.mean(a)) if a else float("nan")


def run_cv(X, y, g, is_edge):
    """주어진 그룹배열로 GroupKFold — (전체,tip-tip,tip-edge,baseline,prf) 평균."""
    gkf = GroupKFold(n_splits=min(5, len(set(g))))
    ap_all, ap_tt, ap_te, ap_base, prf_all = [], [], [], [], []
    for tr, te in gkf.split(X, y, g):
        clf = RandomForestClassifier(
            n_estimators=300, min_samples_leaf=2, class_weight="balanced",
            n_jobs=-1, random_state=0)
        clf.fit(X[tr], y[tr])
        prob = clf.predict_proba(X[te])[:, 1]
        ye, ie = y[te], is_edge[te]
        if ye.sum() == 0:
            continue
        ap_all.append(_ap(ye, prob))
        prf_all.append(_prf(ye, prob))
        if ie.any() and ye[ie].sum():
            ap_te.append(_ap(ye[ie], prob[ie]))
        if (~ie).any() and ye[~ie].sum():
            ap_tt.append(_ap(ye[~ie], prob[~ie]))
        ap_base.append(_ap(ye, -X[te][:, GAPN_IDX]))
    return ap_all, ap_tt, ap_te, ap_base, prf_all


def _report_cv(title, res):
    ap_all, ap_tt, ap_te, ap_base, prf_all = res
    pa = np.mean(prf_all, axis=0)
    print("\n" + "-" * 96)
    print(title)
    print("-" * 96)
    print(f"  PR-AUC 전체      : {_m(ap_all):.3f}   (precision {pa[0]:.3f} · "
          f"recall {pa[1]:.3f} · F1 {pa[2]:.3f})")
    print(f"  PR-AUC tip-tip   : {_m(ap_tt):.3f}")
    print(f"  PR-AUC tip-edge  : {_m(ap_te):.3f}   ← v2 신규(끝점↔edge)")
    print(f"  PR-AUC 거리baseline: {_m(ap_base):.3f}")
    print(f"  향상 (ML − 거리) : {_m(ap_all) - _m(ap_base):+.3f}")


def main(argv=None):
    argv = argv or sys.argv
    mode = argv[1] if len(argv) > 1 else "remote"
    corpus, dskw = MODE_CFG.get(mode, ("remote", {}))
    files = lpd.corpus_files(corpus)
    print("=" * 96)
    print(f"link-prediction v2 학습 [mode={mode} corpus={corpus} {dskw}] — "
          f"답안 SDF {len(files)}건 후보 · 손상시드 {SEEDS} "
          f"· 피처 {len(lpd.FEATURE_NAMES)}종(기하-only)")
    print("=" * 96)

    X, y, fg, cg, used = build_matrix(files, **dskw)
    n_file, n_clust = len(set(fg)), len(set(cg))
    print(f"실사용 망 {used}개 (규모/파싱 필터 통과) — 구조지문 {n_clust}개 클러스터 "
          f"(중복도 {n_file / max(n_clust,1):.1f}x)")
    is_edge = X[:, EDGE_IDX] == 1.0
    print(f"\n후보쌍 총 {len(y)} (양성 {int(y.sum())} · 음성 {int((y == 0).sum())}) "
          f"· 망 {n_file}개")
    print(f"  tip-tip : {int((~is_edge).sum())} (양성 {int(y[~is_edge].sum())})")
    print(f"  tip-edge: {int(is_edge.sum())} (양성 {int(y[is_edge].sum())})")

    # ── per-file CV (near-dup 누출 포함 = 부풀림) vs 구조클러스터 CV (정직) ──
    _report_cv("per-file GroupKFold (near-dup 누출 포함 — 낙관 편향)",
               run_cv(X, y, fg, is_edge))
    if n_clust < n_file:
        _report_cv(f"구조클러스터 GroupKFold ({n_clust}클러스터 de-dup — 정직한 전이성)",
                   run_cv(X, y, cg, is_edge))

    # ── 피처 중요도 + 모델 저장 ──
    final = RandomForestClassifier(
        n_estimators=400, min_samples_leaf=2, class_weight="balanced",
        n_jobs=-1, random_state=0)
    final.fit(X, y)
    imp = sorted(zip(lpd.FEATURE_NAMES, final.feature_importances_),
                 key=lambda t: -t[1])
    print("\n피처 중요도(전체학습):")
    for name, v in imp:
        print(f"  {name:20} {v:.3f}")

    MODEL_DIR.mkdir(exist_ok=True)
    import joblib
    fname = "linkpred_rf_v2.joblib" if mode == "remote" else f"linkpred_rf_v2_{mode}.joblib"
    out = MODEL_DIR / fname
    joblib.dump({"model": final, "features": lpd.FEATURE_NAMES}, out)
    print(f"\n모델 저장: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
