# -*- coding: utf-8 -*-
"""link-prediction 연결복원 — 학습 + 망단위 교차검증.

정답 SDF(스프링클러 REMOTE 43건)를 합성손상해 만든 후보쌍으로 트리 분류기를
학습한다. 핵심 검증: **망(network) 단위 GroupKFold** — 학습에 안 쓴 도면에서도
'끊긴 끝단쌍 연결' 판정이 되는가(전이성). 거리만 쓰는 baseline 과 비교해
ML 의 부가가치(직선성·방향·구경 피처)를 정량화한다.

실행:
    python calibration/linkpred_train.py            # 학습+CV 리포트, 모델 저장
"""
from __future__ import annotations

import sys
from pathlib import Path

import numpy as np

_ROOT = Path(__file__).resolve().parent.parent
_HERE = Path(__file__).resolve().parent
for _p in (str(_ROOT), str(_HERE)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

from kfp_sdf_converter import parse_sdf  # noqa: E402
import linkpred_data as lpd  # noqa: E402
import validate_sdf as vs  # noqa: E402

from sklearn.ensemble import RandomForestClassifier  # noqa: E402
from sklearn.model_selection import GroupKFold  # noqa: E402
from sklearn.metrics import average_precision_score, precision_recall_fscore_support  # noqa: E402

SEEDS = (1, 2, 3)  # 도면당 손상 패턴 3종 증강
MODEL_DIR = _HERE / "models"


def build_matrix(files, seeds=SEEDS):
    X, y, groups = [], [], []
    per_net = []
    for gi, f in enumerate(files):
        net = parse_sdf(f)
        n0 = len(X)
        for s in seeds:
            for feat, lab in lpd.dataset_from_net(net, seed=s):
                X.append([feat[k] for k in lpd.FEATURE_NAMES])
                y.append(lab)
                groups.append(gi)
        per_net.append((Path(f).name, len(X) - n0))
    return np.array(X, float), np.array(y, int), np.array(groups, int), per_net


def _metrics(y_true, score, thresh):
    pred = (score >= thresh).astype(int)
    p, r, f1, _ = precision_recall_fscore_support(
        y_true, pred, average="binary", zero_division=0)
    return p, r, f1


def main():
    files = vs._remote_answer_files()
    print("=" * 96)
    print(f"link-prediction 학습 — 정답 SDF(스프링클러 REMOTE) {len(files)}건 · 손상시드 {SEEDS}")
    print("=" * 96)

    X, y, groups, per_net = build_matrix(files)
    print(f"\n후보쌍 총 {len(y)} (양성 {int(y.sum())} · 음성 {int((y == 0).sum())}) "
          f"· 망 {len(set(groups))}개 · 피처 {X.shape[1]}종")

    # ── 망단위 GroupKFold ──
    n_splits = min(5, len(set(groups)))
    gkf = GroupKFold(n_splits=n_splits)
    ml_ap, base_ap = [], []
    ml_prf, base_prf = [], []
    for fold, (tr, te) in enumerate(gkf.split(X, y, groups)):
        clf = RandomForestClassifier(
            n_estimators=300, max_depth=None, min_samples_leaf=2,
            class_weight="balanced", n_jobs=-1, random_state=0)
        clf.fit(X[tr], y[tr])
        prob = clf.predict_proba(X[te])[:, 1]
        # baseline: 거리만 (gap_norm 작을수록 연결) → score = -gap_norm
        gap_norm = X[te][:, lpd.FEATURE_NAMES.index("gap_norm")]
        base_score = -gap_norm

        if y[te].sum() == 0:
            continue
        ml_ap.append(average_precision_score(y[te], prob))
        base_ap.append(average_precision_score(y[te], base_score))
        ml_prf.append(_metrics(y[te], prob, 0.5))
        # baseline 임계: 학습 양성 gap_norm 중앙값
        thr = np.median(X[tr][y[tr] == 1][:, lpd.FEATURE_NAMES.index("gap_norm")])
        base_prf.append(_metrics(y[te], -gap_norm, -thr))

    def _mean(a, i=None):
        arr = np.array(a)
        return arr.mean(axis=0)[i] if i is not None else arr.mean()

    print("\n" + "-" * 96)
    print("망단위 GroupKFold (학습에 안 본 도면으로 평가) — 전이성")
    print("-" * 96)
    print(f"{'':16}{'PR-AUC':>9}{'precision':>11}{'recall':>9}{'F1':>8}")
    print(f"{'ML(랜덤포레스트)':16}{_mean(ml_ap):>9.3f}"
          f"{_mean(ml_prf,0):>11.3f}{_mean(ml_prf,1):>9.3f}{_mean(ml_prf,2):>8.3f}")
    print(f"{'baseline(거리만)':16}{_mean(base_ap):>9.3f}"
          f"{_mean(base_prf,0):>11.3f}{_mean(base_prf,1):>9.3f}{_mean(base_prf,2):>8.3f}")
    lift = (_mean(ml_ap) - _mean(base_ap))
    print(f"\nPR-AUC 향상 (ML − 거리): {lift:+.3f}")

    # ── 피처 중요도 (전체 학습) ──
    final = RandomForestClassifier(
        n_estimators=400, min_samples_leaf=2, class_weight="balanced",
        n_jobs=-1, random_state=0)
    final.fit(X, y)
    imp = sorted(zip(lpd.FEATURE_NAMES, final.feature_importances_),
                 key=lambda t: -t[1])
    print("\n피처 중요도(전체학습):")
    for name, v in imp:
        print(f"  {name:22} {v:.3f}")

    # ── 모델 저장 ──
    MODEL_DIR.mkdir(exist_ok=True)
    import joblib
    out = MODEL_DIR / "linkpred_rf.joblib"
    joblib.dump({"model": final, "features": lpd.FEATURE_NAMES}, out)
    print(f"\n모델 저장: {out}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
