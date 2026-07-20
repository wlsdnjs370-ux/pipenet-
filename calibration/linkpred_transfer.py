# -*- coding: utf-8 -*-
"""link-prediction 연결복원 — 실제 파탄 DXF 전이 점검.

정답 SDF 합성손상으로 학습한 분류기가 **실제로 끊긴 도면**(LH306 계통도 등)의
끝단 연결 판정에 전이되는지 본다. DXF 에는 구경(nominal) 정보가 없으므로
(배관 재질/구경은 DXF 에 없음 — 설계협의 단계 변동) 학습도 **기하 피처만**으로
하여 학습/추론 도메인을 일치시킨다.

실제 DXF 엔 ground truth 가 없어 정량 채점은 불가 → 정성 점검:
  · 끊긴 끝단(dangling) 수, 후보쌍 수
  · 모델이 prob≥0.5/0.8 로 제안한 연결 수
  · 거리만 쓰는 baseline(각 끝단 최근접 연결) 대비 선택성
  · 제안 연결 표본 (gap_norm·직선성·prob)

실행:
    python calibration/linkpred_transfer.py [dxf ...]
"""
from __future__ import annotations

import sys
from collections import defaultdict
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
from clean_candidate_survey import _raw_graph  # noqa: E402
from remote30_prototype import parse_dxf_for_view, _drawing_scale_ratio  # noqa: E402

from sklearn.ensemble import RandomForestClassifier  # noqa: E402
from sklearn.model_selection import GroupKFold  # noqa: E402
from sklearn.metrics import average_precision_score  # noqa: E402

GEOM_FEATURES = [f for f in lpd.FEATURE_NAMES if not f.startswith("nominal")]

DEFAULT_DXF = [
    "data/sample_problem/대명동201동 계통도_최소.dxf",
    "data/sample_problem/대명동201동 계통도.dxf",
    "samples/dxf/계통도_LH_306.dxf",
]


def _build_geom_matrix(files, seeds=(1, 2, 3)):
    idx = [lpd.FEATURE_NAMES.index(f) for f in GEOM_FEATURES]
    X, y, groups = [], [], []
    for gi, f in enumerate(files):
        net = parse_sdf(f)
        for s in seeds:
            for feat, lab in lpd.dataset_from_net(net, seed=s):
                X.append([feat[k] for k in GEOM_FEATURES])
                y.append(lab)
                groups.append(gi)
    return np.array(X, float), np.array(y, int), np.array(groups, int)


def train_geom_model(verbose=True):
    files = vs._remote_answer_files()
    X, y, g = _build_geom_matrix(files)
    if verbose:
        # 망단위 CV 로 기하-only 전이성 확인
        gkf = GroupKFold(n_splits=5)
        aps = []
        for tr, te in gkf.split(X, y, g):
            if y[te].sum() == 0:
                continue
            m = RandomForestClassifier(n_estimators=300, min_samples_leaf=2,
                                       class_weight="balanced", n_jobs=-1, random_state=0)
            m.fit(X[tr], y[tr])
            aps.append(average_precision_score(y[te], m.predict_proba(X[te])[:, 1]))
        print(f"기하-only 망단위 CV PR-AUC: {np.mean(aps):.3f} "
              f"(피처 {GEOM_FEATURES})")
    model = RandomForestClassifier(n_estimators=400, min_samples_leaf=2,
                                   class_weight="balanced", n_jobs=-1, random_state=0)
    model.fit(X, y)
    return model


# ── DXF 끊긴 끝단 → 후보쌍 + 기하 피처 ───────────────────────────────────────
def dxf_candidates(dxf: Path, search_factor=2.5):
    parsed = parse_dxf_for_view(dxf, include_hidden_layers=True)
    graph, edge_len, scale = _raw_graph(parsed["entities"])
    deg = {u: len(vs_) for u, vs_ in graph.items()}
    dangling = [u for u, d in deg.items() if d == 1]
    if not edge_len:
        return [], 0, scale, len(dangling)
    med_edge = sorted(edge_len.values())[len(edge_len) // 2]
    R = search_factor * med_edge

    # dangling u 의 진행방향 = u - (유일 이웃)
    def out_dir(u):
        v = next(iter(graph[u]))
        return lpd._unit(u[0] - v[0], u[1] - v[1]), v

    cell = max(R, 1e-9)
    grid = defaultdict(list)
    for i, u in enumerate(dangling):
        grid[(int(u[0] // cell), int(u[1] // cell))].append(i)

    rows = []
    seen = set()
    R_sq = R * R
    for i, u in enumerate(dangling):
        du, _ = out_dir(u)
        kx, ky = int(u[0] // cell), int(u[1] // cell)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                for j in grid.get((kx + dx, ky + dy), ()):
                    if j <= i:
                        continue
                    key = (i, j)
                    if key in seen:
                        continue
                    w = dangling[j]
                    gap2 = (w[0] - u[0]) ** 2 + (w[1] - u[1]) ** 2
                    if gap2 == 0 or gap2 > R_sq:
                        continue
                    seen.add(key)
                    dw, _ = out_dir(w)
                    feat = _geom_feat(u, du, w, dw, med_edge)
                    rows.append((feat, u, w))
    return rows, med_edge, scale, len(dangling)


def _geom_feat(u, du, w, dw, med_edge):
    import math
    gap = math.hypot(w[0] - u[0], w[1] - u[1])
    ab = lpd._unit(w[0] - u[0], w[1] - u[1])
    ba = (-ab[0], -ab[1])
    align_a = lpd._angle_deg(du, ab)
    align_b = lpd._angle_deg(dw, ba)
    return {
        "gap": gap,
        "gap_norm": gap / med_edge if med_edge > 0 else gap,
        "align_A": align_a,
        "align_B": align_b,
        "align_max": max(align_a, align_b),
        "align_min": min(align_a, align_b),
        "cos_dirs": du[0] * dw[0] + du[1] * dw[1],
    }


def report_dxf(model, dxf: Path):
    rows, med_edge, scale, n_dang = dxf_candidates(dxf)
    print("\n" + "=" * 96)
    print(f"전이 점검 — {dxf.name}")
    print("=" * 96)
    print(f"scale_ratio={scale:.5f} · med_edge={med_edge:.3f} · dangling 끝단 {n_dang} · "
          f"후보쌍 {len(rows)}")
    if not rows:
        print("  후보 없음.")
        return
    X = np.array([[r[0][k] for k in GEOM_FEATURES] for r in rows], float)
    prob = model.predict_proba(X)[:, 1]

    for thr in (0.5, 0.8):
        sel = prob >= thr
        print(f"  prob≥{thr}: 제안 연결 {int(sel.sum())}쌍")

    # baseline: 각 dangling 의 최근접 1쌍을 연결 (거리만) — bridge 식
    gaps = X[:, GEOM_FEATURES.index("gap")]
    # 끝단별 최근접 후보를 baseline 연결로 카운트 (중복 제거)
    nn = {}
    for k, (_f, u, w) in enumerate(rows):
        for pt in (u, w):
            if pt not in nn or gaps[k] < nn[pt][1]:
                nn[pt] = (k, gaps[k])
    base_pairs = {nn[pt][0] for pt in nn}
    print(f"  baseline(최근접 연결): {len(base_pairs)}쌍")

    # 제안 표본 — 고신뢰 상위 6
    order = np.argsort(-prob)
    print("  고신뢰 제안 표본 (prob · gap_norm · align_min°):")
    for k in order[:6]:
        f = rows[k][0]
        print(f"    prob={prob[k]:.2f}  gap_norm={f['gap_norm']:.2f}  "
              f"align_min={f['align_min']:.0f}  cos_dirs={f['cos_dirs']:+.2f}")
    # 거리는 가깝지만(=baseline 이 이을) 모델이 거부한 쌍 — 선택성 증거
    rej = [k for k in base_pairs if prob[k] < 0.5]
    print(f"  baseline 이 잇지만 모델이 거부(prob<0.5): {len(rej)}/{len(base_pairs)}쌍 "
          f"(거리는 가깝지만 방향 불일치)")


def main(argv):
    targets = argv[1:] if len(argv) > 1 else DEFAULT_DXF
    print("=" * 96)
    print("link-prediction 전이 점검 — 기하-only 모델 (DXF 엔 구경정보 없음)")
    print("=" * 96)
    model = train_geom_model(verbose=True)
    for t in targets:
        p = Path(t)
        if not p.is_absolute():
            p = _ROOT / t
        if not p.is_file():
            print(f"  (없음) {t}")
            continue
        report_dxf(model, p)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
