# -*- coding: utf-8 -*-
"""link-prediction 연결복원 v2 — 실제 파탄 DXF 전이 재점검.

v1 은 끝점↔끝점만 봐서 실제 DXF 제안 0건이었다. v2 모델은 tip-edge(T분기)를
1급으로 학습했다. 실제 도면(LH306 계통도·대명동)에서:
  · raw 그래프(끝점 snap only, 거리 브리지 없음)의 모든 edge 를 본관 후보 세그먼트로,
    모든 dangling 끝단을 tip 으로 삼고
  · tip→tip ∪ tip→edge 후보를 만들어 v2 모델로 점수화
  · prob≥0.5/0.8 제안 수를 유형별로 집계 (v1 과 직접 비교)

DXF 엔 ground truth 가 없어 정성 점검(제안 수·표본)만. 구경 없으므로 기하-only.

실행:
    python calibration/linkpred_transfer_v2.py [dxf ...]
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

import linkpred_data_v2 as lpd  # noqa: E402
from clean_candidate_survey import _raw_graph  # noqa: E402
from remote30_prototype import parse_dxf_for_view  # noqa: E402

import joblib  # noqa: E402

DEFAULT_DXF = [
    "data/sample_problem/대명동201동 계통도_최소.dxf",
    "data/sample_problem/대명동201동 계통도.dxf",
    "samples/dxf/계통도_LH_306.dxf",
]


def _model_path(mode):
    fname = "linkpred_rf_v2.joblib" if mode == "remote" else f"linkpred_rf_v2_{mode}.joblib"
    return _HERE / "models" / fname


def graph_to_segs_tips(graph, edge_len):
    """raw 그래프 → (_Seg 리스트, _Tip 리스트, med_edge).

    edge 하나당 _Seg(pts=[a,b]).  dangling 노드 하나당 _Tip(seg_id=자기 edge).
    통합엔진과 그래프를 공유하기 위해 그래프를 인자로 받는다.
    """
    if not edge_len:
        return [], [], 0.0
    med_edge = sorted(edge_len.values())[len(edge_len) // 2]

    segs, edge_seg = [], {}
    for sid, (a, b) in enumerate(edge_len.keys()):
        segs.append(lpd._Seg(id=sid, pts=[a, b], absorbed=set()))
        edge_seg[frozenset((a, b))] = sid

    tips = []
    deg = {u: len(v) for u, v in graph.items()}
    for tid, u in enumerate(d for d, k in deg.items() if k == 1):
        nb = next(iter(graph[u]))
        own = edge_seg.get(frozenset((u, nb)), -1)
        tips.append(lpd._Tip(id=tid, coord=u,
                             out_dir=lpd._unit(u[0] - nb[0], u[1] - nb[1]),
                             node_id="", seg_id=own))
    return segs, tips, med_edge


def dxf_to_segs_tips(dxf: Path):
    """DXF 파싱 → raw 그래프 → (_Seg, _Tip, med_edge)."""
    parsed = parse_dxf_for_view(dxf, include_hidden_layers=True)
    graph, edge_len, scale = _raw_graph(parsed["entities"])
    return graph_to_segs_tips(graph, edge_len)


def report(model, feats, dxf: Path):
    segs, tips, med_edge = dxf_to_segs_tips(dxf)
    print("\n" + "=" * 96)
    print(f"v2 전이 점검 — {dxf.name}")
    print("=" * 96)
    print(f"med_edge={med_edge:.3f} · 본관 edge {len(segs)} · dangling 끝단 {len(tips)}")
    if not tips or not segs:
        print("  후보 없음.")
        return
    rows = lpd.candidate_rows(segs, tips, med_edge, search_factor=2.5)
    n_tt = sum(1 for _f, _l, i in rows if i["kind"] == "tt")
    n_te = sum(1 for _f, _l, i in rows if i["kind"] == "te")
    print(f"후보쌍 {len(rows)} — tip-tip {n_tt} · tip-edge {n_te}")
    if not rows:
        return
    X = np.array([[f[k] for k in feats] for f, _l, _i in rows], float)
    prob = model.predict_proba(X)[:, 1]
    kinds = [i["kind"] for _f, _l, i in rows]

    # 절대 임계 (참고용 — 도메인 시프트로 보정 안 됨)
    for thr in (0.5, 0.8):
        sel = prob >= thr
        print(f"  절대 prob≥{thr}: {int(sel.sum())}쌍")

    # 배포형 사용 = 끝단별 최고점 후보(top-1 ranker). 각 tip 의 최선 연결 1개.
    by_tip = {}   # tip좌표 -> (prob, kind, feat)
    for (f, _l, info), pr in zip(rows, prob):
        key = info["a"]
        if key not in by_tip or pr > by_tip[key][0]:
            by_tip[key] = (pr, info["kind"], f)
    best = sorted(by_tip.values(), key=lambda t: -t[0])
    bp = np.array([b[0] for b in best])
    print(f"  끝단별 top-1 ranker — {len(best)}개 끝단, 최선후보 prob "
          f"p50={np.percentile(bp,50):.2f} p75={np.percentile(bp,75):.2f} "
          f"p90={np.percentile(bp,90):.2f} max={bp.max():.2f}")
    for rel in (0.40, 0.45):
        sel = [b for b in best if b[0] >= rel]
        tt = sum(1 for b in sel if b[1] == "tt")
        te = sum(1 for b in sel if b[1] == "te")
        print(f"    상대컷 prob≥{rel}: 끝단 {len(sel)}개 연결제안 "
              f"(tip-tip {tt} · tip-edge {te})")

    order = np.argsort(-prob)
    print("  최고점 후보 표본 (prob · kind · gap_norm):")
    seen_b = set()
    shown = 0
    for k in order:
        f, _l, info = rows[k]
        if info["a"] in seen_b:      # 끝단당 1개만 표본
            continue
        seen_b.add(info["a"])
        extra = (f"entry={f['te_entry_angle']:.0f}° interior={f['te_proj_interior']:.2f}"
                 if info["kind"] == "te"
                 else f"cos_dirs={f['tt_cos_dirs']:+.2f} align={f['tip_align']:.0f}°")
        print(f"    prob={prob[k]:.2f}  {info['kind']}  "
              f"gap_norm={f['gap_norm']:.2f}  {extra}")
        shown += 1
        if shown >= 8:
            break


def main(argv):
    args = argv[1:]
    mode = "remote"
    if args and args[0] in ("remote", "all", "allt"):
        mode, args = args[0], args[1:]
    targets = args if args else DEFAULT_DXF
    model_path = _model_path(mode)
    if not model_path.is_file():
        print(f"모델 없음 — 먼저 linkpred_train_v2.py {mode} 실행: {model_path}")
        return 1
    bundle = joblib.load(model_path)
    model, feats = bundle["model"], bundle["features"]
    print("=" * 96)
    print(f"link-prediction v2 전이 점검 [corpus={mode}] — 기하-only ({model_path.name})")
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
