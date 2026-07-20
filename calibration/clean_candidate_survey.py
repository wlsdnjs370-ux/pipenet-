# -*- coding: utf-8 -*-
"""link-prediction 학습용 '깨끗한 도면' 후보 채점.

자가지도(self-supervised) 학습은 '거의 완전히 연결된' 정답망을 합성손상시켜
학습쌍을 만든다. 따라서 학습 소스로 쓸 도면은 추출 그래프가 **이미 잘 이어진**
것이어야 한다(=정답 ground truth 신뢰 가능). 이 스크립트는 후보 도면들을 실제
추출 파이프라인(build_system_graph, bridge 전 raw 연결성)으로 돌려 다음으로 점수화:

  · max_comp%   : 최대 component 가 전체 노드의 몇 %  (높을수록 깨끗)
  · n_comp      : bridge 전 component 수             (낮을수록 깨끗)
  · dangling%   : 차수1 끝단 비율                    (낮을수록 깨끗)
  · scale_ratio : 실좌표(=1.0) vs 용지스키매틱(<1)

판정: max_comp ≥ 80% 이고 dangling ≤ 35% 이면 GOOD(정답 신뢰), 아니면 PARTIAL/POOR.

실행:
    python calibration/clean_candidate_survey.py [dxf경로 ...]
    (인자 없으면 기본 후보 목록)
"""
from __future__ import annotations

import sys
from pathlib import Path

_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from remote30_prototype import (  # noqa: E402
    parse_dxf_for_view,
    _categorize_layer,
    _drawing_scale_ratio,
    _build_graph,
    _connected_components,
    _NodeIndex,
    SNAP_TOL_MM,
    MIN_PIPE_EDGE_MM,
)


def _raw_graph(ents):
    """bridge·force 없이 끝점 스냅만으로 빌드한 raw 그래프 (정답 신뢰성 측정용).

    build_system_graph 은 bridge 를 적용한 그래프를 돌려주므로 '깨끗하게 작도됐는가'
    가 아니라 '거리로 이어붙일 수 있는가'를 재게 된다. 학습 정답은 작도자가 실제로
    끝점을 스냅해 둔 도면이어야 신뢰 가능 → bridge 전 raw 로 잰다.
    """
    cats = {ly: _categorize_layer(ly) for ly in {e.get("l", "") for e in ents}}
    pipe_ents = [e for e in ents
                 if e.get("t") in ("L", "PL") and cats.get(e.get("l", ""), "OTHER") == "PIPE"]
    scale = _drawing_scale_ratio(pipe_ents)
    eps = SNAP_TOL_MM * scale
    idx = _NodeIndex(eps if scale < 1.0 else SNAP_TOL_MM)
    graph, edge_len = _build_graph(pipe_ents, node_index=idx,
                                   min_edge_mm=MIN_PIPE_EDGE_MM * scale)
    return graph, edge_len, scale

DEFAULT_CANDIDATES = [
    # 대명동 201동 — 사용자 1차 작업 도면 (실좌표 평면도/계통도)
    "samples/dxf/대명동201동 단위세대_layer정리.dxf",
    "samples/dxf/대명동201동 단위세대_layer정리_배관망.dxf",
    "samples/dxf/대명동201동 단위세대_layer정리_배관망_파이프넷버전.dxf",
    "data/sample_problem/대명동201동 계통도.dxf",
    "data/sample_problem/대명동201동 계통도_최소.dxf",
    # LH306 — 평면도/배관망 (계통도는 파탄 baseline)
    "samples/dxf/LH306동_평면도.dxf",
    "samples/dxf/LH306동_배관망.dxf",
    "samples/dxf/LH306동.dxf",
    "samples/dxf/계통도_LH_306.dxf",            # baseline (파탄)
    "samples/dxf/계통도_LH_306_배관망추출.dxf",  # 손그림 11선
    # 작은 fitting 테스트
    "samples/dxf/분기티.dxf",
]


def survey_one(path: Path):
    parsed = parse_dxf_for_view(path, include_hidden_layers=True)
    ents = parsed["entities"]
    graph, edge_len, scale = _raw_graph(ents)  # bridge 없는 raw 연결성
    n_nodes = len(graph)
    if n_nodes == 0:
        return {"file": path.name, "nodes": 0, "verdict": "EMPTY"}
    comps = _connected_components(graph)
    max_comp = max((len(c) for c in comps), default=0)
    deg1 = sum(1 for u, vs in graph.items() if len(vs) == 1)
    max_pct = max_comp / n_nodes * 100
    dang_pct = deg1 / n_nodes * 100
    good = max_pct >= 80.0 and dang_pct <= 35.0
    partial = max_pct >= 50.0
    verdict = "GOOD" if good else ("PARTIAL" if partial else "POOR")
    return {
        "file": path.name,
        "nodes": n_nodes,
        "edges": len(edge_len),
        "n_comp": len(comps),
        "max_pct": max_pct,
        "dang_pct": dang_pct,
        "scale": round(scale, 6),
        "verdict": verdict,
    }


def main(argv):
    cands = argv[1:] if len(argv) > 1 else DEFAULT_CANDIDATES
    rows = []
    print("=" * 104)
    print("깨끗한 학습용 도면 후보 채점 (bridge 전 raw 연결성)")
    print("=" * 104)
    for c in cands:
        p = Path(c)
        if not p.is_absolute():
            p = _ROOT / c
        if not p.is_file():
            print(f"  (없음) {c}")
            continue
        try:
            rows.append(survey_one(p))
        except Exception as exc:  # noqa: BLE001
            print(f"  (실패) {p.name}: {exc}")

    rows.sort(key=lambda r: (-(r.get("max_pct") or 0), r.get("dang_pct") or 1e9))
    print()
    hdr = f"{'verdict':<8} {'max%':>6} {'dang%':>6} {'n_comp':>7} {'nodes':>6} {'edges':>6} {'scale':>9}  file"
    print(hdr)
    print("-" * 104)
    for r in rows:
        if r.get("nodes", 0) == 0:
            print(f"{r['verdict']:<8} {'-':>6} {'-':>6} {'-':>7} {0:>6} {'-':>6} {'-':>9}  {r['file']}")
            continue
        print(f"{r['verdict']:<8} {r['max_pct']:>5.0f}% {r['dang_pct']:>5.0f}% "
              f"{r['n_comp']:>7} {r['nodes']:>6} {r['edges']:>6} {r['scale']:>9.5f}  {r['file']}")
    print("-" * 104)
    goods = [r for r in rows if r.get("verdict") == "GOOD"]
    print(f"\nGOOD(정답 신뢰 가능) 후보: {len(goods)}건")
    for r in goods:
        print(f"  · {r['file']}  (max {r['max_pct']:.0f}% · dangling {r['dang_pct']:.0f}% · "
              f"{r['nodes']}노드)")
    if not goods:
        print("  — 없음. 추출이 이미 깨끗한 도면이 없으면 합성손상의 정답 토대가 약함.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
