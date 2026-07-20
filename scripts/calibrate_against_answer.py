"""Remote 30 extractor 자동 미세 조정 — 답안지 SDF 와 비교.

답안지: 1-1. 업무시설 201동_3F (자연낙차_감압_저층부)-RV03_NEW.sdf
입력 DXF: static/대명동201동 단위세대_layer정리.dxf

비교 metric:
  - 노드/배관/노즐 수
  - 배관 길이 분포 (평균, 중앙값, 합계)
  - 배관 직경 분포 (히스토그램)
  - 노즐 input_node elevation
  - 알람밸브 좌표

loss = sum of normalized diffs
조정 파라미터 grid search → 최소 loss 조합 보고
"""

from __future__ import annotations

import math
import statistics
import sys
import xml.etree.ElementTree as ET
from collections import Counter
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))
ANSWER_SDF = PROJECT_ROOT / "1-1. 업무시설 201동_3F (자연낙차_감압_저층부)-RV03_NEW.sdf"
INPUT_DXF = PROJECT_ROOT / "static" / "대명동201동 단위세대_layer정리.dxf"
OUT_DIR = PROJECT_ROOT / "data" / "remote30_outputs"


def parse_sdf(path: Path) -> dict:
    """SDF XML 파싱 → 통계 dict."""
    tree = ET.parse(str(path))
    root = tree.getroot()
    nodes = []
    pipes = []
    nozzles = []
    valves = []
    for n in root.iter("Node"):
        pos = n.find("Position")
        if pos is None:
            continue
        nodes.append({
            "label": n.get("label"),
            "elevation": float(n.get("elevation", 0)),
            "x": float(pos.get("x", 0)),
            "y": float(pos.get("y", 0)),
        })
    for p in root.iter("Pipe"):
        pipes.append({
            "label": p.get("label"),
            "input": p.get("input"),
            "output": p.get("output"),
            "bore_m": float(p.get("bore", 0)),
            "length_m": float(p.get("length", 0)),
            "rise_m": float(p.get("rise", 0)),
            "c": float(p.get("roughness-or-c", 120)),
        })
    for nz in root.iter("Nozzle"):
        nozzles.append({
            "label": nz.get("label"),
            "input": nz.get("input"),
            "output": nz.get("output"),
            "status": nz.get("status"),
        })
    # PIPENET SDF 에 Valve 가 없는 경우 있음 — alarm valve 는 별도
    for v in root.iter("Valve"):
        valves.append({
            "label": v.get("label"),
        })

    # 좌표 정규화 (가장 작은 좌표를 0으로)
    if nodes:
        xs = [n["x"] for n in nodes]
        ys = [n["y"] for n in nodes]
        x_min, x_max = min(xs), max(xs)
        y_min, y_max = min(ys), max(ys)
        x_span = x_max - x_min
        y_span = y_max - y_min
    else:
        x_min = x_max = y_min = y_max = x_span = y_span = 0

    bore_hist = Counter(round(p["bore_m"] * 1000) for p in pipes)
    lengths = [p["length_m"] for p in pipes if p["length_m"] > 0]

    return {
        "path": str(path),
        "n_nodes": len(nodes),
        "n_pipes": len(pipes),
        "n_nozzles": len(nozzles),
        "n_valves": len(valves),
        "x_range": (x_min, x_max, x_span),
        "y_range": (y_min, y_max, y_span),
        "bore_mm_hist": dict(bore_hist),
        "length_total": sum(lengths),
        "length_mean": statistics.mean(lengths) if lengths else 0,
        "length_median": statistics.median(lengths) if lengths else 0,
        "length_min_max": (min(lengths), max(lengths)) if lengths else (0, 0),
        "nodes": nodes,
        "pipes": pipes,
        "nozzles": nozzles,
    }


def print_summary(label: str, s: dict) -> None:
    print(f"\n=== {label} ===")
    print(f"  nodes:   {s['n_nodes']}")
    print(f"  pipes:   {s['n_pipes']}")
    print(f"  nozzles: {s['n_nozzles']}")
    print(f"  x range: {s['x_range'][0]:.1f} ~ {s['x_range'][1]:.1f}  span {s['x_range'][2]:.1f}")
    print(f"  y range: {s['y_range'][0]:.1f} ~ {s['y_range'][1]:.1f}  span {s['y_range'][2]:.1f}")
    print(f"  bore_mm distribution: {s['bore_mm_hist']}")
    print(f"  pipe length: total={s['length_total']:.2f}m  mean={s['length_mean']:.2f}m  median={s['length_median']:.2f}m  min/max={s['length_min_max'][0]:.2f}/{s['length_min_max'][1]:.2f}m")
    # 길이 상위 5 pipe
    long_pipes = sorted(s["pipes"], key=lambda p: -p["length_m"])[:5]
    print(f"  longest 5 pipes:")
    for p in long_pipes:
        print(f"    label={p['label']:>5}  len={p['length_m']:>6.2f}m  rise={p['rise_m']:>+6.2f}m  bore={p['bore_m']*1000:>4.0f}mm  io={p['input']}→{p['output']}")
    # rise 합
    total_rise = sum(p["rise_m"] for p in s["pipes"])
    pos_rise = sum(p["rise_m"] for p in s["pipes"] if p["rise_m"] > 0)
    neg_rise = sum(p["rise_m"] for p in s["pipes"] if p["rise_m"] < 0)
    print(f"  rise total={total_rise:+.2f}m  pos={pos_rise:.2f}m  neg={neg_rise:.2f}m")


def diff_metrics(answer: dict, ours: dict) -> dict:
    """답안지 vs 우리 결과의 차이를 정량화."""
    metrics = {}
    metrics["d_nodes"]   = ours["n_nodes"]   - answer["n_nodes"]
    metrics["d_pipes"]   = ours["n_pipes"]   - answer["n_pipes"]
    metrics["d_nozzles"] = ours["n_nozzles"] - answer["n_nozzles"]
    metrics["d_length_total"] = ours["length_total"] - answer["length_total"]
    metrics["d_length_mean"]  = ours["length_mean"]  - answer["length_mean"]

    # 직경 분포 차이 - 같은 직경이 얼마나 비슷한가
    a_hist = answer["bore_mm_hist"]
    o_hist = ours["bore_mm_hist"]
    all_keys = set(a_hist) | set(o_hist)
    bore_diff = sum(abs(a_hist.get(k, 0) - o_hist.get(k, 0)) for k in all_keys)
    metrics["bore_diff"] = bore_diff

    # 총 loss (정규화)
    loss = (
        abs(metrics["d_nodes"]) / max(answer["n_nodes"], 1) +
        abs(metrics["d_pipes"]) / max(answer["n_pipes"], 1) +
        abs(metrics["d_nozzles"]) / max(answer["n_nozzles"], 1) +
        abs(metrics["d_length_total"]) / max(answer["length_total"], 1) +
        bore_diff / max(sum(a_hist.values()), 1)
    )
    metrics["loss"] = loss
    return metrics


def run_extraction_with(overrides: dict) -> dict:
    """Extractor 호출 → 결과 SDF 파싱."""
    from sprinkler_remote30_extractor import run_remote30_extraction
    r = run_remote30_extraction(
        dxf_path=str(INPUT_DXF),
        alarm_xy=None,
        out_dir=OUT_DIR,
        overrides={**{"emit_sdf": True, "emit_csv": True}, **overrides},
    )
    sdf_path = r.get("sdf_path")
    if not sdf_path:
        return None
    return parse_sdf(Path(sdf_path)), r


def main():
    print(f"답안지: {ANSWER_SDF}")
    print(f"입력 DXF: {INPUT_DXF}")

    answer = parse_sdf(ANSWER_SDF)
    print_summary("답안지", answer)

    # 1차: 기본 설정으로 추출
    print("\n\n[1차] 기본 설정으로 추출")
    ours_sdf, raw = run_extraction_with({"remote_mode": "hydraulic"})
    print_summary("우리 (기본)", ours_sdf)
    m = diff_metrics(answer, ours_sdf)
    print(f"\n  diff: {m}")

    # 3차: snap_tol 미세 조정 + 자연낙차 137.35 고정
    candidates = []
    for snap in [300, 400, 500, 600, 700, 800, 1000, 1200]:
        candidates.append({
            "snap_tol": snap,
            "head_to_pipe_tol": 800,
            "diameter_text_search_radius": 1500,
            "natural_fall_height_m": 137.35,
            "elevation_head_m": 2.7,  # 답안지 정확치
        })
    print(f"\n\n[grid search] {len(candidates)}개 조합 시도")
    print(f"  {'snap':>5} {'fall':>6} {'dia_r':>6} | {'nodes':>5} {'pipes':>5} {'noz':>3} | {'len_tot':>8} {'rise':>8} | {'loss':>6}")
    results = []
    for c in candidates:
        try:
            ours_sdf, _ = run_extraction_with({**c, "remote_mode": "hydraulic"})
        except Exception as e:
            print(f"  ERROR with {c}: {e}")
            continue
        m = diff_metrics(answer, ours_sdf)
        results.append((c, ours_sdf, m))
        rise_tot = sum(p["rise_m"] for p in ours_sdf["pipes"])
        print(f"  {c['snap_tol']:>5} {c['natural_fall_height_m']:>6.1f} {c['diameter_text_search_radius']:>6} | "
              f"{ours_sdf['n_nodes']:>5} {ours_sdf['n_pipes']:>5} {ours_sdf['n_nozzles']:>3} | "
              f"{ours_sdf['length_total']:>8.1f} {rise_tot:>+8.1f} | {m['loss']:>6.3f}")

    # 최저 loss
    best = min(results, key=lambda t: t[2]["loss"])
    print(f"\n  ★ best: snap={best[0]['snap_tol']} fall={best[0]['natural_fall_height_m']:.1f}m  → loss={best[2]['loss']:.3f}")
    print(f"\n  답안지 vs best:")
    print(f"    nodes:   {best[1]['n_nodes']} vs {answer['n_nodes']}  (diff={best[2]['d_nodes']:+d})")
    print(f"    pipes:   {best[1]['n_pipes']} vs {answer['n_pipes']}  (diff={best[2]['d_pipes']:+d})")
    print(f"    nozzles: {best[1]['n_nozzles']} vs {answer['n_nozzles']}  (diff={best[2]['d_nozzles']:+d})")
    print(f"    length:  {best[1]['length_total']:.1f}m vs {answer['length_total']:.1f}m  (diff={best[2]['d_length_total']:+.1f}m)")


if __name__ == "__main__":
    main()
