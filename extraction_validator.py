"""검증 하네스 — 우리 계통도 추출 vs 답안 SDF 위상/통계 비교.

답안 SDF 의 노드 좌표는 PIPENET schematic 의 logical position (도면 좌표 아님).
따라서 좌표 직접 매핑 불가, 대신 다음 지표로 비교:

  - 노드 / 파이프 / 노즐 / 밸브 개수
  - 총 길이 (m)
  - elev range (m)
  - 직경 분포 (bore_mm bucket)
  - 직경 별 길이 분포 (hydraulic 가중)
  - "라이저만" 길이 (직경 ≥ 100mm 일반)

CompareResult 로 한 case 결과 정리.
"""

from __future__ import annotations

import json
from collections import Counter
from dataclasses import dataclass, asdict
from pathlib import Path

from sdf_reader import read_sdf, AnswerNetwork
from remote30_prototype import parse_dxf_for_view, extract_system_path


@dataclass
class CaseSpec:
    """검증 한 case 의 입력."""
    label: str
    dxf_path: str
    answer_sdf_path: str
    pump_xy: tuple[float, float]   # 도면 좌표 (mm)
    av_xy: tuple[float, float]


@dataclass
class StatBlock:
    """우리 추출 결과 / 답안 SDF 의 공통 비교 지표."""
    source: str        # "our_extract" / "answer_sdf" / "answer_riser_only"
    node_count: int
    pipe_count: int
    nozzle_count: int
    valve_count: int
    total_length_m: float
    elev_min_m: float
    elev_max_m: float
    elev_span_m: float
    bore_distribution_mm: dict       # bore_mm → pipe count
    bore_length_mm: dict             # bore_mm → total length m
    main_bore_mm: int                # 가장 긴 직경
    main_bore_length_m: float

    @staticmethod
    def from_answer(net: AnswerNetwork, source: str = "answer_sdf") -> "StatBlock":
        elev_min, elev_max = net.elev_range
        bld = net.bore_length_distribution_mm
        main_bore_mm = 0; main_bore_length_m = 0.0
        if bld:
            main_bore_mm, main_bore_length_m = max(bld.items(), key=lambda kv: kv[1])
        return StatBlock(
            source=source,
            node_count=len(net.nodes),
            pipe_count=len(net.pipes),
            nozzle_count=len(net.nozzles),
            valve_count=len(net.valves),
            total_length_m=round(net.total_length_m, 2),
            elev_min_m=round(elev_min, 2),
            elev_max_m=round(elev_max, 2),
            elev_span_m=round(elev_max - elev_min, 2),
            bore_distribution_mm={int(k): v for k, v in net.bore_distribution_mm.items()},
            bore_length_mm={int(k): round(v, 2) for k, v in bld.items()},
            main_bore_mm=int(main_bore_mm),
            main_bore_length_m=round(main_bore_length_m, 2),
        )

    @staticmethod
    def from_our_extract(riser: dict, source: str = "our_extract") -> "StatBlock":
        nodes = riser.get("nodes", [])
        pipes = riser.get("pipes", [])
        elev = [n.get("elevation", 0.0) for n in nodes]
        elev_min = min(elev) if elev else 0.0
        elev_max = max(elev) if elev else 0.0
        # bore 분포 — pipe.dia 는 mm 단위
        bd = Counter(int(p.get("dia", 0)) for p in pipes)
        bl: dict[int, float] = {}
        for p in pipes:
            dia = int(p.get("dia", 0))
            bl[dia] = bl.get(dia, 0.0) + float(p.get("length", 0.0))
        main_bore_mm = 0; main_bore_length_m = 0.0
        if bl:
            main_bore_mm, main_bore_length_m = max(bl.items(), key=lambda kv: kv[1])
        return StatBlock(
            source=source,
            node_count=len(nodes),
            pipe_count=len(pipes),
            nozzle_count=0,            # 우리 path 추출은 nozzle 안 만듬
            valve_count=0,
            total_length_m=round(riser.get("total_pipe_length_m", 0.0), 2),
            elev_min_m=round(elev_min, 2),
            elev_max_m=round(elev_max, 2),
            elev_span_m=round(elev_max - elev_min, 2),
            bore_distribution_mm=dict(bd),
            bore_length_mm={k: round(v, 2) for k, v in bl.items()},
            main_bore_mm=main_bore_mm,
            main_bore_length_m=round(main_bore_length_m, 2),
        )


def riser_only_subnet(net: AnswerNetwork, min_bore_mm: int = 100) -> AnswerNetwork:
    """답안 SDF 에서 라이저만 (직경 ≥ min_bore_mm). 가지관/헤드망 제외."""
    riser_pipes = [p for p in net.pipes if p.bore_mm >= min_bore_mm]
    pipe_nodes = set()
    for p in riser_pipes:
        pipe_nodes.add(p.input_node)
        pipe_nodes.add(p.output_node)
    riser_nodes = [n for n in net.nodes if n.label in pipe_nodes]
    return AnswerNetwork(
        title=f"{net.title} (riser ≥{min_bore_mm}mm)",
        nodes=riser_nodes,
        pipes=riser_pipes,
        nozzles=[],   # 라이저 부분에 노즐 없음
        valves=[v for v in net.valves
                if v.input_node in pipe_nodes and v.output_node in pipe_nodes],
        sdf_version=net.sdf_version,
        source_path=net.source_path,
    )


def compare_case(spec: CaseSpec, min_bore_mm: int = 100) -> dict:
    """한 case 의 baseline 비교."""
    print(f"\n=== {spec.label} ===")

    # 1) 우리 알고리즘 추출
    dxf_p = Path(spec.dxf_path)
    print(f"  DXF: {dxf_p.name} ({dxf_p.stat().st_size//1024} KB)")
    parsed = parse_dxf_for_view(dxf_p, include_hidden_layers=True)
    print(f"  entities: {len(parsed['entities']):,}, layers: {len(parsed.get('layers',[]))}")
    try:
        our_riser = extract_system_path(parsed["entities"], spec.pump_xy, spec.av_xy,
                                         snap_tolerance_mm=5000.0)
        our_stat = StatBlock.from_our_extract(our_riser, source="our_extract")
        our_error = None
    except Exception as e:
        our_riser = None
        our_stat = None
        our_error = str(e)[:200]
        print(f"  OUR ERROR: {our_error}")

    # 2) 답안 SDF 읽기
    ans_p = Path(spec.answer_sdf_path)
    print(f"  Answer: {ans_p.name}")
    ans_net = read_sdf(ans_p)
    ans_stat = StatBlock.from_answer(ans_net, source="answer_full")
    ans_riser_net = riser_only_subnet(ans_net, min_bore_mm=min_bore_mm)
    ans_riser_stat = StatBlock.from_answer(ans_riser_net, source=f"answer_riser_{min_bore_mm}mm+")

    # 3) Comparison metrics
    comp: dict = {}
    if our_stat is not None:
        # 노드/파이프/길이 비율
        comp["node_ratio_our_vs_answer_full"] = round(our_stat.node_count / max(1, ans_stat.node_count), 3)
        comp["node_ratio_our_vs_answer_riser"] = round(our_stat.node_count / max(1, ans_riser_stat.node_count), 3)
        comp["pipe_ratio_our_vs_answer_full"] = round(our_stat.pipe_count / max(1, ans_stat.pipe_count), 3)
        comp["pipe_ratio_our_vs_answer_riser"] = round(our_stat.pipe_count / max(1, ans_riser_stat.pipe_count), 3)
        comp["length_ratio_our_vs_answer_full"] = round(our_stat.total_length_m / max(0.1, ans_stat.total_length_m), 3)
        comp["length_ratio_our_vs_answer_riser"] = round(our_stat.total_length_m / max(0.1, ans_riser_stat.total_length_m), 3)
        comp["length_diff_pct_vs_riser"] = round((our_stat.total_length_m - ans_riser_stat.total_length_m) / max(0.1, ans_riser_stat.total_length_m) * 100, 1)
        comp["elev_span_diff_m"] = round(our_stat.elev_span_m - ans_stat.elev_span_m, 2)
        # 직경 일치
        comp["main_bore_match"] = our_stat.main_bore_mm == ans_riser_stat.main_bore_mm
        comp["our_main_bore_mm"] = our_stat.main_bore_mm
        comp["answer_riser_main_bore_mm"] = ans_riser_stat.main_bore_mm
        # bore overlap (직경 종류 일치 IoU)
        our_bores = set(b for b in our_stat.bore_distribution_mm if b > 0)
        ans_bores = set(b for b in ans_riser_stat.bore_distribution_mm if b > 0)
        if ans_bores or our_bores:
            iou = len(our_bores & ans_bores) / max(1, len(our_bores | ans_bores))
        else:
            iou = 1.0
        comp["bore_diversity_iou_vs_riser"] = round(iou, 3)

    result = {
        "case": spec.label,
        "dxf": dxf_p.name,
        "answer_sdf": ans_p.name,
        "our": asdict(our_stat) if our_stat else None,
        "our_error": our_error,
        "answer_full": asdict(ans_stat),
        "answer_riser_only": asdict(ans_riser_stat),
        "comparison": comp,
    }
    return result


def print_case_summary(result: dict) -> None:
    print(f"\n  ── 결과 ──")
    if result["our"] is None:
        print(f"  OUR: 실패 ({result.get('our_error', 'unknown')})")
        return
    o = result["our"]; a = result["answer_full"]; ar = result["answer_riser_only"]
    print(f"  {'':25s} {'우리':>15s}  {'답안(전체)':>15s}  {'답안(라이저)':>15s}")
    print(f"  {'노드 수':25s} {o['node_count']:>15}  {a['node_count']:>15}  {ar['node_count']:>15}")
    print(f"  {'파이프 수':25s} {o['pipe_count']:>15}  {a['pipe_count']:>15}  {ar['pipe_count']:>15}")
    print(f"  {'노즐 (헤드) 수':25s} {o['nozzle_count']:>15}  {a['nozzle_count']:>15}  {ar['nozzle_count']:>15}")
    print(f"  {'총 길이 (m)':25s} {o['total_length_m']:>15.2f}  {a['total_length_m']:>15.2f}  {ar['total_length_m']:>15.2f}")
    print(f"  {'elev range (m)':25s} {o['elev_max_m']-o['elev_min_m']:>15.2f}  {a['elev_max_m']-a['elev_min_m']:>15.2f}  {ar['elev_max_m']-ar['elev_min_m']:>15.2f}")
    print(f"  {'주 직경 (mm)':25s} {o['main_bore_mm']:>15}  {a['main_bore_mm']:>15}  {ar['main_bore_mm']:>15}")
    c = result["comparison"]
    print(f"\n  ── 비교 (vs 답안 라이저 부분) ──")
    print(f"    노드 비율: {c.get('node_ratio_our_vs_answer_riser','?')}")
    print(f"    파이프 비율: {c.get('pipe_ratio_our_vs_answer_riser','?')}")
    print(f"    길이 차이: {c.get('length_diff_pct_vs_riser','?')}%")
    print(f"    주 직경 일치: {c.get('main_bore_match','?')} (우리 {c.get('our_main_bore_mm','?')}mm vs 답안 {c.get('answer_riser_main_bore_mm','?')}mm)")
    print(f"    직경 다양성 IoU: {c.get('bore_diversity_iou_vs_riser','?')}")


# ── 검증 case 정의 ──
CASES = [
    CaseSpec(
        label="대명동 201동 — 28F 자연낙차 라이저",
        dxf_path="data/sample_problem/대명동201동 계통도.dxf",
        answer_sdf_path="data/sample_problem/MSP 중층부(17,28층)/1-1. 업무시설 201동_28F (자연낙차)-RV03_NEW.sdf",
        pump_xy=(-644264.212, 178419.8529),
        av_xy=(-660081.0832, 133254.6858),
    ),
    # 양주옥정 / 다이소 추가는 도면 좌표 파악 후 — 우선 대명동만 시연
]


def main():
    results = []
    for spec in CASES:
        try:
            r = compare_case(spec)
            print_case_summary(r)
            results.append(r)
        except Exception as e:
            print(f"  EXC: {e}")
            import traceback; traceback.print_exc()
    out = Path("data/validation_baseline.json")
    out.write_text(json.dumps({"baseline_results": results}, ensure_ascii=False, indent=2),
                   encoding="utf-8")
    print(f"\n저장: {out}")


if __name__ == "__main__":
    main()
