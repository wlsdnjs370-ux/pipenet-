"""47 DXF 도면 전수 분석 — 일반화 작업용 패턴 학습.

수집:
  per DXF: entity 분포, layer 목록 + 카운트, bbox, INSERT 블록명, TEXT 값,
           기존 키워드 매칭 결과, 후보 신규 키워드, 텍스트 인코딩 통계

집계 (across 47):
  - 가장 흔한 layer 이름 / 패턴 (pipe-layer 후보)
  - 가장 흔한 INSERT 블록명 (head 블록 후보)
  - 텍스트 패턴 빈도 (직경, 층, 기기명)
  - DXF 간 bbox 스케일 분포

출력: data/reference_analysis.json
"""

from __future__ import annotations

import json
import re
import time
import traceback
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict, field
from pathlib import Path

from remote30_prototype import (
    parse_dxf_for_view,
    _auto_pipe_layer_filter,
    _extract_dia_text_points,
    _extract_floor_labels,
    SYSTEM_PIPE_LAYER_KEYWORDS,
)


REF_ROOT = Path("data/reference_library")


@dataclass
class DXFAnalysis:
    rel_path: str
    project: str
    folder: str
    name: str
    is_diagram: bool        # 계통도
    is_plan: bool           # 평면도
    # 파싱 결과
    parse_ok: bool = True
    parse_error: str = ""
    parse_seconds: float = 0.0
    file_size_kb: int = 0
    entity_count: int = 0
    entity_types: dict = field(default_factory=dict)
    layer_count: int = 0
    bbox: dict = field(default_factory=dict)
    # 레이어 분석
    line_layers_top: dict = field(default_factory=dict)   # LINE 이 들어있는 layer 의 entity 수
    auto_matched_layers: list = field(default_factory=list)  # 우리 키워드와 매칭된 layer
    # INSERT 블록명
    insert_blocks_top: dict = field(default_factory=dict)
    # 텍스트 분석
    text_count: int = 0
    text_distinct_values: int = 0
    dia_label_count: int = 0
    floor_label_count: int = 0
    # 후보 새 키워드 (현재 키워드와 매칭 안 되지만 비슷한 layer 이름)
    candidate_pipe_layers: list = field(default_factory=list)


def _project_from_path(p: Path) -> str:
    for part in p.parts:
        if "다이소" in part: return "다이소"
        if "양주옥정" in part: return "양주옥정"
    return "?"


def _classify_dxf(name: str) -> tuple[bool, bool]:
    return ("계통도" in name) or ("흐름도" in name), "평면도" in name


# 더 넓은 pipe-layer 휴리스틱 — 현재 keyword 외에 candidate 식별
_PIPE_KEYWORDS_EXTENDED = (
    # 현재 keywords
    "HSP", "LSP", "MSP", "LLSP", "SP", "배관", "PIPE", "RISER",
    "입상", "가지", "분기", "감압밸브",
    # candidate (양주옥정 MF-016 에서 발견)
    "Sprinkler", "Spr", "sprinkler", "SC ", "Sp-", "Sp ",
    "in-h", "OPLSP", "OPSP", "OPHSP", "OPLLS",
    # 한국어 더
    "스프링", "소화", "헤드망", "헤드선",
)


def _categorize_layer_name(layer: str) -> str:
    """레이어 이름 → 추정 카테고리."""
    n = layer.upper()
    if any(k in n for k in ("HEAD", "헤드", "스프링클러헤드")):
        return "HEAD"
    if any(k in n for k in ("TEXT", "DIM", "DIMENSION", "표")):
        return "TEXT"
    if any(k in n for k in ("벽", "건축", "WALL", "건물", "ARCH")):
        return "ARCH"
    if any(k in n for k in ("HSP", "LSP", "MSP", "LLSP", "SP", "배관", "PIPE", "RISER",
                             "입상", "가지", "분기", "감압",
                             "SPRINKLER", "OPSP", "OPLLS", "IN-H", "SP-", "SPR",
                             "스프링", "소화")):
        return "PIPE"
    return "OTHER"


def _layer_pipe_score(layer: str, entity_count: int) -> float:
    """layer 가 pipe layer 일 가능성 점수 — 이름 + entity 수 기반."""
    cat = _categorize_layer_name(layer)
    if cat != "PIPE":
        return 0.0
    if entity_count < 5:
        return 0.3
    if entity_count < 30:
        return 0.6
    return 1.0


def analyze_dxf(path: Path) -> DXFAnalysis:
    """한 DXF 의 전수 분석."""
    rel = path.relative_to(REF_ROOT)
    name = path.name
    is_diag, is_plan = _classify_dxf(name)
    a = DXFAnalysis(
        rel_path=str(rel),
        project=_project_from_path(path),
        folder=str(path.parent.relative_to(REF_ROOT)),
        name=name,
        is_diagram=is_diag,
        is_plan=is_plan,
        file_size_kb=int(path.stat().st_size / 1024),
    )
    t0 = time.time()
    try:
        parsed = parse_dxf_for_view(path, include_hidden_layers=True)
    except Exception as e:
        a.parse_ok = False
        a.parse_error = str(e)[:300]
        a.parse_seconds = round(time.time() - t0, 1)
        return a
    a.parse_seconds = round(time.time() - t0, 1)

    ents = parsed.get("entities", [])
    layers_info = parsed.get("layers", [])
    a.entity_count = len(ents)
    a.layer_count = len(layers_info)
    a.bbox = parsed.get("bbox", {})

    # entity types
    a.entity_types = dict(Counter(e.get("t") for e in ents).most_common())

    # LINE 이 있는 layer 별 LINE 개수
    line_by_layer = Counter()
    for e in ents:
        if e.get("t") in ("L", "PL"):
            line_by_layer[e.get("l", "?")] += 1
    a.line_layers_top = dict(line_by_layer.most_common(15))

    # 우리 자동 키워드 매칭 결과
    a.auto_matched_layers = sorted(_auto_pipe_layer_filter(ents))

    # INSERT 블록명 top
    blocks = Counter(e.get("n", "?") for e in ents if e.get("t") == "I")
    a.insert_blocks_top = dict(blocks.most_common(20))

    # TEXT 분석
    texts = [e for e in ents if e.get("t") in ("T", "M")]
    a.text_count = len(texts)
    distinct_vals = set((e.get("v") or "").strip() for e in texts)
    distinct_vals.discard("")
    a.text_distinct_values = len(distinct_vals)

    # 직경 라벨 / 층 라벨
    try:
        dias = _extract_dia_text_points(ents)
        floors = _extract_floor_labels(ents)
        a.dia_label_count = len(dias)
        a.floor_label_count = len(floors)
    except Exception:
        pass

    # 후보 신규 pipe layer (우리 자동 매칭에 안 잡혔지만 점수 ≥ 0.6)
    candidates = []
    matched_set = set(a.auto_matched_layers)
    for layer, n in line_by_layer.most_common(50):
        if layer in matched_set:
            continue
        score = _layer_pipe_score(layer, n)
        if score >= 0.6:
            candidates.append({"layer": layer, "lines": n, "score": score})
    a.candidate_pipe_layers = candidates[:10]
    return a


def aggregate(analyses: list[DXFAnalysis]) -> dict:
    """47 분석 결과 → 집계 인사이트."""
    ok = [a for a in analyses if a.parse_ok]
    failed = [a for a in analyses if not a.parse_ok]

    # 전체 매칭된 layer 빈도 (여러 도면에 걸쳐 같은 layer 이름이 매칭됐는지)
    matched_layer_freq: Counter = Counter()
    for a in ok:
        for ly in a.auto_matched_layers:
            matched_layer_freq[ly] += 1

    # 후보 새 키워드 — 여러 도면에서 추천된 layer
    candidate_freq: Counter = Counter()
    candidate_lines: defaultdict = defaultdict(int)
    for a in ok:
        for cand in a.candidate_pipe_layers:
            candidate_freq[cand["layer"]] += 1
            candidate_lines[cand["layer"]] += cand["lines"]

    # 블록명 전수
    block_freq: Counter = Counter()
    for a in ok:
        for bn, cnt in a.insert_blocks_top.items():
            block_freq[bn] += cnt

    # entity 분포
    total_entities = sum(a.entity_count for a in ok)
    total_layers = sum(a.layer_count for a in ok)
    total_texts = sum(a.text_count for a in ok)
    total_dia_labels = sum(a.dia_label_count for a in ok)
    total_floor_labels = sum(a.floor_label_count for a in ok)

    # 매칭 layer 단어 추출 — 토큰화하여 자주 등장하는 단어
    word_freq: Counter = Counter()
    for ly, freq in matched_layer_freq.items():
        for tok in re.findall(r"[A-Za-z가-힣\-]+", ly):
            if len(tok) >= 2:
                word_freq[tok] += freq

    return {
        "parsed_ok_count": len(ok),
        "parsed_fail_count": len(failed),
        "parse_failures": [{"path": a.rel_path, "error": a.parse_error} for a in failed],
        "totals": {
            "entities": total_entities,
            "layers": total_layers,
            "texts": total_texts,
            "dia_labels": total_dia_labels,
            "floor_labels": total_floor_labels,
        },
        "matched_layer_freq_top": dict(matched_layer_freq.most_common(40)),
        "candidate_new_pipe_layers": [
            {"layer": l, "doc_count": candidate_freq[l], "total_lines": candidate_lines[l]}
            for l in candidate_freq if candidate_freq[l] >= 1
        ][:30],
        "top_insert_blocks": dict(block_freq.most_common(50)),
        "pipe_word_freq": dict(word_freq.most_common(40)),
    }


def main():
    dxfs = sorted(REF_ROOT.rglob("*.dxf"))
    print(f"=== {time.strftime('%H:%M:%S')} 분석 시작 — {len(dxfs)} DXF ===", flush=True)
    analyses: list[DXFAnalysis] = []
    t0 = time.time()
    for i, p in enumerate(dxfs, 1):
        try:
            a = analyze_dxf(p)
            analyses.append(a)
            tag = "OK " if a.parse_ok else "FAIL"
            print(f"  [{i}/{len(dxfs)}] {tag} {a.parse_seconds:5.1f}s "
                  f"{a.file_size_kb:>6}KB ents={a.entity_count:>7} layers={a.layer_count:>3} "
                  f"matched={len(a.auto_matched_layers):>2} cand={len(a.candidate_pipe_layers):>2} "
                  f"text={a.text_count:>4} dia={a.dia_label_count:>3} floor={a.floor_label_count:>3} "
                  f"| {a.rel_path}", flush=True)
        except Exception as e:
            print(f"  [{i}/{len(dxfs)}] EXC {p.name}: {e}", flush=True)
            traceback.print_exc()

    print(f"\n=== {time.strftime('%H:%M:%S')} 집계 ===", flush=True)
    agg = aggregate(analyses)

    out = {
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "total_seconds": round(time.time() - t0, 1),
        "aggregate": agg,
        "per_dxf": [asdict(a) for a in analyses],
    }
    out_path = Path("data/reference_analysis.json")
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n=== 완료 — 총 {round(time.time()-t0,1)}s ===", flush=True)
    print(f"  결과: {out_path}", flush=True)

    # 요약
    print(f"\n=== 집계 요약 ===")
    print(f"  파싱 OK {agg['parsed_ok_count']} / FAIL {agg['parsed_fail_count']}")
    print(f"  total entities: {agg['totals']['entities']:,}")
    print(f"  total layers:   {agg['totals']['layers']}")
    print(f"  total texts:    {agg['totals']['texts']:,}")
    print(f"  dia labels:     {agg['totals']['dia_labels']}")
    print(f"  floor labels:   {agg['totals']['floor_labels']}")

    print(f"\n=== 매칭된 layer 이름 (top 20, 도면 빈도) ===")
    for ly, c in list(agg['matched_layer_freq_top'].items())[:20]:
        print(f"  {c:>3}x  {ly}")

    print(f"\n=== 후보 신규 pipe layer (현재 매칭 X, 점수 ≥ 0.6, top 20) ===")
    for c in agg['candidate_new_pipe_layers'][:20]:
        print(f"  {c['doc_count']:>3}x ({c['total_lines']:>5} LINE) — {c['layer']}")

    print(f"\n=== pipe word freq (자주 같이 등장하는 단어, top 25) ===")
    for w, c in list(agg['pipe_word_freq'].items())[:25]:
        print(f"  {c:>4}x  {w}")


if __name__ == "__main__":
    main()
