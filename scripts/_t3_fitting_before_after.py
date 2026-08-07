# -*- coding: utf-8 -*-
"""작업 3 실측 — 직류/분류티 구분 도입 전후의 부속 계상 비교.

새 규칙으로 만든 tables 와 **같은 입력**에서 옛 규칙(차수≥3 전량 티,
엘보 43.5~46.5/≥70 만 인정)을 그대로 재현해 나란히 센다. 한 번의 추출로
before/after 를 뽑으므로 형상 차이가 끼어들 여지가 없다.
"""
from __future__ import annotations

import sys
from collections import Counter, defaultdict
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))

import remote30_prototype as rp  # noqa: E402

TARGETS = [
    ("대명동 201동 단위세대", REPO / "samples/dxf/대명동201동 단위세대_layer정리.dxf"),
    ("B1F 현장조사 소화설비", REPO / "samples/dxf/B1F 현장조사 소화설비 평면도.dxf"),
]


def old_rule_counts(selection, tables, edge_key_to_pipe):
    elbows = Counter()
    for edge_key, recs in selection.elbow_fittings.items():
        if edge_key not in edge_key_to_pipe:
            continue
        for _pos, ang in recs:
            if 43.5 <= ang <= 46.5:
                elbows["elbow-45"] += 1
            elif ang >= 70:
                elbows["elbow"] += 1
    deg = Counter()
    for p in tables.pipes:
        deg[p["in"]] += 1
        deg[p["out"]] += 1
    tees = sum(1 for p in tables.pipes if deg[p["in"]] >= 3)
    elbows["tee"] = tees
    return elbows


def edge_map(selection, tables):
    """build_input_tables 와 같은 방식으로 edge_key→pipe 를 되만든다."""
    src = selection.source_pos
    parent_map, _d, order = rp.rooted_traversal(src, selection.edges)
    ordered = rp.traversal_ordered_edges(selection.edges, parent_map, order)
    return {(min(a, b), max(a, b)): str(10 + i)
            for i, (a, b, _L, _off) in enumerate(ordered)}


for title, path in TARGETS:
    if not path.exists():
        print(f"{title}: 파일 없음 — {path}")
        continue
    bundle = rp.parse_dxf_bundle(path)
    ents = rp.filter_pipenet_only(bundle)
    cats = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    sel = rp.select_worst30_heads(ents, cats)
    tables = rp.build_input_tables(sel, pipe_entities=ents, project_title=path.stem)

    after = Counter(f["type"] for f in tables.fittings)
    before = old_rule_counts(sel, tables, edge_map(sel, tables))
    meta = dict(tables.meta)

    print(f"\n=== {title} ===")
    print(f"  헤드 {len(sel.heads)} / 관로 {len(tables.pipes)}")
    for k in ("tee", "elbow", "elbow-45"):
        print(f"  {k:9s} before {before.get(k, 0):4d}  →  after {after.get(k, 0):4d}")
    print(f"  {'합계':9s} before {sum(before.values()):4d}  →  after {sum(after.values()):4d}")
    print(f"  판정 불가(엘보/티): {meta.get('부속 각도 판정 불가 (엘보 / 티)')}")
