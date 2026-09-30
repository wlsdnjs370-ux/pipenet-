"""Explicit non-calculable plan exports; never invent missing bore/source data."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import json
import math
import xml.etree.ElementTree as ET
from typing import Any, Protocol, TypeVar

NOTICE = "평면도 검토용 — 관경 미지정 항목 및 실제 펌프·수원 조건은 PIPENET에서 정의해야 하며, 이 파일만으로 수리계산할 수 없습니다."


class ReviewTables(Protocol):
    """The graph-table contract shared by the conversion adapters."""
    nodes: list[dict[str, Any]]
    pipes: list[dict[str, Any]]
    nozzles: list[dict[str, Any]]


T = TypeVar('T', bound=ReviewTables)


def review_tables(tables: T) -> T:
    """Validate graph references and blank unknown/conflicting bores on a copy."""
    draft = deepcopy(tables)
    nodes = {str(n['label']) for n in draft.nodes}
    if len(nodes) != len(draft.nodes):
        raise ValueError("중복 노드가 있어 검토용 SDF를 만들 수 없습니다.")
    for node in draft.nodes:
        if any(not math.isfinite(float(node.get(key, 0))) for key in ('x','y','elevation')):
            raise ValueError(f"노드 {node['label']}: 좌표 또는 표고 값이 유한하지 않습니다.")
    labels = set()
    for pipe in draft.pipes:
        label = str(pipe['label'])
        if label in labels or {str(pipe['in']), str(pipe['out'])} - nodes:
            raise ValueError(f"배관 {label}: 중복 또는 연결 노드 누락")
        labels.add(label)
        for key in ('length', 'elev', 'c'):
            if not math.isfinite(float(pipe.get(key, 0))):
                raise ValueError(f"배관 {label}: {key} 값이 유한하지 않습니다.")
        if float(pipe.get('length', 0)) <= 0:
            raise ValueError(f"배관 {label}: 길이가 0 이하입니다.")
        if float(pipe.get('c', 120)) <= 0:
            raise ValueError(f"배관 {label}: C 값이 0 이하입니다.")
        if not math.isfinite(float(pipe.get('dia') or 0)) or float(pipe.get('dia') or 0) < 0:
            raise ValueError(f"배관 {label}: 관경 값이 올바르지 않습니다.")
        if not pipe.get('dia') or (pipe.get('bore_provenance') or {}).get('block_export'):
            pipe['dia'] = 0  # SDF's unassigned bore, not a hydraulic diameter.
            pipe.pop('inner_mm', None)
    for head in draft.nozzles:
        if str(head['in']) not in nodes:
            raise ValueError(f"헤드 {head['label']}: 연결 노드 누락")
    return draft


def mark_plan_review(path: str | Path, tables: ReviewTables) -> None:
    """Mark the SDF itself and retain all unresolved input evidence in a sidecar."""
    path = Path(path)
    tree = ET.parse(path)
    spray = tree.getroot().find('.//Network-spray')
    title = spray.find('Title')
    if title is None:
        title = ET.SubElement(spray, 'Title')
    title.text = 'PLAN REVIEW / NOT CALCULATION READY — ' + NOTICE
    tree.write(path, encoding='utf-8', xml_declaration=True)
    record = {'schema': 'module-h-plan-review-v1', 'calculation_ready': False,
              'notice': NOTICE, 'unknown_bore_labels': [str(p['label']) for p in tables.pipes
                  if not p.get('dia') or (p.get('bore_provenance') or {}).get('block_export')],
              'tables': {key: getattr(tables, key, []) for key in
                         ('nodes', 'pipes', 'nozzles', 'fittings', 'equipment')}}
    path.with_suffix('.review.json').write_text(json.dumps(record, ensure_ascii=False, indent=2, allow_nan=False), encoding='utf-8')
