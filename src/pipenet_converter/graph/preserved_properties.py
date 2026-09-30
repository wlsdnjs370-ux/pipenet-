"""Cycle-safe explicit property updates without tree-coordinate relayout."""
from __future__ import annotations

import math
from typing import Protocol


class ReviewTables(Protocol):
    """Only the structured rows needed by cycle-safe property edits."""
    nodes: list[dict]
    pipes: list[dict]
    pipe_labels: dict


def apply_review_properties(tables: ReviewTables, keys: dict, rows: list) -> None:
    """Update table properties, retaining physical XY and checking every rise."""
    pipe_keys = {tuple(k):str(lab) for lab,k in keys.get('pipe',{}).items()}
    node_keys = {tuple(k):str(lab) for lab,k in keys.get('node',{}).items()}
    pipes = {str(p['label']):p for p in tables.pipes}
    nodes = {str(n['label']):n for n in tables.nodes}
    for row in rows:
        field, key = row.get('field'), tuple(row.get('key') or ())
        mapping, table = (node_keys,nodes) if field=='elevation' else (pipe_keys,pipes)
        label = mapping.get(key)
        if label not in table:
            raise ValueError('저장된 요소 수정이 현재 루프 선정 범위에 없습니다. 편집 기록을 확인하세요.')
        value = row.get('new')
        if field != 'type' and (not math.isfinite(float(value)) or
                                (field in ('dia','length','c') and float(value)<=0)):
            raise ValueError('요소 수정값이 유효하지 않습니다.')
        table[label][field] = value
        if field == 'dia':
            evidence = table[label].setdefault('bore_provenance',{})
            evidence.update(source='user',confirmed=True,block_export=False,
                            manual={'value_mm':int(value),'note':row.get('reason',row.get('note',''))})
            table[label]['dia_src']='user'
    for pipe in tables.pipes:
        pipe['elev'] = round(float(nodes[str(pipe['out'])]['elevation'])-float(nodes[str(pipe['in'])]['elevation']),6)
        if abs(pipe['elev']) > float(pipe['length'])+1e-6:
            raise ValueError(f"{pipe['label']}: 표고차가 배관 길이보다 큽니다. 길이·표고를 함께 확인하세요.")


def sync_review_kfp(tables: ReviewTables, net: dict, node_labels: dict[str, str]) -> None:
    """Copy explicit table values back to the same graph, never reconstruct it."""
    for row in tables.nodes:
        node = net['nodes_meta_runtime'][node_labels[str(row['label'])]]
        node['coords'][2] = node['elevation_m'] = float(row['elevation'])
    reverse = {lab:pid for pid,lab in tables.pipe_labels.items()}
    for row in tables.pipes:
        pipe = net['pipe_data'][reverse[row['label']]]
        pipe.update(length_m=float(row['length']),C=float(row['c']),nominal_mm=int(row['dia']),
                    pipe_type=row['type'],bore_provenance=dict(row.get('bore_provenance') or {}))
