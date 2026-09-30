"""Canonical straight runs shared by editing, inspection and ordinary export.

This is degree-two contraction, not a spanning tree or a visual grouping.
Lengths/losses are added; branches, bends, attachments and review stops survive.
Original rows remain available as evidence, while undo uses the editor history.
"""
from __future__ import annotations

from copy import deepcopy
from typing import Iterable

from .bore_provenance import describe_bore, evidence
from .export_compaction import CompactionAudit, T, compact_export


def _aggregate(rows: list[dict]) -> dict:
    """Keep all segment evidence without certifying a mixed/uncertain run."""
    info = [describe_bore(row) for row in rows]
    origins = []
    for row in rows:
        saved = evidence(row)
        if saved.get('continuous_sources'):
            origins.extend(deepcopy(saved['continuous_sources']))
        else:
            origins.append(deepcopy(row))
    sources = {r.get('source', 'unknown') for r in info}
    categories = {r['category'] for r in info}
    rec = dict(version=3, source=next(iter(sources)) if len(sources)==1 else 'mixed',
               auto_mm=rows[0].get('dia'), method='연속 동일관 통합',
               serial_category=next(iter(categories)) if len(categories)==1 else 'review',
               serial_edited=any(r['edited'] for r in info),
               review_only=any(r.get('review_only') or p.get('review_only') for r,p in zip(info,rows)),
               continuous_sources=origins)
    if all(r.get('policy')=='drawing_first_v1' for r in info):
        rec.update(policy='drawing_first_v1',block_export=any(r.get('block_export') for r in info),
                   source_edges=[e for r in info for e in r.get('source_edges',[])])
        definitions={r.get('definition_source') for r in info}
        rec['definition_source']=next(iter(definitions)) if len(definitions)==1 else 'drawing_mixed'
        if all(r.get('reason') == 'no_match' and r.get('block_export') for r in info):
            rec.update(reason='no_match', auto_mm=None, source='unresolved')
    if rec['source']=='review_default':
        rec['review_default_mm'] = rows[0].get('dia')
    # Preserve a uniform manual override; mixed edits remain individually visible.
    if all(r.get('manual') for r in info):
        rec['manual'] = dict(value_mm=rows[0].get('dia'), note='통합 전 모든 구간에 직접 입력 · 구간별 근거 참조')
        rec['serial_manual'] = True
    return rec


def compact_continuous(tables: T, *, keep_nodes: Iterable[str] = (),
                       allow_unassigned: bool = False) -> tuple[T, CompactionAudit]:
    """Return canonical tables on a copy; preserve deliberate editor split points."""
    keep = set(map(str, keep_nodes))
    for node in tables.nodes:
        # User-created points may be waiting for a branch/fitting in the next edit.
        if node.get('editor_added') or node.get('editor_note'):
            keep.add(str(node['label']))
        if node.get('editor_parent') is not None:
            keep.add(str(node['editor_parent']))
        keep.update(map(str, (node.get('editor_between') or [])[:2]))
    result, audit = compact_export(tables, keep_nodes=keep, allow_unassigned=allow_unassigned)
    for row in result.pipes:
        originals = audit.source_pipes[str(row['label'])]
        if len(originals)>1:
            row['bore_provenance'] = _aggregate(originals)
    return result, audit
