"""Reconcile unresolved source spots with explicit, endpoint-bound user fittings.

Keep the original issues: removing a fitting or undoing an edit must restore the
blocker. A nearby fitting, valve, or pipe-only fitting does not resolve a spot.
This changes review state only, never topology, elevations or loss quantities.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass
from typing import Any, Mapping


@dataclass(frozen=True)
class FittingReview:
    """Current pending spots and auditable manual resolutions."""

    pending: tuple[dict, ...]
    resolved: tuple[dict, ...]
    unlocated_count: int

    @property
    def count(self) -> int:
        """Number of still-unconfirmed fittings, including legacy count-only data."""
        return self.unlocated_count + sum(_count(r.get('n', 1)) for r in self.pending)


def _count(value: Any) -> int:
    number = float(value)
    if not number.is_integer() or number < 0:
        raise ValueError('부속 미확정 개수는 0 이상의 정수여야 합니다.')
    return int(number)


def fitting_review(tables: Any) -> FittingReview:
    """Resolve only exact pipe + endpoint spots with a user-selected loss fitting."""
    source = getattr(tables, 'unresolved', None) or {}
    issues = source.get('kind_items') or []
    reported = max([_count(v) for k, v in (getattr(tables, 'meta', None) or [])
                    if k == '부속 판정 불가'] or [0])
    unlocated = _count(source.get('kind_unlocated_count',
        max(0, reported - sum(_count(r.get('n', 1)) for r in issues))))
    pipes = {str(p['label']): p for p in tables.pipes}
    equipment: dict[tuple[str, str], list[dict]] = {}
    allowed = {'ELBOW_90_STD', 'ELBOW_45', 'TEE_BRANCH'}
    for row in tables.equipment:
        pid, node = str(row.get('pipe')), str(row.get('editor_node', ''))
        pipe = pipes.get(pid)
        if (pipe and node and node in {str(pipe['in']), str(pipe['out'])}
                and row.get('editor_library') in allowed):
            equipment.setdefault((pid, node), []).append(row)
    # If several source issues share a spot, do not use one fitting twice.
    needed: dict[tuple[str, str], int] = {}
    for row in issues:
        key = (str(row.get('pipe_label', row.get('pipe'))), str(row.get('node_label', '')))
        needed[key] = needed.get(key, 0) + _count(row.get('n', 1))
    pending, resolved = [], []
    for row in issues:
        key = (str(row.get('pipe_label', row.get('pipe'))), str(row.get('node_label', '')))
        matches = equipment.get(key, [])
        if (key[1] and row.get('where') != '구간 내부 다중 접속'
                and sum(_count(e.get('count', 1)) for e in matches) >= needed[key]):
            resolved.append(dict(row, resolution='editor_fitting',
                equipment_labels=[str(e.get('label')) for e in matches],
                fitting_ids=[str(e['editor_library']) for e in matches]))
        else:
            pending.append(dict(row))
    return FittingReview(tuple(pending), tuple(resolved), unlocated)


def sync_fitting_review(tables: Any) -> None:
    """Refresh the cached count without discarding source issues or legacy gaps."""
    review = fitting_review(tables)
    source = getattr(tables, 'unresolved', None)
    if source is not None and 'kind_items' in source:
        tables.unresolved = dict(source, kind_unlocated_count=review.unlocated_count)
    meta = list(getattr(tables, 'meta', None) or [])
    if any(k == '부속 판정 불가' for k, _ in meta):
        tables.meta = [(k, str(review.count) if k == '부속 판정 불가' else v) for k, v in meta]


def remap_fitting_review(source: dict | None, *, node_labels: Mapping[str, str],
                         pipe_labels: Mapping[str, str]) -> dict:
    """Carry source identities plus shifted display labels through a merge."""
    out = deepcopy(source or {})
    for group in ('kind_items', 'length_items', 'applied'):
        for row in out.get(group, []):
            pid = str(row.get('pipe_label', row.get('pipe', '')))
            if pid in pipe_labels:
                row['pipe_label'] = pipe_labels[pid]
            nid = str(row.get('node_label', ''))
            if nid in node_labels:
                row['node_label'] = node_labels[nid]
    return out
