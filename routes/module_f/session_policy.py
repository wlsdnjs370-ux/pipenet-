"""Audited mutation boundaries; unknown routes retain the full-copy fallback."""
from __future__ import annotations

VIEW_FIELDS = frozenset({'net_rev','n_bodies','stat_rev','body_stat','head_rev',
                        'wet_rev','aj_sent','worst_rev'})
SELECTION_FIELDS = VIEW_FIELDS | frozenset({
    'edit','water_path','flow_report','worst','worst_edit','worst_cand','worst_zones',
    'worst_k','worst_edits','worst_args','design_settings','selection_mode',
})


def snapshot_fields(path: str, sess: dict) -> frozenset[str] | None:
    """None means conservative whole session; a set includes absent output keys."""
    if path in ('/api/module-f/edit/flow','/api/module-f/edit/worst',
                '/api/module-f/edit/network-mode'):
        return SELECTION_FIELDS
    if path == '/api/module-f/merge/mode':
        return frozenset({'supply_mode','source_drop_m','pump_spec'})
    return None


def background_snapshot_fields(path: str) -> frozenset[str] | None:
    """Audited job outputs: rollback snapshot, not a synchronous draft.

Sizing workers publish only these two keys. Their CAD/merged input is read-only.
Unrecognised jobs keep the existing conservative full-session snapshot.
"""
    if path in ('/api/module-f/merge/sizing/run','/api/module-f/merge/sizing/emit'):
        return frozenset({'sizing_report','sizing_files'})
    return None
