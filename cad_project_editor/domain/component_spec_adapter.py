"""Compatibility adapter between legacy flat node meta and ComponentSpec meta.

This is intentionally a read-side bridge. It lets loaders accept a future
ComponentSpec-shaped payload while the current save format keeps writing the
legacy flat Node fields that existing .kfp files and tests rely on.
"""
from __future__ import annotations

import copy
from typing import Any


COMPONENT_SPEC_KEYS = ("component_spec", "componentSpec")

_SPEC_TYPE_KEYS = ("type_id", "kind", "type", "component_type", "componentType")
_SPEC_PAYLOAD_KEYS = ("fields", "data", "props", "properties", "attributes")
_SPEC_RESERVED_KEYS = {
    *COMPONENT_SPEC_KEYS,
    *_SPEC_PAYLOAD_KEYS,
    "schema",
    "schema_version",
    "schemaVersion",
    "version",
}


def normalize_node_meta_for_load(meta: dict[str, Any]) -> dict[str, Any]:
    """Return legacy-flat node meta, accepting an optional ComponentSpec payload.

    Precedence is deliberately conservative:
    - ComponentSpec fields fill in missing legacy-flat fields.
    - Existing flat fields win if both shapes are present.

    That keeps current 3.x flat files stable while allowing spec-only payloads
    to load through the old Node.update_from_meta path.
    """
    if not isinstance(meta, dict):
        return {}

    flat_meta = copy.deepcopy(meta)
    spec_meta = _flatten_component_spec(flat_meta)
    if not spec_meta:
        return flat_meta

    merged = spec_meta
    merged.update(flat_meta)
    return merged


def _flatten_component_spec(meta: dict[str, Any]) -> dict[str, Any]:
    spec = _first_dict(meta, COMPONENT_SPEC_KEYS)
    if not spec:
        return {}

    flat: dict[str, Any] = {}

    type_id = _first_value(spec, _SPEC_TYPE_KEYS)
    if type_id is not None and str(type_id).strip():
        flat["type_id"] = str(type_id).strip()

    category_id = spec.get("category_id")
    if category_id is not None:
        flat["category_id"] = category_id

    for key in _SPEC_PAYLOAD_KEYS:
        payload = spec.get(key)
        if isinstance(payload, dict):
            flat.update(copy.deepcopy(payload))

    # Accept shallow spec fields too. This lets a compact future shape such as
    # {"component_spec": {"kind": "pump", "rated_q": 1000}} load without a
    # nested fields/properties object.
    for key, value in spec.items():
        if key in _SPEC_RESERVED_KEYS or key in _SPEC_TYPE_KEYS:
            continue
        if key in flat:
            continue
        flat[key] = copy.deepcopy(value)

    return flat


def _first_dict(source: dict[str, Any], keys: tuple[str, ...]) -> dict[str, Any]:
    for key in keys:
        value = source.get(key)
        if isinstance(value, dict):
            return value
    return {}


def _first_value(source: dict[str, Any], keys: tuple[str, ...]) -> Any:
    for key in keys:
        if key in source:
            return source[key]
    return None
