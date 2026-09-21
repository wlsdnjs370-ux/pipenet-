"""Preserve fitting origins across explicit head takeoff expansion.

No coordinate matching is performed. A vertical branch replacing a plan arm
does not add a port; a peeled head takeoff adds exactly one head connection.
"""
from __future__ import annotations

from collections.abc import Collection, Mapping
from dataclasses import dataclass


@dataclass(frozen=True)
class FittingOrigin:
    """Original graph vertex and independently added head port count."""

    source_vertex: int
    added_ports: int = 0


def fitting_origins(
    node_ref: Mapping[str, int],
    head_takeoffs: Mapping[str, str],
    active_nodes: Collection[str],
) -> dict[str, FittingOrigin]:
    """Resolve origins by recorded identities, never display/grid proximity.

    ``head_takeoffs`` maps a new pipe-line connection to its original head node.
    Original head identities remain untouched for nozzle/editor addressing.
    Missing references remain unknown; unrelated generated elbows inherit none.
    """
    active = set(active_nodes)
    result = {nid: FittingOrigin(vid) for nid, vid in node_ref.items() if nid in active}
    for connection, head in head_takeoffs.items():
        if connection in active and head in active and head in node_ref:
            result[connection] = FittingOrigin(node_ref[head], added_ports=1)
    return result
