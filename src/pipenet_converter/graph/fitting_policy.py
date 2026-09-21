"""Module F fitting policy; connectivity and separate equipment are unchanged."""
from collections.abc import Iterable, Mapping
from typing import Any

CALCULATION_FITTINGS = frozenset({"tee", "elbow", "elbow-45"})
STRAIGHT_TEE_KINDS = frozenset({"tee-run", "cross-run"})


def without_straight_tees(rows: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
    """Copy native fitting rows without legacy zero-loss straight-through tees.

    Do not simplify nodes/edges or touch separately stored valve/equipment loss.
    Unknown types remain available to validation, never silently become elbows.
    """
    return [dict(row) for row in rows if row.get("type") not in STRAIGHT_TEE_KINDS]
