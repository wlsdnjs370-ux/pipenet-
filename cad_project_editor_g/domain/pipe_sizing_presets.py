"""Head-count diameter presets for the unified pipe-sizing dialog.

The user-editable catalog lives in ``pipe_sizing_library.json`` (factory seed
at the workspace root, per-user copy under %LOCALAPPDATA%/K-Fire, loaded by
LibraryService).  This module is the pure-domain side of that file: schema
constants, dict <-> dataclass conversion, and editing validation.  The
built-in catalog below is the last-resort fallback (file missing/corrupt) and
the factory-seed reference; a test pins the seed file to it so they cannot
drift.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from typing import Any, Mapping, Sequence

from domain.pipe_sizing import (
    HEAD_RANGE_FIXED_NOMINALS,
    OPEN_ENDED_HEAD_UNTIL_UI,
    RangeRow,
    ranges_to_fixed_head_slots,
)

PIPE_SIZING_LIBRARY_FILENAME = "pipe_sizing_library.json"
PIPE_SIZING_LIBRARY_SCHEMA_VERSION = "1.0"
# Forward-compat: presets carry a kind so other table types can join later.
HEAD_COUNT_PRESET_KIND = "head_count"


@dataclass(frozen=True)
class HeadRangePreset:
    """One reusable head-count → DN table."""

    preset_id: str
    # System presets: locale key / display label source string (Korean SSOT
    # for _t).  User presets: the raw name typed by the user (shown as-is).
    label: str
    head_ranges: tuple[RangeRow, ...]
    system: bool = False

    def slots(
        self,
        *,
        nominals: Sequence[int] = HEAD_RANGE_FIXED_NOMINALS,
    ) -> list[tuple[int | None, int]]:
        return ranges_to_fixed_head_slots(self.head_ranges, nominals=nominals)


# Built-in catalog — fallback + factory-seed reference. Library rows mirror
# this shape (preset_id, display_name, head_ranges[{start,end,nominal_mm}]).
_BUILTIN_HEAD_RANGE_PRESETS: tuple[HeadRangePreset, ...] = (
    HeadRangePreset(
        preset_id="nfsc_k80_standard",
        label="NFSC K80 Standard",
        system=True,
        head_ranges=(
            RangeRow(1.0, 2.0, 25),
            RangeRow(3.0, 3.0, 32),
            RangeRow(4.0, 5.0, 40),
            RangeRow(6.0, 10.0, 50),
            RangeRow(11.0, 30.0, 65),
            RangeRow(31.0, 60.0, 80),
            RangeRow(61.0, 100.0, 100),
            RangeRow(101.0, 160.0, 125),
            # DN150 covers 161+ (open-ended in domain; UI uses sentinel until).
            RangeRow(161.0, None, 150),
        ),
    ),
)


def builtin_head_range_presets() -> tuple[HeadRangePreset, ...]:
    return _BUILTIN_HEAD_RANGE_PRESETS


def list_head_range_presets() -> Sequence[HeadRangePreset]:
    """Built-in fallback catalog (callers without a loaded library)."""
    return _BUILTIN_HEAD_RANGE_PRESETS


def get_head_range_preset(preset_id: str) -> HeadRangePreset | None:
    wanted = str(preset_id or "").strip()
    if not wanted:
        return None
    for preset in _BUILTIN_HEAD_RANGE_PRESETS:
        if preset.preset_id == wanted:
            return preset
    return None


# ---------------------------------------------------------------------------
# Library dict <-> presets
# ---------------------------------------------------------------------------

def _coerce_rows(raw_rows: Any) -> tuple[RangeRow, ...]:
    """Coerce JSON rows tolerantly; invalid rows are dropped, never fatal."""
    if not isinstance(raw_rows, Sequence) or isinstance(raw_rows, (str, bytes)):
        return ()
    rows: list[RangeRow] = []
    for raw in raw_rows:
        row = RangeRow.from_dict(raw)
        if row._input_errors or not math.isfinite(row.start):
            continue
        if row.end is not None and not math.isfinite(row.end):
            continue
        if int(row.nominal_mm) <= 0:
            continue
        rows.append(row)
    return tuple(rows)


def presets_from_library(data: Mapping[str, Any] | None) -> list[HeadRangePreset]:
    """Parse the library dict tolerantly: broken entries are skipped.

    Unknown ``kind`` values are ignored (forward compatibility), duplicate
    preset ids keep the first occurrence.
    """
    if not isinstance(data, Mapping):
        return []
    raw_presets = data.get("presets")
    if not isinstance(raw_presets, list):
        return []
    result: list[HeadRangePreset] = []
    seen_ids: set[str] = set()
    for index, entry in enumerate(raw_presets):
        if not isinstance(entry, Mapping):
            continue
        kind = str(entry.get("kind") or HEAD_COUNT_PRESET_KIND).strip()
        if kind != HEAD_COUNT_PRESET_KIND:
            continue
        label = str(entry.get("display_name") or "").strip()
        if not label:
            continue
        rows = _coerce_rows(entry.get("head_ranges"))
        if not rows:
            continue
        preset_id = str(entry.get("preset_id") or "").strip() or f"preset_{index + 1}"
        if preset_id in seen_ids:
            continue
        seen_ids.add(preset_id)
        result.append(
            HeadRangePreset(
                preset_id=preset_id,
                label=label,
                head_ranges=rows,
                system=bool(entry.get("system")),
            )
        )
    return result


def presets_to_library(presets: Sequence[HeadRangePreset]) -> dict[str, Any]:
    return {
        "schema_version": PIPE_SIZING_LIBRARY_SCHEMA_VERSION,
        "presets": [
            {
                "preset_id": preset.preset_id,
                "kind": HEAD_COUNT_PRESET_KIND,
                "display_name": preset.label,
                "system": bool(preset.system),
                "head_ranges": [row.to_dict() for row in preset.head_ranges],
            }
            for preset in presets
        ],
    }


def builtin_library_dict() -> dict[str, Any]:
    """Factory-seed shaped dict of the built-in catalog."""
    return presets_to_library(_BUILTIN_HEAD_RANGE_PRESETS)


def normalize_open_ended_tail(rows: Sequence[RangeRow]) -> list[RangeRow]:
    """UI sentinel (9999) on the last row means 'and above' → open end.

    Keeps the stored library round-trip clean: ranges_to_fixed_head_slots
    renders an open end back as the same sentinel.
    """
    result = list(rows)
    if result:
        last = result[-1]
        if (
            last.end is not None
            and math.isfinite(last.end)
            and int(round(float(last.end))) >= OPEN_ENDED_HEAD_UNTIL_UI
        ):
            result[-1] = RangeRow(last.start, None, last.nominal_mm)
    return result


# ---------------------------------------------------------------------------
# Editing validation — issue codes for the library editor dialog. Mirrors the
# progression rules PipeSizingDialog.get_head_ranges() enforces on its slots.
# ---------------------------------------------------------------------------

SLOT_ISSUE_EMPTY = "empty"
SLOT_ISSUE_NOT_NUMERIC = "not_numeric"
SLOT_ISSUE_BELOW_ONE = "below_one"
SLOT_ISSUE_DECREASING = "decreasing"


def head_slot_text_issues(texts: Sequence[str]) -> list[tuple[int, str]]:
    """Validate cumulative 'until' texts over the fixed DN ladder.

    Returns (row_index, issue_code) pairs; row_index -1 flags the whole
    table (no ranges entered). Empty cells mean 'DN unused' and are fine.
    """
    issues: list[tuple[int, str]] = []
    prev_until = 0
    any_value = False
    for index, raw in enumerate(texts):
        text = str(raw or "").strip()
        if not text:
            continue
        try:
            until = int(round(float(text)))
        except (TypeError, ValueError, OverflowError):
            issues.append((index, SLOT_ISSUE_NOT_NUMERIC))
            continue
        any_value = True
        if until < 1:
            issues.append((index, SLOT_ISSUE_BELOW_ONE))
            continue
        if until < prev_until:
            issues.append((index, SLOT_ISSUE_DECREASING))
        prev_until = max(prev_until, until)
    if not any_value and not issues:
        issues.append((-1, SLOT_ISSUE_EMPTY))
    return issues
