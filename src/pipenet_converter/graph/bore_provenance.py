"""Read-only bore evidence and edit audit records; never choose a pipe size.

All diameters here are nominal millimetres, not hydraulic inner diameters.
Unknown legacy/default sources must not be described as a code-table fallback.
"""
from __future__ import annotations

from copy import deepcopy
from typing import Any, Mapping

FIELD = "bore_provenance"
SOURCES = {"text": "drawing", "nfpc_fallback": "rule", "nfpc_min": "review", "review_default": "review"}
MANUAL = {"user", "사람", "직접 편집 · 라이브러리", "user_fix"}


def evidence(row: Mapping[str, Any]) -> dict[str, Any]:
    """Copy stored evidence, or explicitly identify the limits of a legacy row."""
    saved = row.get(FIELD)
    if isinstance(saved, dict):
        return deepcopy(saved)
    source = row.get("dia_src") or row.get("dia_source") or row.get("src")
    result = {"version": 1, "source": source if source in SOURCES else "unknown",
              "auto_mm": row.get("dia") if source in SOURCES else None}
    if source in MANUAL:
        result["manual"] = {"value_mm": row.get("dia"), "note": "이전 직접 입력 · 원근거 미기록"}
    return result


def record_manual(row: dict[str, Any], value_mm: int | float,
                  note: str = "", *, previous_mm: int | float | None = None) -> None:
    """Retain the first automatic evidence while recording the latest override."""
    rec = evidence(row)
    # A subsequent typed/library override no longer derives from that text.
    rec.pop('reference_annotation', None)
    rec.pop('serial_manual', None)
    rec["manual"] = {"previous_mm": row.get("dia") if previous_mm is None else previous_mm,
                     "value_mm": value_mm, "note": str(note or "수정 메모 없음")}
    rec["block_export"] = False
    row[FIELD] = rec


def record_adjustment(row: dict[str, Any], previous_mm: int | float,
                      reason: str) -> None:
    """Annotate an existing automatic adjustment without changing its result."""
    if row.get("dia") == previous_mm:
        return
    before = dict(row, dia=previous_mm)
    rec = evidence(before)
    rec["adjustment"] = {"previous_mm": previous_mm, "value_mm": row.get("dia"),
                         "reason": reason}
    row[FIELD] = rec


def describe_bore(row: Mapping[str, Any]) -> dict[str, Any]:
    """Return a JSON-safe property-card record without mutating a table row."""
    rec = evidence(row)
    source = rec.get("source", "unknown")
    category = rec.get("serial_category") or SOURCES.get(source, "unknown")
    manual = rec.get("manual")
    adjustment = rec.get("adjustment")
    # A stale/legacy provenance field cannot certify a subsequently changed DN.
    expected = (manual or adjustment or {}).get("value_mm", rec.get("auto_mm"))
    untracked = expected is not None and row.get("dia") != expected
    if adjustment or untracked or rec.get("review_reasons") or rec.get("inferred") or rec.get("block_export"):
        category = "review"
    return {**rec, "category": category, "current_mm": row.get("dia"),
            "edited": bool(manual or rec.get('serial_edited')), "review": category == "review",
            "untracked_change": untracked,
            "method": rec.get("method") or ("최근접 문자 매칭" if source in ("text", "nfpc_min") else None)}
