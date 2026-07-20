"""
change_log.py — 변경이력 (Change Log) + 3중 Trace 기록

Per the v4 pipeline §⑧, every redesign attempt and every automated decision
must carry a 3-source citation (NFTC + HB + PhD) and a 4-tuple log entry
(diagnosis · option · result · timestamp). This module provides the data
structures and the file-system journal.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from nftc_rules import TripleTrace


# ---------------------------------------------------------------------------
# 1. ChangeLog entry
# ---------------------------------------------------------------------------


@dataclass
class ChangeLogEntry:
    """A single redesign attempt entry in the change log."""

    version: str                         # "v1", "v2", ...
    timestamp: str                       # ISO-8601 UTC
    triggered_by: str                    # "auto" / "human" / "system"
    diagnosis: dict[str, Any]            # {zone, metric, value, ...}
    option: str                          # "⑦A_diameter" / "⑦B_routing" / "⑦E_loop" / "⑦F_pump_change" / ...
    parameters: dict[str, Any]           # parameter changes applied
    kpi_before: dict[str, Any]           # 6대 + 3대 imbalance metrics
    kpi_after: dict[str, Any]
    verdict: str                         # "PASS" / "FAIL" / "REVIEW" / "SIGN-OFF"
    trace_links: list[str] = field(default_factory=list)  # references to NFTC/HB/PhD clauses
    note: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


# ---------------------------------------------------------------------------
# 2. ChangeLogger
# ---------------------------------------------------------------------------


class ChangeLogger:
    """Append-only change log writer.

    The log is stored as a JSON Lines file under `data/change_logs/<project_id>.jsonl`.
    Read access is provided via `iter_entries`.
    """

    def __init__(self, project_id: str, *, log_dir: Path | str = "data/change_logs") -> None:
        self.project_id = project_id
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = self.log_dir / f"{project_id}.jsonl"
        self._version_counter = self._init_version_counter()

    def _init_version_counter(self) -> int:
        """Resume version counter from existing log if present."""
        if not self.log_path.exists():
            return 0
        try:
            with self.log_path.open(encoding="utf-8") as f:
                count = sum(1 for _ in f)
            return count
        except OSError:
            return 0

    def append(
        self,
        *,
        diagnosis: dict[str, Any],
        option: str,
        parameters: dict[str, Any],
        kpi_before: dict[str, Any],
        kpi_after: dict[str, Any],
        verdict: str,
        triggered_by: str = "auto",
        trace_links: list[str] | None = None,
        note: str = "",
    ) -> ChangeLogEntry:
        """Append a new ChangeLogEntry."""
        self._version_counter += 1
        entry = ChangeLogEntry(
            version=f"v{self._version_counter}",
            timestamp=datetime.now(timezone.utc).isoformat(),
            triggered_by=triggered_by,
            diagnosis=diagnosis,
            option=option,
            parameters=parameters,
            kpi_before=kpi_before,
            kpi_after=kpi_after,
            verdict=verdict,
            trace_links=trace_links or [],
            note=note,
        )
        with self.log_path.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry.to_dict(), ensure_ascii=False) + "\n")
        return entry

    def iter_entries(self):
        """Iterate change log entries (oldest → newest)."""
        if not self.log_path.exists():
            return
        with self.log_path.open(encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    yield json.loads(line)
                except json.JSONDecodeError:
                    continue

    def latest(self, n: int = 10) -> list[dict[str, Any]]:
        entries = list(self.iter_entries())
        return entries[-n:]

    def render_table(self) -> list[dict[str, Any]]:
        """Render entries as table rows for HTML display."""
        rows: list[dict[str, Any]] = []
        for ent in self.iter_entries():
            rows.append({
                "version": ent.get("version"),
                "timestamp": ent.get("timestamp"),
                "diagnosis": _summarize_diagnosis(ent.get("diagnosis", {})),
                "option": ent.get("option"),
                "verdict": ent.get("verdict"),
                "triggered_by": ent.get("triggered_by"),
                "kpi_delta": _compute_kpi_delta(
                    ent.get("kpi_before", {}),
                    ent.get("kpi_after", {}),
                ),
            })
        return rows


def _summarize_diagnosis(d: dict[str, Any]) -> str:
    """One-line diagnosis summary for table rendering."""
    if not d:
        return "—"
    parts: list[str] = []
    for key in ("zone", "metric", "value"):
        if key in d:
            parts.append(f"{key}={d[key]}")
    other = [k for k in d if k not in {"zone", "metric", "value"}]
    if other:
        parts.append(f"+{len(other)} fields")
    return " ".join(parts) or "—"


def _compute_kpi_delta(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    """Compute numeric KPI deltas between before/after."""
    delta: dict[str, Any] = {}
    for k in set(before) | set(after):
        b = before.get(k)
        a = after.get(k)
        try:
            if isinstance(b, (int, float)) and isinstance(a, (int, float)):
                delta[k] = round(a - b, 4)
        except Exception:
            pass
    return delta


# ---------------------------------------------------------------------------
# 3. TripleTracer — running collector for all decisions
# ---------------------------------------------------------------------------


class TripleTracer:
    """Collect every TripleTrace produced during a pipeline run.

    Each rule decision returns a TripleTrace; this collector preserves all of
    them so the final report can answer "which clause produced which value?"
    """

    def __init__(self) -> None:
        self._traces: list[dict[str, Any]] = []

    def record(self, *, decision_key: str, trace: TripleTrace, value: Any = None) -> None:
        self._traces.append({
            "key": decision_key,
            "value": value,
            **trace.to_dict(),
            "timestamp": datetime.now(timezone.utc).isoformat(),
        })

    def all(self) -> list[dict[str, Any]]:
        return list(self._traces)

    def by_source(self) -> dict[str, list[dict[str, Any]]]:
        """Group traces by source (NFTC / HB / PhD)."""
        grouped: dict[str, list[dict[str, Any]]] = {"NFTC": [], "HB": [], "PhD": []}
        for t in self._traces:
            if t.get("NFTC"):
                grouped["NFTC"].append(t)
            if t.get("HB"):
                grouped["HB"].append(t)
            if t.get("PhD"):
                grouped["PhD"].append(t)
        return grouped

    def summary(self) -> dict[str, int]:
        grouped = self.by_source()
        return {k: len(v) for k, v in grouped.items()}


__all__ = [
    "ChangeLogEntry",
    "ChangeLogger",
    "TripleTracer",
]
