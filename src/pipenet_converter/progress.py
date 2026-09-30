"""Optional, read-only processing observations; never part of calculation state."""
from __future__ import annotations

from collections.abc import Callable, Iterator
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class ProgressEvent:
    """A detached preview in CAD millimeters, not a validated network result."""

    kind: str
    phase: str
    data: dict[str, Any] = field(default_factory=dict)


Reporter = Callable[[ProgressEvent], None]
_REPORTER: ContextVar[Reporter | None] = ContextVar("converter_reporter", default=None)


def current_reporter() -> Reporter | None:
    """Read the observer scoped to this request/worker, if explicitly enabled."""
    return _REPORTER.get()


@contextmanager
def reporting(reporter: Reporter | None) -> Iterator[None]:
    """Temporarily observe this context; restore it even on cancellation."""
    token = _REPORTER.set(reporter)
    try:
        yield
    finally:
        _REPORTER.reset(token)


def report(kind: str, phase: str, **data: Any) -> None:
    """Notify a presentation observer without allowing its errors to alter output."""
    callback = _REPORTER.get()
    if callback is not None:
        try:
            callback(ProgressEvent(kind, phase, data))
        except Exception:
            pass  # A disconnected/failed progress surface is not a calculation failure.


class WorldPreview:
    """Publish newly decoded primitives in bounded batches, without retaining CAD."""

    def __init__(self) -> None:
        self.offsets = {"segs": 0, "circles": 0, "arcs": 0}
        self.count = 0
        report("reset", "도형 해석", mode="world")

    def flush(self, world: Any) -> None:
        """Copy only new line/circle/arc coordinates; originals remain untouched."""
        batch = []
        for name in self.offsets:
            rows = getattr(world, name)
            end = len(rows)
            for index in range(self.offsets[name], end):
                row = rows[index]
                if name == "segs":
                    coords = [*row[2][:2], *row[3][:2]]
                elif name == "arcs":
                    angles = world.arc_ang[index] if index < len(world.arc_ang) else None
                    coords = [*row[2:5], *(angles or (0, 360))]
                else:
                    coords = list(row[2:5])
                batch.append({"layer": str(row[0]), "kind": name, "xy": coords})
                self.count += 1
                if len(batch) == 256:
                    report("world", "도형 해석", items=batch, count=self.count)
                    batch = []
            self.offsets[name] = end
        if batch:
            report("world", "도형 해석", items=batch, count=self.count)
