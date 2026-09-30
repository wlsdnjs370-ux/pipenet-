"""Bounded, content-free timing records for Module F requests and workers."""
from __future__ import annotations

from collections import deque
from contextlib import contextmanager
from threading import Lock
from time import perf_counter
from typing import Iterator

_EVENTS: deque[dict] = deque(maxlen=400)
_LOCK = Lock()


@contextmanager
def measure(phase: str, *, sid: str = '', **counts: int | str) -> Iterator[dict]:
    """Measure elapsed time without traversing models or enabling memory tracing."""
    event = dict(phase=phase, sid=str(sid), **counts)
    started = perf_counter()
    try:
        yield event
    finally:
        event['elapsed_ms'] = round((perf_counter()-started)*1000, 3)
        with _LOCK:
            _EVENTS.append(event)


def report(sid: str) -> list[dict]:
    """Return only this session's recent timings, never drawing/user content."""
    with _LOCK:
        return [dict(e) for e in _EVENTS if e['sid'] == str(sid)]
