"""Bounded, operation-scoped H previews; no access to mutable session drafts."""
from __future__ import annotations

from collections import deque
from dataclasses import asdict, dataclass, field
from threading import Lock
import re
import time
from urllib.parse import urlsplit

from flask import Flask, g, jsonify, request
from src.pipenet_converter.progress import ProgressEvent, reporting


@dataclass
class PreviewChannel:
    """Small replay buffer. Dropped events are explicitly marked, never concealed."""

    seq: int = 0
    touched: float = field(default_factory=time.monotonic)
    events: deque = field(default_factory=lambda: deque(maxlen=192))
    diameters: dict = field(default_factory=dict)
    annotations: list | None = None


class PreviewStore:
    """Thread-safe storage of detached observations, independent of F sessions."""

    def __init__(self) -> None:
        self.channels: dict[str, PreviewChannel] = {}
        self.lock = Lock()

    def _get(self, operation: str) -> PreviewChannel:
        now = time.monotonic()
        for key in list(self.channels):
            if now - self.channels[key].touched > 600:
                del self.channels[key]
        if operation not in self.channels and len(self.channels) >= 12:
            del self.channels[min(self.channels, key=lambda k: self.channels[k].touched)]
        channel = self.channels.setdefault(operation, PreviewChannel())
        channel.touched = now
        return channel

    def publish(self, operation: str, event: ProgressEvent) -> None:
        """Keep only serialized preview values, not references to engine objects."""
        item = asdict(event)
        with self.lock:
            channel = self._get(operation)
            channel.seq += 1
            channel.events.append(dict(item, seq=channel.seq))
            if event.kind == 'diameter':
                data = item['data']
                if data.get('reset'):
                    channel.diameters.clear()
                    channel.annotations = data.get('annotations', [])
                if data.get('edge') and (len(channel.diameters) < 50000 or tuple(data['edge']) in channel.diameters):
                    channel.diameters[tuple(data['edge'])] = data

    def read(self, operation: str, after: int) -> dict:
        """Read at most 48 real events per response, with a detectable replay gap."""
        with self.lock:
            channel = self._get(operation)
            rows = [event for event in channel.events if event["seq"] > after][:48]
            gap = bool(rows and rows[0]["seq"] > after + 1)
            snapshot = list(channel.diameters.values()) if gap else []
            return {"ok": True, "events": rows, "diameter_snapshot": snapshot,
                    "annotation_snapshot": channel.annotations if gap else None,
                    "snapshot_cursor": channel.seq if gap else None,
                    "cursor": rows[-1]["seq"] if rows else after,
                    "gap": gap}


def install(app: Flask) -> None:
    """Observe only requests originating from H with an existing operation token."""
    store = PreviewStore()
    app.extensions["module_h_preview"] = store

    @app.before_request
    def h_observe_request() -> None:
        operation = request.headers.get("X-Module-F-Operation", "")
        if (request.path.startswith("/api/module-f/")
                and urlsplit(request.referrer or "").path == "/module-h"
                and re.fullmatch(r"[a-zA-Z0-9_-]{16,80}", operation)):
            scope = reporting(lambda event: store.publish(operation, event))
            scope.__enter__()
            g.h_progress_scope = scope

    @app.teardown_request
    def h_release_observer(error: BaseException | None) -> None:
        scope = g.pop("h_progress_scope", None)
        if scope is not None:
            scope.__exit__(None, None, None)

    @app.get("/api/module-h/progress")
    def h_progress():
        operation = request.args.get("operation", "")
        if not re.fullmatch(r"[a-zA-Z0-9_-]{16,80}", operation):
            return jsonify(ok=False, message="작업 번호가 올바르지 않습니다."), 400
        try:
            after = max(0, int(request.args.get("after", "0")))
        except ValueError:
            return jsonify(ok=False, message="진행 위치가 올바르지 않습니다."), 400
        response = jsonify(store.read(operation, after))
        response.headers["Cache-Control"] = "no-store"
        return response
