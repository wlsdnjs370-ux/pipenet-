"""Bounded, disposable replay checkpoints; JSON commands remain authoritative.

Each checkpoint owns a private graph copy. Callers receive a copy, never the
stored graph. The immutable cache can therefore be shared by session snapshots.
At most two graph checkpoints are retained, not one snapshot per edit.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import dataclass, field
import hashlib
import json

from core.network_editor import Network, apply_edit
from routes.module_f.performance import measure

CHECKPOINT_INTERVAL = 10
MAX_CHECKPOINTS = 2


def _prefix(commands: list[dict], cursor: int) -> str:
    return hashlib.sha256(json.dumps(commands[:cursor], sort_keys=True,
                                     ensure_ascii=False).encode()).hexdigest()


@dataclass(frozen=True)
class _Checkpoint:
    cursor: int
    digest: str
    _graph: Network = field(repr=False, compare=False)


@dataclass(frozen=True)
class ReplayCache:
    """Functional cache: adding/pruning returns a new cache, never mutates one."""

    basis: tuple[str, str | None]
    _entries: tuple[_Checkpoint, ...] = ()

    def __deepcopy__(self, memo: dict) -> ReplayCache:
        memo[id(self)] = self
        return self

    @property
    def count(self) -> int:
        """Number of retained graph checkpoints, for diagnostics/tests."""
        return len(self._entries)

    def pruned(self, commands: list[dict]) -> ReplayCache:
        return ReplayCache(self.basis, tuple(p for p in self._entries
            if p.cursor <= len(commands) and p.digest == _prefix(commands, p.cursor)))

    def remember(self, commands: list[dict], cursor: int, graph: Network) -> ReplayCache:
        """Store periodic copies only; never retain discarded redo branches."""
        cache = self.pruned(commands)
        if not cursor or cursor % CHECKPOINT_INTERVAL or any(p.cursor == cursor for p in cache._entries):
            return cache
        point = _Checkpoint(cursor, _prefix(commands, cursor), deepcopy(graph))
        return ReplayCache(self.basis, (*cache._entries, point)[-MAX_CHECKPOINTS:])

    def start(self, commands: list[dict], cursor: int, base: Network) -> tuple[int, Network]:
        """Leave at least the final command for replay to preserve selection data."""
        points = [p for p in self.pruned(commands)._entries if p.cursor < cursor]
        point = max(points, key=lambda p: p.cursor, default=None)
        return (point.cursor, deepcopy(point._graph)) if point else (0, deepcopy(base))


def cache_for(editor: dict) -> ReplayCache:
    """A rebuilt base or changed library invalidates all cached graphs."""
    basis = (editor['base_hash'], editor.get('library'))
    cache = editor.get('_replay_cache')
    return cache if isinstance(cache, ReplayCache) and cache.basis == basis else ReplayCache(basis)


def remember(editor: dict) -> None:
    """Attach a new cache to a candidate editor, without changing the old editor."""
    editor['_replay_cache'] = cache_for(editor).remember(
        editor['commands'], editor['cursor'], editor['current'])


def replay(editor: dict, commands: list[dict], cursor: int, library: dict,
           *, sid: str = '', warm: bool = False) -> tuple[Network, dict | None]:
    """Rebuild from the closest valid checkpoint; optionally warm a fresh load."""
    cache = cache_for(editor)
    start, candidate = cache.start(commands, cursor, editor['base'])
    selection = None
    with measure('history_replay', sid=sid, commands=cursor, replayed=cursor-start):
        for index in range(start, cursor):
            candidate, selection = apply_edit(candidate, commands[index], library)
            if warm:
                cache = cache.remember(commands, index+1, candidate)
    if warm:
        editor['_replay_cache'] = cache
    return candidate, selection
