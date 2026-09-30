"""Isolated, opt-in drafts for audited in-memory Module F operations.

Unowned state is read-only by contract. File-producing routes deliberately stay
on the legacy path until their persistence has a corresponding transaction.
"""
from __future__ import annotations

from routes.module_f.performance import measure


class SessionConflict(RuntimeError):
    """A concurrent request or changed basis prevents accepting a result."""


class SessionWorkspace:
    """Copy only owned fields and publish them after successful validation."""

    def __init__(self, live: dict, fields: frozenset[str]):
        from routes.module_f.cancellation import _session_copy
        self.live = live
        self.fields = fields
        self.basis = (live.get('_state_revision', 0), live.get('active'), live.get('key'))
        self.original = dict(live)
        self.working = dict(live)
        self.working.update(_session_copy(live, fields))

    def validate(self) -> None:
        """Reject stale results and undeclared top-level mutations."""
        now = (self.live.get('_state_revision', 0), self.live.get('active'), self.live.get('key'))
        if now != self.basis:
            raise SessionConflict('도면 상태가 바뀌었습니다. 최신 상태에서 다시 실행하세요.')
        runtime = {'job', 'log', 'touched'}
        absent = object()
        for key in (set(self.original) | set(self.working)) - self.fields - runtime:
            if self.original.get(key, absent) is not self.working.get(key, absent):
                raise SessionConflict(f'작업 범위 밖의 상태 변경을 차단했습니다: {key}')

    def commit(self) -> None:
        """Publish under the operation's lock; never merge partial results."""
        self.validate()
        with measure('session_publish', sid=self.live.get('id', ''), keys=len(self.fields)):
            self.live.update({k: self.working[k] for k in self.fields if k in self.working})
            for key in self.fields - self.working.keys():
                self.live.pop(key, None)

