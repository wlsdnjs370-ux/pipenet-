# -*- coding: utf-8 -*-
"""설계 세션 저장소 — 지시서 §C.2 / §C.3.

설계 도서는 날인 문서다. 어떤 값이 어디서 왔는지 재구성할 수 없으면 감리 대응이
불가능하므로, 산출물은 덮어쓰지 않고(`constraints.json` 은 불변) 사람이 만진 것은
전부 `audit.ndjson` 에 남긴다.

동시 편집은 두 방식으로 막는다 — GATE 편집은 **낙관적 잠금**(`if_version` 불일치
시 409), C2 이후 실행은 **단계별 뮤텍스**로 직렬화한다. 뮤텍스는 프로세스 안에서만
유효하다(서버가 waitress 단일 프로세스·다중 스레드라 충분하다). 워커를 여러
프로세스로 늘리면 파일 잠금으로 바꿔야 한다.
"""
from __future__ import annotations

import json
import os
import re
import threading
import uuid
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any

SESSION_TTL_DAYS = 30
KST = timezone(timedelta(hours=9))

# UUID4 정규형만 받는다. 세션 ID 가 그대로 디렉터리 이름이 되므로 여기가 유일한
# 경로 조작 방어선이다.
_SESSION_ID_RE = re.compile(r"^[0-9a-f]{8}-[0-9a-f]{4}-4[0-9a-f]{3}-[89ab][0-9a-f]{3}-[0-9a-f]{12}$")

_STAGE_LOCKS: dict[tuple[str, str], threading.Lock] = {}
_STAGE_LOCKS_GUARD = threading.Lock()


class SessionNotFound(Exception):
    pass


class ArtifactNotFound(Exception):
    def __init__(self, name: str):
        self.name = name
        super().__init__(f"{name} 이(가) 아직 없습니다")


class VersionConflict(Exception):
    """다른 탭이 먼저 저장했다. 현재 내용을 함께 돌려줘 사용자가 판단하게 한다."""

    def __init__(self, name: str, expected: int, current: int, data: dict):
        self.name = name
        self.expected = expected
        self.current = current
        self.data = data
        super().__init__(f"{name} 버전 불일치 — 보낸 값 {expected}, 현재 {current}")


class ImmutableArtifact(Exception):
    """`constraints.json` 은 한 번 쓰면 불변이다. 어떤 제약으로 설계했는지가 감사 대상."""


class StageBusy(Exception):
    def __init__(self, stage: str):
        self.stage = stage
        super().__init__(f"{stage} 단계가 이미 실행 중입니다")


def now_iso() -> str:
    return datetime.now(KST).isoformat(timespec="seconds")


def is_session_id(sid: str) -> bool:
    return bool(_SESSION_ID_RE.match(sid or ""))


class DesignSession:
    """`<root>/<session_id>/` 한 벌. 파일별 `version` 으로 낙관적 잠금을 건다."""

    IMMUTABLE = ("constraints.json",)

    def __init__(self, root: Path, sid: str):
        if not is_session_id(sid):
            raise SessionNotFound(sid)
        self.sid = sid
        self.dir = Path(root) / sid

    # ── 생성·열기 ────────────────────────────────────────────────────────
    @classmethod
    def create(cls, root: Path, *, operator: str | None = None) -> "DesignSession":
        sid = str(uuid.uuid4())
        sess = cls(root, sid)
        sess.dir.mkdir(parents=True, exist_ok=False)
        created = now_iso()
        sess.write("meta.json", {
            "schema": "fncadnet.design_session/1",
            "session_id": sid,
            "created_at": created,
            "expires_at": (datetime.now(KST) + timedelta(days=SESSION_TTL_DAYS)).isoformat(timespec="seconds"),
            "operator": operator,
            "stage": "c1",
            "gate_passed": False,
        })
        sess.audit("system", "session", "created", {"operator": operator})
        return sess

    @classmethod
    def open(cls, root: Path, sid: str) -> "DesignSession":
        sess = cls(root, sid)
        if not (sess.dir / "meta.json").is_file():
            raise SessionNotFound(sid)
        return sess

    # ── 파일 입출력 ──────────────────────────────────────────────────────
    def path(self, name: str) -> Path:
        return self.dir / name

    def read(self, name: str) -> tuple[dict, int]:
        """반환 `(내용, version)`. 없으면 `ArtifactNotFound`."""
        try:
            raw = self.path(name).read_text(encoding="utf-8")
        except FileNotFoundError:
            raise ArtifactNotFound(name) from None
        data = json.loads(raw)
        return data, int(data.get("version") or 0)

    def read_or_none(self, name: str) -> tuple[dict | None, int]:
        try:
            return self.read(name)
        except ArtifactNotFound:
            return None, 0

    def write(self, name: str, data: dict, *, if_version: int | None = None) -> int:
        """원자적 쓰기 + 버전 증가. 반환은 새 version.

        `if_version` 을 주면 현재 버전과 다를 때 `VersionConflict` 를 던진다.
        생략하면 검사 없이 덮어쓴다(생성 시점·서버 자신의 산출물).
        """
        current_data, current = self.read_or_none(name)
        if name in self.IMMUTABLE and current_data is not None:
            raise ImmutableArtifact(name)
        if if_version is not None and if_version != current:
            raise VersionConflict(name, if_version, current, current_data or {})

        payload = dict(data)
        payload["version"] = current + 1
        payload["updated_at"] = now_iso()
        dest = self.path(name)
        tmp = dest.with_suffix(f".{uuid.uuid4().hex[:12]}.tmp")
        tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, dest)
        return payload["version"]

    def update_meta(self, **fields: Any) -> None:
        meta, _ = self.read("meta.json")
        meta.update(fields)
        self.write("meta.json", meta)

    # ── 감사 로그 (append-only) ──────────────────────────────────────────
    def audit(self, actor: str, stage: str, event: str, detail: dict | None = None) -> None:
        line = json.dumps({
            "ts": now_iso(), "actor": actor, "stage": stage, "event": event,
            "detail": detail or {},
        }, ensure_ascii=False)
        self.dir.mkdir(parents=True, exist_ok=True)
        with self.path("audit.ndjson").open("a", encoding="utf-8") as fp:
            fp.write(line + "\n")

    def audit_entries(self) -> list[dict]:
        try:
            raw = self.path("audit.ndjson").read_text(encoding="utf-8")
        except FileNotFoundError:
            return []
        return [json.loads(line) for line in raw.splitlines() if line.strip()]

    # ── 단계 뮤텍스 ──────────────────────────────────────────────────────
    def stage_lock(self, stage: str) -> "_StageLock":
        with _STAGE_LOCKS_GUARD:
            lock = _STAGE_LOCKS.setdefault((self.sid, stage), threading.Lock())
        return _StageLock(lock, stage)

    # ── 상태 요약 ────────────────────────────────────────────────────────
    def status(self) -> dict[str, Any]:
        meta, _ = self.read("meta.json")
        artifacts = {}
        for name in ("building.json", "constraints.json", "design.json",
                     "graph.json", "checks.json"):
            data, version = self.read_or_none(name)
            artifacts[name] = None if data is None else {
                "version": version, "updated_at": data.get("updated_at")}
        timings, _ = self.read_or_none("timings.json")
        return {"meta": meta, "artifacts": artifacts, "timings": timings or {}}


class _StageLock:
    def __init__(self, lock: threading.Lock, stage: str):
        self._lock = lock
        self._stage = stage

    def __enter__(self) -> None:
        if not self._lock.acquire(blocking=False):
            raise StageBusy(self._stage)

    def __exit__(self, *exc: Any) -> None:
        self._lock.release()
