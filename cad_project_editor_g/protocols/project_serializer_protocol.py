"""
ProjectSerializerProtocol — 프로젝트 저장/불러오기 교체 가능 인터페이스

직렬화 구현체(infra.project_serializer.ProjectSerializer)를
호출부와 분리하기 위한 Protocol 정의.
"""
from __future__ import annotations

from typing import Protocol, runtime_checkable


@runtime_checkable
class ProjectSerializerProtocol(Protocol):
    """프로젝트 직렬화기가 구현해야 하는 메서드 집합."""

    def to_dict(self) -> dict:
        ...

    def to_runtime_snapshot(self) -> dict:
        ...

    def to_json_serializable(self) -> dict:
        ...

    def from_dict(self, data: dict) -> None:
        ...

    def from_runtime_snapshot(self, snapshot: dict) -> None:
        ...

    def save_json(self, filepath: str) -> None:
        ...

    def load_json(self, filepath: str, strict: bool = False) -> dict:
        ...


__all__ = ["ProjectSerializerProtocol"]
