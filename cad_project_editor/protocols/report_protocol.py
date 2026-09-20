"""
protocols/report_protocol.py

K-Fire — PDF 보고서 생성기 프로토콜 정의.
구체적인 구현(infra/report_generator.py)과 호출 코드를 분리한다.
"""
from __future__ import annotations
from typing import Protocol, runtime_checkable


@runtime_checkable
class ReportGeneratorProtocol(Protocol):
    """PDF 보고서 생성기가 반드시 구현해야 하는 인터페이스."""

    def build(self) -> None:
        """
        보고서를 조립하고 PDF 파일로 저장한다.
        저장 경로는 생성자에서 `filename`으로 전달받는다.
        """
        ...


__all__ = ["ReportGeneratorProtocol"]
