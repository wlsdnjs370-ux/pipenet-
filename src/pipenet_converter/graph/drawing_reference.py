"""Explicit native-text references; nominal mm never replaces library inner bore."""
from __future__ import annotations

from dataclasses import asdict
from typing import Any

from .bore_provenance import record_manual
from .diameter_inference import DiameterAnnotation


def record_drawing_reference(row: dict[str, Any], annotation: DiameterAnnotation) -> None:
    """Retain automatic evidence and attach the explicitly chosen original text."""
    if not annotation.identity or not annotation.raw_text.strip():
        raise ValueError('원본 관경 텍스트가 확인된 표기만 참조할 수 있습니다.')
    record_manual(row, annotation.nominal_mm, '원본 도면 참조 내경 변경')
    row['bore_provenance']['reference_annotation'] = asdict(annotation)


def annotation_from_snapshot(snapshot: dict[str, Any]) -> DiameterAnnotation:
    """Validate a persisted command's immutable native-text snapshot for replay."""
    try:
        annotation = DiameterAnnotation(**snapshot)
        if not annotation.identity or not annotation.raw_text.strip():
            raise ValueError
        if int(annotation.nominal_mm) != annotation.nominal_mm:
            raise ValueError
        return annotation
    except (TypeError, ValueError, AttributeError, OverflowError):
        raise ValueError('참조 관경 텍스트 정보가 올바르지 않습니다.') from None
