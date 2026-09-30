"""Reject unresolved competing diameter interpretations before file generation."""
from __future__ import annotations

from collections.abc import Iterable, Mapping
from typing import Any


def require_resolved_diameters(rows: Iterable[Mapping[str, Any]]) -> None:
    """Allow legacy/fallback rows, but do not silently export a known conflict."""
    rows = list(rows)
    unknown = [str(row.get('label', '?')) for row in rows
               if (row.get('bore_provenance') or {}).get('policy') == 'drawing_first_v1'
               and not row.get('dia')]
    if unknown:
        raise ValueError(f"관경 미지정 {len(unknown)}개: {', '.join(unknown[:12])}. "
                         "도면 근거·일괄 수정 또는 통합망 역산으로 관경을 정의한 뒤 출력하세요.")
    blocked = [str(row.get('label', '?')) for row in rows
               if (row.get('bore_provenance') or {}).get('block_export')]
    if blocked:
        labels = ', '.join(blocked[:12]) + (' …' if len(blocked) > 12 else '')
        raise ValueError(f"관경 표기 충돌 {len(blocked)}개: {labels}. "
                         "보라색 배관의 근거를 확인하고 구간 분할 또는 관경 직접 수정을 완료한 뒤 다시 생성하세요.")
