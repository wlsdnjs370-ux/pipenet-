"""Strict, configured nominal-diameter grammar for already mapped text layers."""
from __future__ import annotations

from dataclasses import dataclass
import math
import re
from typing import Iterable, Sequence


@dataclass(frozen=True)
class DiameterLabelConfig:
    """Explicit project grammar; no arbitrary number or layer discovery."""

    nominal_sizes_mm: tuple[int, ...]
    prefixes: tuple[str, ...] = ()

    def __post_init__(self) -> None:
        if not self.nominal_sizes_mm or any(not isinstance(n,int) or n<=0 for n in self.nominal_sizes_mm):
            raise ValueError('허용 호칭경 목록은 양의 정수여야 합니다.')
        if any(not isinstance(p,str) or not p.strip() for p in self.prefixes):
            raise ValueError('관경 접두어는 빈 문자열일 수 없습니다.')


def extract_mapped_diameter_points(texts: Iterable[Sequence], config: DiameterLabelConfig) -> list[tuple[float,float,int]]:
    """Read one diameter per mapped text; reject quantities and mixed-size notes.

    Accepted examples are 25, 25A, DN25, Ø25, 25mm, and configured prefixes such
    as SP 150 or H/newline/100. Multiple different sizes in one entity require
    separate annotation/leader interpretation, never taking its first number.
    """
    prefixes='|'.join(re.escape(p.strip()) for p in config.prefixes)
    prefix=rf'(?:(?:{prefixes})\s*[-:－]?\s*)?' if prefixes else ''
    pattern=re.compile(rf'^\s*{prefix}(?:DN\s*|[ØøΦφ]\s*)?(\d{{2,3}})\s*(?:A|mm)?\s*$',re.IGNORECASE)
    result=[]
    for row in texts:
        if len(row)!=6:
            continue
        _,_,x,y,_,raw=row
        match=pattern.fullmatch(str(raw or ''))
        if match and int(match[1]) in config.nominal_sizes_mm and all(math.isfinite(float(v)) for v in (x,y)):
            result.append((float(x),float(y),int(match[1])))
    return result
