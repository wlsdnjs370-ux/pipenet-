"""
domain/fitting_models.py

K-Fire v3.0 — 부속·밸브 등가길이 라이브러리 도메인 모델.
ADR-001 §Detailed Design 구현체.
PySide6 의존성 없음.
"""
from __future__ import annotations

import uuid
from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import TypedDict, Optional, Union

from domain.constants import DN_STANDARD, FittingCategory, CellOrigin


# ---------------------------------------------------------------------------
# M1.4 — SourceTag TypedDict
# ---------------------------------------------------------------------------

class SourceTagNFPA(TypedDict):
    kind: str       # "NFPA13"
    edition: str    # "2025"
    table: str      # "28.2.4.4.1"


class SourceTagUser(TypedDict, total=False):
    kind: str       # "USER"
    note: str


SourceTag = Union[SourceTagNFPA, SourceTagUser]


def make_nfpa_source() -> SourceTagNFPA:
    """NFPA 13 2025판 출처 태그 생성 헬퍼."""
    return {"kind": "NFPA13", "edition": "2025", "table": "28.2.4.4.1"}


def make_user_source(note: str = "") -> SourceTagUser:
    """사용자 정의 출처 태그 생성 헬퍼."""
    return {"kind": "USER", "note": note} if note else {"kind": "USER"}


# ---------------------------------------------------------------------------
# M1.5 — CellValue dataclass (H3: CellOrigin enum 사용)
# ---------------------------------------------------------------------------

@dataclass
class CellValue:
    """등가길이 셀 1개 (호칭경 1개에 해당하는 값).

    origin은 반드시 CellOrigin enum 멤버를 사용한다.
    문자열 리터럴 직접 비교는 H3 원칙 위반 — CellOrigin 멤버로 비교할 것.
    """
    value: Optional[float]              # m 단위. null = NFPA 미정의
    origin: CellOrigin                  # 출처 구분 (enum)
    original_value: Optional[float] = None   # override/filled 직전 값(감사용)
    updated_at: Optional[str] = None         # ISO8601 최종 수정 시각
    updated_by: str = "local"                # 최종 수정자

    def to_dict(self) -> dict:
        return {
            "value": self.value,
            "origin": self.origin.value,     # str로 직렬화
            "original_value": self.original_value,
            "updated_at": self.updated_at,
            "updated_by": self.updated_by,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "CellValue":
        origin_str = d.get("origin", "")
        try:
            origin = CellOrigin(origin_str)
        except ValueError:
            raise ValueError(
                f"CellValue.from_dict: 알 수 없는 origin 값 '{origin_str}'. "
                f"허용값: {[e.value for e in CellOrigin]}"
            )
        raw_value = d.get("value")
        value = float(raw_value) if raw_value is not None else None
        raw_orig = d.get("original_value")
        original_value = float(raw_orig) if raw_orig is not None else None
        return cls(
            value=value,
            origin=origin,
            original_value=original_value,
            updated_at=d.get("updated_at"),
            updated_by=d.get("updated_by", "local"),
        )


def make_nfpa_cell(value: Optional[float]) -> CellValue:
    """NFPA 시드 셀 생성 헬퍼. null은 value=None으로."""
    return CellValue(value=value, origin=CellOrigin.NFPA_ORIGINAL)


def make_user_null_cell() -> CellValue:
    """USER 부속의 미입력 셀 생성 헬퍼 (USER_FILLED, value=None)."""
    return CellValue(value=None, origin=CellOrigin.USER_FILLED)


# ---------------------------------------------------------------------------
# M1.6 — FittingDefinition dataclass
# ---------------------------------------------------------------------------

@dataclass
class FittingDefinition:
    """부속·밸브 라이브러리 항목 1개.

    equivalent_lengths 키는 int(DN) 기준. JSON 저장 시 str("15" 등)으로
    직렬화되므로 로딩 시 반드시 _get_cell() 헬퍼로 접근한다.
    """
    id: str
    name_ko: str
    name_en: str
    category: FittingCategory
    source: SourceTag
    equivalent_lengths: dict[int, CellValue]    # key: DN(int)
    notes: str = ""
    audit: dict = field(default_factory=dict)

    def __post_init__(self) -> None:
        if not isinstance(self.category, FittingCategory):
            try:
                self.category = FittingCategory(str(self.category))
            except ValueError:
                self.category = FittingCategory.USER_DEFINED

        if not self.audit:
            now = _now_iso()
            self.audit = {
                "created_at": now,
                "created_by": "system",
                "updated_at": now,
                "updated_by": "system",
                "revision": 1,
            }

    # --- M1.8: DN 키 정규화 접근자 ---

    def get_cell(self, dn: int) -> Optional[CellValue]:
        """호칭경 DN으로 셀 조회. int 키로만 접근 (문자열 직접 접근 금지)."""
        return self.equivalent_lengths.get(dn)

    # --- 직렬화 ---

    def to_dict(self) -> dict:
        eql: dict = {}
        for dn, cell in self.equivalent_lengths.items():
            eql[str(dn)] = cell.to_dict()
        return {
            "id": self.id,
            "name_ko": self.name_ko,
            "name_en": self.name_en,
            "category": self.category.value,
            "source": dict(self.source),
            "equivalent_lengths": eql,
            "notes": self.notes,
            "audit": dict(self.audit),
        }

    @classmethod
    def from_dict(cls, d: dict) -> "FittingDefinition":
        category_str = d.get("category", "")
        try:
            category = FittingCategory(category_str)
        except ValueError:
            category = FittingCategory.USER_DEFINED

        raw_eql = d.get("equivalent_lengths", {})
        # normalize_equivalent_lengths: str 키 → int 키 변환 + CellValue 역직렬화
        eql = normalize_equivalent_lengths(raw_eql)

        return cls(
            id=d.get("id", str(uuid.uuid4().hex)),
            name_ko=d.get("name_ko", ""),
            name_en=d.get("name_en", ""),
            category=category,
            source=d.get("source", make_user_source()),
            equivalent_lengths=eql,
            notes=d.get("notes", ""),
            audit=d.get("audit", {}),
        )


# ---------------------------------------------------------------------------
# M1.7 — 스키마 검증 함수 4종
# ---------------------------------------------------------------------------

class FittingSchemaError(ValueError):
    """FittingDefinition 스키마 위반 예외."""


def validate_equivalent_lengths_keys(fitting: FittingDefinition) -> None:
    """검증 ①: equivalent_lengths 키 집합 = DN_STANDARD 정확 일치."""
    expected = set(DN_STANDARD)
    actual = set(fitting.equivalent_lengths.keys())
    if actual != expected:
        missing = expected - actual
        extra = actual - expected
        parts = []
        if missing:
            parts.append(f"누락 DN: {sorted(missing)}")
        if extra:
            parts.append(f"초과 DN: {sorted(extra)}")
        raise FittingSchemaError(
            f"[{fitting.id}] equivalent_lengths 키 오류 — {'; '.join(parts)}"
        )


def validate_cell_values(fitting: FittingDefinition) -> None:
    """검증 ②: value > 0 & < 1000 (sanity bound). null은 통과."""
    for dn, cell in fitting.equivalent_lengths.items():
        if cell.value is not None:
            if not (0 < cell.value < 1000):
                raise FittingSchemaError(
                    f"[{fitting.id}] DN{dn} 셀 값 범위 위반: "
                    f"{cell.value} (허용: 0 < value < 1000)"
                )


def validate_nfpa_uniqueness(fittings: list[FittingDefinition]) -> None:
    """검증 ③: NFPA13 출처 항목은 카테고리당 최대 1개."""
    seen: dict[FittingCategory, str] = {}
    for f in fittings:
        if f.source.get("kind") == "NFPA13":
            if f.category in seen:
                raise FittingSchemaError(
                    f"카테고리 '{f.category.value}'에 NFPA13 출처 항목이 2개 이상: "
                    f"'{seen[f.category]}' 및 '{f.id}'"
                )
            seen[f.category] = f.id


def validate_user_name_ko_unique(fittings: list[FittingDefinition]) -> None:
    """검증 ④: USER 출처 항목은 동일 카테고리 내 name_ko 중복 금지."""
    seen: dict[tuple[FittingCategory, str], str] = {}
    for f in fittings:
        if f.source.get("kind") == "USER":
            key = (f.category, f.name_ko)
            if key in seen:
                raise FittingSchemaError(
                    f"USER 부속 name_ko 중복: 카테고리 '{f.category.value}', "
                    f"name_ko '{f.name_ko}' — '{seen[key]}' 및 '{f.id}'"
                )
            seen[key] = f.id


def validate_fitting(fitting: FittingDefinition) -> None:
    """단일 FittingDefinition 전체 검증 (①+②)."""
    validate_equivalent_lengths_keys(fitting)
    validate_cell_values(fitting)


def validate_library(fittings: list[FittingDefinition]) -> None:
    """라이브러리 전체 검증 (①+②+③+④)."""
    for f in fittings:
        validate_fitting(f)
    validate_nfpa_uniqueness(fittings)
    validate_user_name_ko_unique(fittings)


# ---------------------------------------------------------------------------
# M1.8 — JSON DN 키 정규화 헬퍼 (모듈 레벨)
# ---------------------------------------------------------------------------

def normalize_equivalent_lengths(raw: dict) -> dict[int, CellValue]:
    """JSON에서 읽은 equivalent_lengths를 int 키 dict로 정규화.

    키가 문자열 "15" 여도 int(15)로 변환.
    CellValue.from_dict()로 역직렬화.
    잘못된 origin 값이 있으면 ValueError 발생 (silent bug 방지).
    """
    result: dict[int, CellValue] = {}
    for k, v in raw.items():
        try:
            dn = int(k)
        except (TypeError, ValueError):
            raise ValueError(f"DN 키를 int로 변환 불가: '{k}'")
        result[dn] = CellValue.from_dict(v)
    return result


def get_cell(eql: dict[int, CellValue], dn: int) -> Optional[CellValue]:
    """equivalent_lengths dict에서 DN(int)으로 안전 조회.

    문자열 키 직접 접근 금지 패턴 대신 이 함수를 사용할 것.
    """
    return eql.get(dn)


# ---------------------------------------------------------------------------
# 내부 유틸
# ---------------------------------------------------------------------------

def _now_iso() -> str:
    return datetime.now(tz=timezone.utc).isoformat(timespec="seconds")


def new_fitting_id() -> str:
    """USER 부속 신규 등록 시 UUID 자동 부여. 사용자에게는 노출하지 않음."""
    return uuid.uuid4().hex


# ---------------------------------------------------------------------------
# __all__
# ---------------------------------------------------------------------------

__all__ = [
    "SourceTag",
    "SourceTagNFPA",
    "SourceTagUser",
    "make_nfpa_source",
    "make_user_source",
    "CellValue",
    "make_nfpa_cell",
    "make_user_null_cell",
    "FittingDefinition",
    "FittingSchemaError",
    "validate_fitting",
    "validate_library",
    "validate_equivalent_lengths_keys",
    "validate_cell_values",
    "validate_nfpa_uniqueness",
    "validate_user_name_ko_unique",
    "normalize_equivalent_lengths",
    "get_cell",
    "new_fitting_id",
]
