"""domain.fdt — K-Fire FDT(물 도달시간) 물리 코어 패키지.

건식/준비작동식 스프링클러의 트립(블로다운)·충수(트랜짓) 과도해석 물리 코어.
순수 Python으로 유지한다 — PySide6·WNTR·EPANET 의존 금지(계획서 §6 4계층 규약).

- air_model: 오리피스 공기 배출식·가스법칙·초크 판정 (Phase A·B 공용)
- blowdown:  Phase A 트립시간 적분기 + 트립임계 환산 헬퍼
"""
from __future__ import annotations
