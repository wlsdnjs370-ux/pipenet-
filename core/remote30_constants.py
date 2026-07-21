# -*- coding: utf-8 -*-
"""remote30 파이프라인 튜닝 상수 (Phase2 분할 — 무-코드 의존, 순환 차단용).

카테고리/레이어 필터, 스냅/브릿지 허용거리, ladder 합성 임계, 배관 재질 상수 등.
"""
from __future__ import annotations

import re


PIPENET_CATEGORIES = {"PIPE", "HEAD", "TEXT", "ALARM"}
KEEP_BASE_LAYERS = {"0"}  # INSERT BYLAYER 공통 + 도면 컨텍스트
_DIA_TEXT_PATTERNS = (
    re.compile(r"\b(\d{2,3})\s*A\b"),                  # 25A
    re.compile(r"^\s*(\d{2,3})\s*$"),                  # 순수 숫자
    re.compile(r"[Øø]\s*(\d{2,3})"),                   # Ø25
    re.compile(r"DN\s*(\d{2,3})"),                     # DN25
    re.compile(r"(?<![0-9])(\d{2,3})\s*mm(?![0-9])"),  # 25mm
)
_DIA_TEXT_NOISE_KW = ("호스", "방수구", "소화전", "옥내", "HOSE", "EA", "KG",
                      "SET", "SCALE", "PUMP", "펌프", "TANK", "탱크", "SIZE")
_VALID_DIA_MM = frozenset((15, 20, 25, 32, 40, 50, 65, 80, 100, 125, 150, 200, 250, 300))
_FLOOR_LABEL_PATTERNS = (
    (re.compile(r"지상\s*(\d{1,2})\s*층"), "ground"),     # 지상N층 → +N
    (re.compile(r"지하\s*(\d{1,2})\s*층"), "basement"),   # 지하N층 → -N
    (re.compile(r"B\s*(\d{1,2})\s*F", re.I), "basement"),  # B1F → -1
    (re.compile(r"(?<![A-Za-z])(\d{1,2})\s*F(?![A-Za-z])"), "ground"),  # 5F → +5
)
_FLOOR_LABEL_SPECIAL = {"옥상": 99, "옥탑": 99, "ROOF": 99, "R/F": 99, "RF": 99}
MACHINE_ROOM_SP_LAYERS = {"-소화(SP-고)", "-소화(SP-저)"}
SNAP_TOL_MM = 50.0
HEAD_BRIDGE_MAX_MM = 5000.0  # 헤드 INSERT 좌표 ↔ 가장 가까운 그래프 노드 brigde 허용 거리.
SOURCE_BRIDGE_MAX_MM = 25000.0  # 알람밸브 (source) ↔ 배관망 nearest bridge 허용 (25m).
MIN_PIPE_EDGE_MM = 50.0
CLOSED_PL_TOL_MM = 5.0  # PL 의 첫점과 마지막점이 이 거리 안이면 closed polygon 으로 간주 → 그래프 제외
LADDER_MAX_RUNG_MM = 300.0     # rung (짧은 cross 변) 최대 길이. 단위세대 도면 기준.
LADDER_MIN_RAIL_RATIO = 3.0    # rail / rung 평균 길이 비. 정사각형 (=1) 은 합성 안 됨.
LADDER_PARALLEL_COS = 0.985    # 두 rail 의 방향 cos 유사도 임계값 (≈ 10도 안)
LADDER_MAX_ITER = 10           # collapse 반복 횟수 (합성 후 새 ladder 생길 수 있음)
STEEL_PIPE_TYPE = "KSD 3507"
STEEL_C_FACTOR = "120"
CPVC_PIPE_TYPE = "CPVC2"
CPVC_C_FACTOR = "150"
ORTHO_SNAP_TOL_DEG = 20.0


# ── 신축배관(FX) 규격 프로파일 — "표 1" (원본 규격표) ──────────────
# 값은 여기에만 존재한다. build_input_tables 등 사용처는 반드시 이 dict를 참조.
# 프로파일 추가 시 여기에만 항목을 늘린다.
# 설계사별 신축배관 입력유형(신축배관 테이블.pptx) 중 eq_len 이 정상 입력된 5종만 등재.
#   제외: C사·D사(등가길이 별도 안잡음 → FX 손실 0), G사(C값 조작 22.4 — 가드레일 위반).
# phys_len_m 는 참고용 — 파이프 기하(도면 거리)에 이미 포함되므로 eq_len/길이에 가산 금지.
FX_SPEC_PROFILES: dict[str, dict] = {
    "사내표준": {              # 구 F사 유형 — 한백 사내표준 (2026-07 확정)
        "eq_len_m": 22.4,     # Equipment 등가길이
        "nominal_dn": 25,     # 25A — 말단 파이프 최소 호칭경과 정합
        "inner_dia_mm": 28.0, # SLF Size-definition 검증 대상
        "c_factor": 120,
        "phys_len_m": 0.7,    # 참고용. 가산 금지.
    },
    "A사": {
        "eq_len_m": 15.6,
        "nominal_dn": 20,
        "inner_dia_mm": 21.6,
        "c_factor": 120,
        "phys_len_m": 0.7,
    },
    "B사": {
        "eq_len_m": 11.5,
        "nominal_dn": 20,
        "inner_dia_mm": 21.5,
        "c_factor": 120,
        "phys_len_m": 0.6,
    },
    "E사": {
        "eq_len_m": 20.0,
        "nominal_dn": 20,
        "inner_dia_mm": 21.6,
        "c_factor": 120,
        "phys_len_m": 1.0,
    },
    "H사": {
        "eq_len_m": 7.8,
        "nominal_dn": 20,
        "inner_dia_mm": 21.6,
        "c_factor": 120,
        "phys_len_m": 0.7,
    },
}
FX_DEFAULT_PROFILE = "사내표준"
AV_EQ_LEN_M = 12.9             # 알람밸브 등가길이 (기존값 상수화만, 값 변경 없음)



