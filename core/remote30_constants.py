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



