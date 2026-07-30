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
# auto_snap_eps 후보 — 도면별 이음매 간격 실측용. 하한은 SNAP_TOL_MM(종전 고정값),
# 상한 200mm 는 정상 배관 간 최소 이격보다 작아 서로 다른 배관을 병합하지 않는 선.
SNAP_EPS_CANDIDATES_MM = (50.0, 60.0, 75.0, 90.0, 110.0, 130.0, 160.0, 200.0)
# 과대 병합 안전판 — 총 연장이 기준 대비 이 비율 밑이면 그 후보는 버린다.
SNAP_EPS_MIN_LEN_RATIO = 0.90
HEAD_BRIDGE_MAX_MM = 5000.0  # 헤드 INSERT 좌표 ↔ 가장 가까운 그래프 노드 brigde 허용 거리.
# 헤드 기호 ↔ 배관 결합선(drop line) 상한. 이건 배관이 아니라 "기호가 어느 배관에
# 달렸는가" 의 판정이므로 기호 도시 오차 수준이어야 한다. 대명동 서측 실측 분포가
# 중앙 25mm 인데 상한 5000mm 를 쓰면 수 m 떨어진 남의 가지에 헤드가 붙는다.
# 넘는 헤드는 붙이지 않고 unreachable 로 보고한다 — 추정 연결 폐지 원칙.
HEAD_DROP_MAX_MM = 300.0
SOURCE_BRIDGE_MAX_MM = 25000.0  # 알람밸브 (source) ↔ 배관망 nearest bridge 허용 (25m).
ANCHOR_W_MARGIN_MM = 3000.0  # anchored 작업창 W = convex_hull(head_region ∪ {alarm_xy}) 팽창 여유.
MIN_PIPE_EDGE_MM = 50.0
# 느슨한 끝점 ↔ 다른 edge 내부(수선발) 거리가 이 안이면 T분기로 보고 edge 를 쪼갠다.
# 도면 스케일 비례(적응형) 금지 — 이 갭은 CAD 작도 정밀도이지 도면 크기의 함수가 아니다.
# 4개 도면 실측(대각선 45m~784m)에서 갭 분포가 동일하게 양봉이었다: 정확한 T분기는
# ≤5mm 에 몰리고, 20mm 를 넘으면 평행 2줄 배관의 rail 간격 같은 다른 집단이 섞인다
# (B1F 는 20→50mm 에서만 +481건, 50→100mm 에서 +34,811건). 대명동 산출물은 2~50mm 에서
# 완전히 동일하고 100mm 에서 붕괴(배출망 73.5m→100.9m)하므로, 네 도면의 공통 골짜기인
# 20mm 를 택한다. 상한은 SNAP_TOL_MM — 분기점을 수선발이 아닌 끝점 u 로 잡는 근거가
# "둘이 노드 동등성 epsilon 안"이라, 그 값을 넘으면 전제 자체가 깨진다.
TEE_SPLIT_MAX_MM = 20.0
# 헤드 기호가 끊어놓은 동일선상 배관을 잇는 상한. 일부 도면은 가지관 런을 헤드마다
# 끊어 그리고 그 틈에 기호를 넣는다 — B1F 실측 간극 200~400mm(기호는 148.6mm 십자).
# 이 간격은 이음매 간격(SNAP_EPS_CANDIDATES_MM 상한 200mm)을 넘어 끝점 클러스터로는
# 안 붙고, 같은 축이라 T분기 복원에도 안 걸린다. 기호가 들어가는 틈이므로 상한은
# 기호 크기 + 양쪽 여백 수준이면 충분하다.
HEAD_GAP_JOIN_MAX_MM = 400.0
# 위 판정의 축직교 오프셋 허용(배관 동일선상 · 헤드가 그 선 위). 실측은 배관 오프셋
# 1mm 미만 · 헤드 오프셋 0mm 로 사실상 정확히 일치하며, 1~20mm 구간은 5532쌍 중 7쌍뿐
# 이라 값에 둔감하다. 기호 도시 오차를 감안해 MIN_PIPE_EDGE_MM 수준으로 둔다.
HEAD_GAP_JOIN_TOL_MM = 50.0
# 관통 교차(끝점이 어느 쪽에도 얹히지 않은 직교 X)를 티로 해석할 때의 축평행 허용치.
# 교차점 좌표를 두 선에서 그대로 합성하므로 허용치가 크면 선에서 벗어난 자리에 접속이
# 생긴다. 실측 배관은 축 오차 1mm 미만이라 좁게 잡아도 잃는 게 없다.
CROSS_TEE_AXIS_TOL_MM = 5.0
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
# 프리셋은 "평균"(구 A사 유형 값) 1종만 등재 — 그 외 규격은 편집기의 "직접 입력"으로 처리.
# phys_len_m 는 참고용 — 파이프 기하(도면 거리)에 이미 포함되므로 eq_len/길이에 가산 금지.
FX_SPEC_PROFILES: dict[str, dict] = {
    "평균": {                  # 구 A사 유형 값 — 규격 평균 프리셋
        "eq_len_m": 15.6,     # Equipment 등가길이
        "nominal_dn": 20,     # 20A
        "inner_dia_mm": 21.6, # SLF Size-definition 검증 대상
        "c_factor": 120,
        "phys_len_m": 0.7,    # 참고용. 가산 금지.
    },
}
FX_DEFAULT_PROFILE = "평균"
AV_EQ_LEN_M = 12.9             # 알람밸브 등가길이 (기존값 상수화만, 값 변경 없음)

# ── FX 실배관 materialize 상수 (Stage-6 방출부에서 사용) ──────────────
# 참조 SDF(201동)의 FX 파이프: <Pipe bore=".." length="0.7" rise="-0.1" roughness-or-c="120">
# 내부에 등가길이 Equipment 를 담는 구조. 아래 값은 그 참조와 정합.
FX_SCHEDULE_ROUGHNESS = 0.065  # SLF Metric-definition roughness (Colebrook 전용, C=120 계산엔 무영향 — 참고값)
FX_RISE_M = -0.1               # FX 파이프 rise(입->출 표고차). 참조 SDF 하드코드값.


def fx_schedule_name(nominal_dn: int, inner_dia_mm: float) -> str:
    """규격 기하 → PIPENET-safe ASCII 스케줄명. (한글/공백 없이 SLF Item-name 겸용)

    예: (25, 28.0) → "FX_25A_28",  (20, 21.6) → "FX_20A_216",  (20, 21.5) → "FX_20A_215".
    호칭경이 같아도 내경이 다르면 별개 스케줄로 구분된다.
    """
    inner_str = ("%g" % float(inner_dia_mm)).replace(".", "")
    return f"FX_{int(nominal_dn)}A_{inner_str}"


def fx_geometry_key(profile: dict) -> tuple[int, float]:
    """FX 프로파일 → dedup 키 (nominal_dn, inner_dia_mm). 같은 기하는 스케줄 1개로 병합."""
    return (int(profile["nominal_dn"]), float(profile["inner_dia_mm"]))



