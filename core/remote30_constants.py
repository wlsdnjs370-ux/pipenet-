# -*- coding: utf-8 -*-
"""remote30 파이프라인 튜닝 상수 (Phase2 분할 — 무-코드 의존, 순환 차단용).

카테고리/레이어 필터, 스냅/브릿지 허용거리, ladder 합성 임계, 배관 재질 상수 등.
"""
from __future__ import annotations

import re


PIPENET_CATEGORIES = {"PIPE", "HEAD", "TEXT", "ALARM"}
# 배관 geometry 가 될 수 없는 카테고리 — 이름 분류를 못 믿고 지오메트리로 배관을
# 되찾을 때(그래프 fallback · 연결관 승격) 공통으로 쓰는 배제 목록.
NON_PIPE_GEOMETRY_CATS = {"HEAD", "TEXT", "ALARM", "ARCH", "EXCLUDE"}
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
    # 접두어 없는 "18층" — 압력표에 사람이 실제로 쓰는 표기. 지상/지하가 먼저
    # 걸리므로 여기까지 오는 건 층수만 적힌 경우뿐이다.
    (re.compile(r"(\d{1,2})\s*층"), "ground"),
)
_FLOOR_LABEL_SPECIAL = {"옥상": 99, "옥탑": 99, "ROOF": 99, "R/F": 99, "RF": 99}
MACHINE_ROOM_SP_LAYERS = {"-소화(SP-고)", "-소화(SP-저)"}
SNAP_TOL_MM = 50.0
# auto_snap_eps 후보 — 도면별 이음매 간격 실측용. 하한은 SNAP_TOL_MM(종전 고정값).
#
# 상한을 200 → 800mm 로 넓혔다 [2026-08-21]. 200 은 «정상 배관 간 최소 이격보다
# 작은 선» 이라는 근거였는데, 타현장 도면 실측에서 그 전제가 틀렸다. 이음매 간격을
# 지배하는 것은 배관 이격이 아니라 **그 도면이 배관선을 끊어놓은 기호 획의 길이**다.
#   대명동 70mm(전체 세그먼트의 55.8%) · 대구오페라세대 76mm(31.5%)
#   대구오페라 288mm(21.1%) · 청라스타필드 MF-304 240mm(56.6%)+318mm(12.2%)
# eps 가 그 길이를 넘어야 획의 양 끝점이 한 노드로 눌려 사라지고 배관이 이어진다.
# 대명동이 잘 돌던 이유가 바로 «70 < 75» 였다 — 우연히 사다리 안에 들어와 있었다.
# 240~318mm 를 쓰는 도면은 200 에서 막혀 배관망이 만 조각으로 부서졌다
# (MF-304: 최대조각 544 · 성분 9,647 · 헤드 5,504 중 도달 233).
# 넓힌 뒤 실측 정점: 대명동 75(불변) · 대구오페라 400 · 죽전 650 · MF-304 500.
SNAP_EPS_CANDIDATES_MM = (50.0, 60.0, 75.0, 90.0, 110.0, 130.0, 160.0, 200.0,
                          250.0, 300.0, 400.0, 500.0, 650.0, 800.0)
# 과대 병합 안전판 — 연장이 기준 대비 이 비율 밑이면 그 후보는 버린다.
SNAP_EPS_MIN_LEN_RATIO = 0.90
# ★ 가드는 «긴 간선» 연장만 잰다 [2026-08-21]. 총 연장으로 재면 위의 기호 획이
# 눌려 사라지는 것까지 «배관이 없어졌다» 로 세어, 정작 옳은 eps 를 걷어찼다
# (MF-304 eps 500: 총연장비 0.795 → 탈락. 그런데 긴 간선만 보면 0.958 로 통과).
# 가드의 본뜻은 «실제 배관이 병합돼 사라지는 것을 막는다» 이고 실배관 런은 길다.
# 기호 획은 이 문턱 아래라 애초에 집계에 안 들어온다.
SNAP_EPS_GUARD_MIN_EDGE_MM = 1000.0
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
# 간선을 가로/세로로 분류할 때의 축평행 허용치 — 관통 교차 티 복원과 겹쳐 그린 중복
# 선분 제거가 공유한다. 전자는 교차점 좌표를 두 선에서 그대로 합성하고 후자는 두 선이
# 같은 선 위인지 보므로, 허용치가 크면 선에서 벗어난 자리를 같은 선으로 오인한다.
# 실측 배관은 축 오차 1mm 미만이라 좁게 잡아도 잃는 게 없다.
CROSS_TEE_AXIS_TOL_MM = 5.0
# 관통 교차 자리에 그려진 부속 기호(티·엘보)를 그 교차의 것으로 인정하는 반경.
# B1F 의 부속 기호는 INSERT 가 아니라 배관 레이어에 직접 그린 ARC 쌍이다 — 교차점을
# 중심(거리 0.0)으로 r=180 짜리 호 두 개가 314~46° / 134~226° 로 마주 본다. 중심이
# 교차점과 정확히 일치하므로 값에 둔감하다(5~800mm 스윕에서 판정 불변). 기호 반지름
# 180mm + 도시 오차를 감안해 300mm 로 둔다.
CROSS_TEE_SYMBOL_TOL_MM = 300.0
# 교차점이 두 간선 중 한쪽의 *끝* 근처면 그 배관이 거기서 끝난다는 뜻 — 스쳐 지나감이
# 아니라 진짜 T 다. 기호가 없어도 이 증거만으로 접속을 인정한다. B1F 실측 무릎이
# 200→400mm 사이에 있고(주망헤드 1538→1856), 400→900mm 에서는 더 늘지 않는다.
# 헤드 틈 접속 상한(HEAD_GAP_JOIN_MAX_MM)과 같은 값 — 둘 다 "기호 하나가 들어가는
# 틈"이라는 같은 도면 관습을 재는 자다.
CROSS_TEE_END_TOL_MM = 400.0
# 이름이 배관과 무관한 레이어를 지오메트리만으로 배관으로 승격하는 문턱 — 그 레이어
# 안에서 "동일선상 두 조각이 헤드 하나를 사이에 두고 벌어진" 헤드 틈이 몇 번 나오는가.
# 실측(B1F 업로드본 39레이어 · B1F 최소 13 · 대명동 · LH306 · LH지하): 이 지문이 잡히는
# 비배관 레이어는 "현장조사#셔터" 하나뿐(15건)이고 건축·밸브·소화전·치수는 전부 0건이다.
# 배관 레이어는 수백~수천 건이라 3 은 우연(한 쌍이 어쩌다 맞아떨어지는 경우)만 걸러낸다.
HEADGAP_PIPE_PROMOTE_MIN = 3
# 헤드 ↔ 가지관 사이의 연결관(후렉시블·드롭)이 배관 레이어가 아니라 기본 '0' 등에
# 그려진 경우의 승격 상한. 대명동 201동 실측: 헤드 118 중 42개가 이 때문에 그래프에
# 붙지 못했고(부착 헤드는 25mm 거리에 `SP 후렉시블` PIPE 런이 있는데 미부착 헤드는
# 그 자리에 같은 모양의 선이 레이어 '0' 으로 그려져 있다), 최근접 배관까지 거리는
# 중앙 617 · p75 790mm 로 300mm 상한을 넘는다. 실제 후렉시블 길이(도면 458~658mm)에
# 꺾임 여유를 더한 값. 승격은 "헤드에서 300mm 안에 끝점이 있고 반대쪽이 배관에 닿는"
# 실제로 그려진 선에만 적용된다 — 없는 선을 만들지 않으므로 추정 연결이 아니다.
HEAD_CONNECTOR_MAX_MM = 1500.0
# 연결관의 반대쪽 끝이 배관에 "닿았다"고 볼 거리 — 끝점 동등성(SNAP_TOL_MM)과 같은 자.
HEAD_CONNECTOR_TOUCH_MM = 50.0
# 연결관으로 인정할 최대 선분 수. 후렉시블은 수직 드롭 + 꺾임 + 수평 접근의 3토막이
# 최대치다(대명동 `SP 후렉시블` 실측: 175 → 141 → 319mm). 더 길게 허용하면 건축선을
# 이어 붙여 없는 배관을 만든다.
HEAD_CONNECTOR_MAX_SEGS = 3
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

# ── 구역 유형 → 관종 매핑 ──────────────────────────────────────────
# 사용자가 화면에서 그린 구역에는 유형 태그가 붙는다. 유형은 "헤드를 어디서 고를까"
# 와 무관하게 "이 구간 배관이 무엇으로 되어 있나"만 뜻한다.
# DXF 에 재질 정보가 없으므로 유형은 사람이 지정해야 하고, 지정이 없으면 강관이다.
ZONE_KIND_UNIT_DWELLING = "unit_dwelling"
ZONE_KIND_PARKING = "parking"
ZONE_KIND_CORRIDOR = "corridor"
ZONE_MATERIAL_MAP: dict[str, tuple[str, str]] = {
    ZONE_KIND_UNIT_DWELLING: (CPVC_PIPE_TYPE, CPVC_C_FACTOR),
    ZONE_KIND_PARKING: (STEEL_PIPE_TYPE, STEEL_C_FACTOR),
    ZONE_KIND_CORRIDOR: (STEEL_PIPE_TYPE, STEEL_C_FACTOR),
}
DEFAULT_ZONE_MATERIAL = (STEEL_PIPE_TYPE, STEEL_C_FACTOR)


# ── 신축배관(FX) 규격 프로파일 — "표 1" (원본 규격표) ──────────────
# 값은 여기에만 존재한다. build_input_tables 등 사용처는 반드시 이 dict를 참조.
# 프로파일 추가 시 여기에만 항목을 늘린다.
# 등재되지 않은 규격은 편집기의 "직접 입력"으로 처리.
# phys_len_m 는 참고용 — 파이프 기하(도면 거리)에 이미 포함되므로 eq_len/길이에 가산 금지.
FX_SPEC_PROFILES: dict[str, dict] = {
    "평균": {                  # 구 A사 유형 값 — 규격 평균 프리셋
        "eq_len_m": 15.6,     # Equipment 등가길이
        "nominal_dn": 20,     # 20A
        "inner_dia_mm": 21.6, # SLF Size-definition 검증 대상
        "c_factor": 120,
        "phys_len_m": 0.7,    # 참고용. 가산 금지.
    },
    "한백표준": {              # 한백에프앤씨 사내 표준(F사 유형) — KSD 25A
        "eq_len_m": 22.4,
        "nominal_dn": 25,
        "inner_dia_mm": 28.0,
        "c_factor": 120,
        "phys_len_m": 0.7,
    },
}
# 기본값은 "평균" 유지 — 사용자 지시로 지시서 T3 의 기본값 교체안은 채택하지 않는다.
# 한백표준은 편집기에서 골라 쓰는 라이브러리 항목이다.
FX_DEFAULT_PROFILE = "평균"
AV_EQ_LEN_M = 12.9             # 알람밸브 등가길이 (기존값 상수화만, 값 변경 없음)

# ── 펌프 ────────────────────────────────────────────────────────────
# <Pump-fan> 의 <Library-pump> 기본 이름. **이 이름이 SLF 의 Pump-section 에
# 실재해야 한다** — 없으면 PIPENET 이 성능곡선 없이 양정을 스스로 선정해 버려서,
# 살 수 없는 펌프로 계산서가 나와도 출력물만 봐서는 알 수 없다. SLF 사본마다
# 펌프 이름이 다르므로(현재 표준 SLF 에는 이 이름이 없다) 실재 여부는 방출
# 시점에 대조하고, 없으면 정격유량·양정으로 곡선을 주입하거나 미확정으로 올린다.
DEFAULT_PUMP_LIBRARY_NAME = "SP_162M_2900LPM"

# ── FX 실배관 materialize 상수 (Stage-6 방출부에서 사용) ──────────────
# 참조 SDF(201동)의 FX 파이프: <Pipe bore=".." length="0.7" rise="-0.1" roughness-or-c="120">
# 내부에 등가길이 Equipment 를 담는 구조. 아래 값은 그 참조와 정합.
FX_SCHEDULE_ROUGHNESS = 0.065  # SLF Metric-definition roughness (Colebrook 전용, C=120 계산엔 무영향 — 참고값)
FX_RISE_M = -0.1               # FX 파이프 rise(입->출 표고차). 참조 SDF 하드코드값.


# ── 표고 출처 ─────────────────────────────────────────────────────
# 근거가 센 것부터. 근거가 아예 없는 값은 0 으로 때우지 않고 UNRESOLVED 로 센다 —
# 0 은 "수평"이라는 주장이라서, 모르는 것과 구분되어야 한다.
ELEV_SOURCE_USER = "user_confirmed"
ELEV_SOURCE_DRAWING = "drawing_estimated"
ELEV_SOURCE_DEFAULT = "default"
ELEV_SOURCE_UNRESOLVED = "unresolved"
ELEV_SOURCE_ORDER = (ELEV_SOURCE_USER, ELEV_SOURCE_DRAWING,
                     ELEV_SOURCE_DEFAULT, ELEV_SOURCE_UNRESOLVED)

# ── 층 내 국소 표고 규칙 (FNCADnet 작업지시서 T6-b) ─────────────────
# 층간 낙차(압력표)와 달리 이건 한 층 안에서의 오르내림이다. 사내 도면 통계가
# 나오면 값만 갈아끼운다. None 은 "아직 근거 없음" — 0 과 다르다.
LOCAL_RISE_RULES: dict[str, float | None] = {
    "parking_beam_drop_m": -0.3,       # 지하주차장 보 하단 하향
    # 상향식 촛대 — KS D 3507 배관용 탄소강관. 수직이라 길이가 곧 상승분이다.
    # 세대 내 하향식은 신축배관(FX)이 맡는다: 길이 0.7m 는 FX_SPEC_PROFILES
    # phys_len_m, 표고차는 FX_RISE_M. 관이 휘어 돌아 길이 ≠ 낙차다.
    "upright_riser_nipple_m": 0.3,
}


# ── 라이저 형상 기본값 (FNCADnet 작업지시서 4-1/4-3) ────────────────
# 아래 네 값은 **대명동 201동 실측값**이다. 근거는 그 현장의 수작업 PIPENET 모델
# (data/sample_problem) — LSP·MSP·LLSP 20여 모델에서 한 자도 다르지 않다:
#   1→2 length=20.95, 3→4 length=14.93, 7→8 length=0.5, 5→10 length=1.5 rise=1.
# 현장이 바뀌면 전부 틀린 값이므로 ProjectContext 로 올려 사람이 확정하기 전에는
# [미확정] 로 표시한다. 여기 값은 "아무도 안 정했을 때의 자리" 이지 표준이 아니다.
RISER_ROOF_RUN_TO_RISER_M = 20.95   # 수원 → 옥상 수평 (r1)
RISER_ROOF_RUN_AFTER_DROP_M = 14.93  # 옥상 하강 뒤 수평 (r3)
RISER_PRV_APPROACH_M = 0.5          # 라이저 → PRV 입구 (r5)
# T분기 → 알람밸브. 수작업 모델 전량이 rise=+1.0(알람밸브가 위), length=1.5 —
# 즉 수직 1.0m + 수평 0.5m 의 L 자다. 현업 관행문(수평 0.6~1.0m)과는 수평분이
# 다르지만, 실측이 있는 쪽을 기본값으로 둔다.
TEE_TO_ALARM_VALVE_RISE_M = 1.0
TEE_TO_ALARM_VALVE_RUN_M = 0.5

# 현업 관행값 — 산출물에는 아직 반영하지 않는다(표고 기준면 미확정, BLOCKED.md §26).
# 값만 서류에 남기고 "채웠으니 반영됐다" 로 읽히지 않게 기록 전용으로 표시한다.
TEE_BRANCH_ABOVE_SLAB_M = 0.6       # 입상관 T분기 = 그 층 바닥 + 600mm
TOP_FLOOR_EXTRA_HEIGHT_M = 0.25     # 최상층은 단열재분만큼 층고가 높다

# 관경 전이(레듀서)는 T분기점에서 이 거리 안에 있으면 T분기점에 귀속시킨다.
# 현업 관행 — 300mm 를 독립 노드로 살리면 관리 포인트만 늘고 그 차이는 오차범위
# 안이라는 판단. 25/32 가 갈릴 때는 불리한 작은 쪽을 택한다.
# 현재 추출망은 전이가 이미 노드에서만 일어나 결과적으로 같은 동작을 한다. 상수를
# 두는 목적은 동작 변경이 아니라 자동본이 수동본과 다를 때 이유를 말할 수 있게
# 하는 것이므로, 값을 쓰지 않고 **귀속된 지점 수만** 세어 방출 리포트에 낸다.
REDUCER_SNAP_TO_TEE_MM = 300.0


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



