# -*- coding: utf-8 -*-
"""인식 셸의 임계값 전부 — 지시서 §3.1 "코드에 하드코딩 금지".

한 곳에 모으는 이유가 둘이다. 흩어지면 어떤 값으로 어떤 결과를 얻었는지 되짚을
수 없고, 인식 캐시 키(§11.2)가 이 파일의 해시에 걸려 있어 값이 코드에 박히면
파라미터를 튜닝해도 옛 결과가 그대로 나온다.

지시서 표에 근거가 있는 값과 이 구현이 정한 값을 구분한다. 후자는 `# 미검증`
(작업 규칙 5) — 실 도면 벤치마크로 확인하기 전까지는 주장이 아니라 출발점이다.
"""
from __future__ import annotations

# ── 단위 ────────────────────────────────────────────────────────────────
# 국내 건축 DXF 는 대부분 mm 단위지만 확정은 GATE 가 한다. 여기 값들은 전부 mm
# 기준이므로 `unit_to_mm` 이 틀리면 모든 판정이 함께 틀린다.
UNIT_TO_MM_DEFAULT = 1.0
# 건물 한 층의 대각 길이가 이 범위를 벗어나면 단위 가정이 틀렸다는 신호다.
PLAUSIBLE_DIAG_MIN_MM = 3_000.0        # 미검증 — 3m 미만은 층 평면이 아니다
PLAUSIBLE_DIAG_MAX_MM = 1_000_000.0    # 미검증 — 1km 초과는 mm 가정이 아니다

# ── C130 평행쌍 (§3.1 표) ───────────────────────────────────────────────
PARALLEL_ANGLE_TOL_DEG = 2.0           # CAD 스냅 오차 흡수
PARALLEL_OVERLAP_MIN_RATIO = 0.30      # 스쳐 지나가는 선 배제
PARALLEL_OFFSET_MIN_MM = 50.0          # 국내 벽 두께 실측 범위 하한
PARALLEL_OFFSET_MAX_MM = 500.0         # 상한
OFFSET_HIST_BIN_MM = 10.0
OFFSET_PEAK_MIN_SHARE = 0.08           # 전체 평행쌍의 8% 미만은 노이즈
OFFSET_PEAK_MAX = 3                    # 상위 3 peak

# ── C130 그 밖의 지문 ───────────────────────────────────────────────────
ANGLE_BUCKET_DEG = 2.0                 # 평행 후보 각도 버킷. 허용오차와 같게 둔다
ARC_ATTACH_TOL_MM = 100.0              # 미검증 — ARC 끝점이 LINE 끝점에 "붙었다"
DOOR_ARC_RADIUS_MIN_MM = 600.0         # §3.2 DOOR 판정의 반경 범위
DOOR_ARC_RADIUS_MAX_MM = 1200.0
CLOSED_GAP_TOL_MM = 50.0               # 미검증 — 첫점-끝점 간극이 이 이하면 폐합
SMALL_CLOSED_AREA_MAX_M2 = 6.0         # §3.2 SHAFT 소형 폐합 기준
GRID_ALIGN_TOL_MM = 50.0               # 미검증 — 같은 통심에 정렬됐다고 볼 오차
LONG_LINE_DIAG_RATIO = 0.8             # §3.2 GRID — bbox 대각의 0.8 이상이면 초장선
STAIR_GAP_CV_MAX = 0.15                # §3.2 STAIR 간격 변동계수 상한
STAIR_GAP_TOL_RATIO = 0.15             # 미검증 — 등간격 run 을 잇는 상대 허용오차
STAIR_LEN_CV_MAX = 0.30                # 미검증 — 디딤판은 길이도 고르다
STAIR_BUNDLE_MIN = 6                   # §3.2 STAIR 다발 최소 개수
NODE_SNAP_TOL_MM = 20.0                # 미검증 — 끝점 군집 공차(격자 snap 아님)
QUAD_MAX_DEGREE = 8                    # 사각형 탐색에서 건너뛸 과밀 노드 차수

# 숫자 전용 TEXT 판정 — 치수선은 "3,600" "Ø100" "12.5m²" 처럼 쓰인다.
NUMERIC_TEXT_CHARS = "0123456789 .,-+×xX*/%'\"°ØΦø∅@#()~=<>mM㎜㎡²³"

# ── C140 카테고리 판정 (§3.2 표) ────────────────────────────────────────
# 기본 신뢰도는 표에 그대로 있다. 임계값도 대부분 표에 있고, 표가 "TEXT 이고"
# "INSERT 밀집" 처럼 말로만 적은 것은 여기서 수치로 정하고 `# 미검증` 을 단다.
NAME_HINT_BONUS = 0.15                 # 이름은 가산점 전용 — 판정 단독 근거 금지
OTHER_CONFIDENCE = 0.30                # 미검증 — "봤는데 그 밖" 이지 "안 봤다" 가 아니다

WALL_PARALLEL_MIN_RATIO = 0.55
WALL_LEN_MEDIAN_MIN_MM = 800.0
CONF_WALL = 0.80

DOOR_ARC_ATTACH_MIN_RATIO = 0.35
DOOR_RADIUS_MIN_RATIO = 0.50
CONF_DOOR = 0.75

CONF_WINDOW = 0.65

COLUMN_REPEAT_MIN = 0.70
COLUMN_GRID_ALIGN_MIN = 0.60
COLUMN_AREA_MIN_M2 = 0.10
COLUMN_AREA_MAX_M2 = 4.00
CONF_COLUMN = 0.80

# 미검증 — §3.2 STAIR 행은 "다발 ≥ 6개" 라는 절대 개수뿐이다. 20,000선짜리
# 레이어에서는 등간격 6선 다발이 우연히 거의 항상 나와서, 실 도면의 가장 큰
# 레이어 셋이 전부 STAIR 0.70 으로 잡혔다. 다발이 레이어에서 차지하는 몫도 본다.
STAIR_BUNDLE_MIN_SHARE = 0.05
CONF_STAIR = 0.70

# 미검증 — §3.2 SHAFT 행에는 상한(6㎡)만 있고 하한이 없다. 실 도면에서 가구·해치
# 조각 같은 0.001㎡ 폐합 수천 개가 그대로 SHAFT 로 잡혔다. COLUMN 행의 하한과
# 같은 값을 쓴다: 이보다 작으면 사람이 드나드는 수직 관통부가 아니다.
SHAFT_AREA_MIN_M2 = 0.10
CONF_SHAFT_MULTIFLOOR = 0.85
CONF_SHAFT_SINGLE = 0.40               # 다층 반복을 확인 못 하면 여기까지만

DIM_NUMERIC_MIN_RATIO = 0.80
CONF_DIM = 0.85

ROOM_TEXT_SHARE_MIN = 0.50             # 미검증 — 레이어가 TEXT 위주여야 한다
ROOM_TEXT_NUMERIC_MAX_RATIO = 0.30
ROOM_TEXT_LEN_MIN = 1
ROOM_TEXT_LEN_MAX = 20
CONF_ROOM_TEXT = 0.70

GRID_LONG_LINE_MIN_RATIO = 0.20        # 미검증 — 통심선이 레이어의 이 비율 이상
CONF_GRID = 0.75

FURNITURE_INSERT_SHARE_MIN = 0.50      # 미검증 — "INSERT 밀집" 의 수치화
CONF_FURNITURE = 0.40

# ── 성능 예산 (부록 C.1) ────────────────────────────────────────────────
LAYER_SAMPLE_LIMIT = 20_000            # 레이어당 엔티티 상한. 초과 시 무작위 표본
SAMPLE_SEED = 20260805                 # 표본이 바뀌면 지문도 바뀐다 — 고정한다


def recognize_params() -> dict:
    """인식 캐시 키(§11.2)에 실릴 파라미터 전체.

    이름을 열거하지 않고 모듈에서 걷는 이유는, 새 임계값을 추가하고 캐시 키에
    넣는 것을 잊는 실수를 원천 차단하기 위해서다. 잊으면 튜닝해도 옛 결과가
    나오고, 그 사실을 알아챌 방법이 없다.
    """
    return {
        key: value
        for key, value in sorted(globals().items())
        if key.isupper() and isinstance(value, (int, float, str, tuple, list))
    }
