"""
도메인 계층 — 펌프/노즐/배관 유틸리티

PySide6 의존성이 없는 순수 파이썬 함수.
"""
from dataclasses import dataclass, field
import math


@dataclass
class PumpCurveValidationError(Exception):
    """펌프 검증 오류 코드 + 파라미터 운반체. UI 레이어에서 번역한다."""

    code: str
    params: dict = field(default_factory=dict)

    def __str__(self) -> str:
        return f"PumpCurveValidationError({self.code}, {self.params})"


PUMP_ERR_INVALID_RATED = "pump.err.invalid_rated"
PUMP_ERR_RATED_Q_NONPOSITIVE = "pump.err.rated_q_nonpositive"
PUMP_ERR_RATED_P_NONPOSITIVE = "pump.err.rated_p_nonpositive"
PUMP_ERR_TOO_FEW_POINTS = "pump.err.too_few_points"
PUMP_ERR_TOO_MANY_POINTS = "pump.err.too_many_points"
PUMP_ERR_INVALID_POINT_VALUE = "pump.err.invalid_point_value"
PUMP_ERR_NEGATIVE_HEAD = "pump.err.negative_head"
PUMP_ERR_FIRST_FLOW_NOT_ZERO = "pump.err.first_flow_not_zero"
PUMP_ERR_ZERO_HEAD_NOT_LAST = "pump.err.zero_head_not_last"
PUMP_ERR_FLOW_NOT_INCREASING = "pump.err.flow_not_increasing"
PUMP_ERR_HEAD_RISING = "pump.err.head_rising"
PUMP_ERR_SHUTOFF_TOO_HIGH = "pump.err.shutoff_too_high"
PUMP_ERR_150PCT_OUT_OF_RANGE = "pump.err.150pct_out_of_range"
PUMP_ERR_150PCT_HEAD_TOO_LOW = "pump.err.150pct_head_too_low"
PUMP_ERR_PARAM_NONPOSITIVE = "pump.err.param_nonpositive"
PUMP_ERR_SHUTOFF_BELOW_RATED = "pump.err.shutoff_below_rated"
PUMP_ERR_RATED_BELOW_PEAK_P = "pump.err.rated_below_peak_p"
PUMP_ERR_PEAK_Q_BELOW_RATED = "pump.err.peak_q_below_rated"

PUMP_PARAM_RATED_Q = "pump.param.rated_q"
PUMP_PARAM_RATED_P = "pump.param.rated_p"
PUMP_PARAM_SHUTOFF_P = "pump.param.shutoff_p"
PUMP_PARAM_PEAK_Q = "pump.param.peak_q"
PUMP_PARAM_PEAK_P = "pump.param.peak_p"

# ---------------------------------------------------------------------------
# 10점 자동 커브 생성용 상수 (정규화 기준: x = Q/rated_q, y = H/rated_p)
# ---------------------------------------------------------------------------
PUMP_AUTO_SHUTOFF_HEAD_RATIO: float = 1.20   # 체절점: 정격양정의 120%
PUMP_AUTO_OVERLOAD_FLOW_RATIO: float = 1.50  # 과부하 유량: 정격유량의 150%
PUMP_AUTO_OVERLOAD_HEAD_RATIO: float = 0.65  # 과부하 양정: 정격양정의 65%
PUMP_OVERLOAD_HEAD_TOLERANCE_RATIO: float = 0.05  # 150% 지점 양정 검증 허용폭: ±5%p
PUMP_AUTO_ZERO_HEAD_RATIO: float = 0.0       # 0 bar runout 양정비
PUMP_AUTO_ZERO_HEAD_FLOW_RATIO: float = 2.0506659273218872  # 0 bar runout 유량비
PUMP_AUTO_FLOW_RATIOS: tuple[float, ...] = (
    0.0, 0.3, 0.6, 0.8, 1.0, 1.2, 1.35, 1.5, 1.75, PUMP_AUTO_ZERO_HEAD_FLOW_RATIO
)
PUMP_CURVE_MIN_POINTS: int = 6
PUMP_CURVE_MAX_POINTS: int = 10
# 부동소수점 비교 오차 흡수용 epsilon (공학 기준 완화가 아닌 float 오차 전용)
PUMP_CURVE_EPS: float = 1e-4  # 4자리 반올림 최대 오차(5e-5 bar) × 2배 여유

# 표시·스핀박스·내부 저장 정밀도 SSOT.
# 표시(_fmt_num)·검증(validate 5행)·라이브러리 매칭·수정됨 표시·저장 정규화가
# 모두 이 상수들을 같이 참조한다. 숫자를 각 파일에 흩뿌리면 SSOT가 깨진다.
PUMP_FLOW_DISPLAY_DECIMALS: int = 1   # 정격유량 스핀박스·화면 표시 자리
PUMP_HEAD_DISPLAY_DECIMALS: int = 2   # 정격압력 스핀박스·화면 표시 자리
PUMP_CURVE_STORE_DECIMALS: int = 4    # 테이블 내부 저장 자리 (표시 역류 방지용)

# power-law 지수 C: 기준 3점 (0,1.20), (1.0,1.0), (1.5,0.65) 에서 산출
# y = 1.20 - 0.20 * x^C
# C = log((1.20-1.00)/(1.20-0.65)) / log(1.00/1.50)
_PUMP_CURVE_EXPONENT: float = math.log(
    (PUMP_AUTO_SHUTOFF_HEAD_RATIO - 1.0) /
    (PUMP_AUTO_SHUTOFF_HEAD_RATIO - PUMP_AUTO_OVERLOAD_HEAD_RATIO)
) / math.log(1.0 / PUMP_AUTO_OVERLOAD_FLOW_RATIO)


def pump_value_eq(left, right, *, flow: bool) -> bool:
    """두 펌프 수치값이 화면 표시 정밀도에서 같은지 판정한다.

    SSOT 비교 함수 — 표시(`_fmt_num`)·5행 검증·라이브러리 매칭·수정됨 표시가
    모두 같은 기준(표시 정밀도)을 쓰도록 한다. "화면에 같아 보이면 같다"를
    보장하므로, 스핀박스 양자화(유량 0.1)와 내부 4자리 저장 사이의 미세 차이를
    에러로 오판하지 않는다.

    Args:
        left, right: 비교할 값 (float 변환 가능해야 함)
        flow: True → 유량(소수 1자리 기준), False → 양정/압력(소수 2자리 기준)

    Returns:
        표시 정밀도에서 같으면 True, 다르거나 변환 불가면 False
    """
    try:
        a = float(left)
        b = float(right)
    except (TypeError, ValueError):
        return False
    decimals = PUMP_FLOW_DISPLAY_DECIMALS if flow else PUMP_HEAD_DISPLAY_DECIMALS
    return f"{a:.{decimals}f}" == f"{b:.{decimals}f}"


def calculate_pump_curve_points(rated_q: float, rated_p: float) -> dict:
    """
    펌프 정격 사양으로부터 특성 곡선 3점을 자동 산출합니다.

    표준 근사식:
      - 체절압(Shutoff)  = 정격압 × 1.2  (운전점 120%)
      - 최대유량(Peak Q) = 정격유량 × 1.5  (운전점 150%)
      - 최대유량점 압력  = 정격압 × 0.65  (운전점 65%)

    Args:
        rated_q: 정격유량 (L/min)
        rated_p: 정격압력 (bar)

    Returns:
        dict with keys: shutoff_p, peak_q, peak_p
    """
    return {
        "shutoff_p": rated_p * 1.2,
        "peak_q":    rated_q * 1.5,
        "peak_p":    rated_p * 0.65,
    }


def validate_pump_params(
    rated_q: float, rated_p: float, shutoff_p: float,
    peak_q: float, peak_p: float,
) -> list[PumpCurveValidationError] | None:
    """
    펌프 커브 5개 파라미터의 공학적 유효성을 검증한다.

    Returns:
        유효하면 None, 위반 시 오류 코드 목록
    """
    errors: list[PumpCurveValidationError] = []

    for name_key, val in [
        (PUMP_PARAM_RATED_Q, rated_q), (PUMP_PARAM_RATED_P, rated_p),
        (PUMP_PARAM_SHUTOFF_P, shutoff_p), (PUMP_PARAM_PEAK_Q, peak_q),
        (PUMP_PARAM_PEAK_P, peak_p),
    ]:
        if val <= 0:
            errors.append(
                PumpCurveValidationError(PUMP_ERR_PARAM_NONPOSITIVE, {"name": name_key})
            )
    if errors:
        return errors

    if shutoff_p < rated_p:
        errors.append(PumpCurveValidationError(PUMP_ERR_SHUTOFF_BELOW_RATED))
    if rated_p <= peak_p:
        errors.append(PumpCurveValidationError(PUMP_ERR_RATED_BELOW_PEAK_P))
    if rated_q >= peak_q:
        errors.append(PumpCurveValidationError(PUMP_ERR_PEAK_Q_BELOW_RATED))

    return errors if errors else None


def calculate_nozzle_k_factor(nozzle_item: dict) -> float:
    """
    노즐 라이브러리 항목으로부터 K-factor(SI)를 산출합니다.

    노즐 타입:
      - 'FIXED': 이미 K_val 필드에 K값이 고정 저장됨
      - 'CALC' : P_bar · Q_lpm 실측값에서 역산: K = Q / √P

    Args:
        nozzle_item: 노즐 라이브러리 단일 항목 dict
            - 'K_val': float  (FIXED 타입의 K값)
            - 'type':  'FIXED' | 'CALC'
            - 'P_bar': float  (CALC 타입 — 실측 압력)
            - 'Q_lpm': float  (CALC 타입 — 실측 유량)

    Returns:
        K-factor (float, SI 단위)
    """
    k_val = float(nozzle_item.get("K_val", 0.0))
    ntype = nozzle_item.get("type", "FIXED")
    if ntype == "CALC":
        p = float(nozzle_item.get("P_bar", 0.0))
        q = float(nozzle_item.get("Q_lpm", 0.0))
        if p > 0:
            k_val = q / (p ** 0.5)
    return k_val


def get_safe_pump_params(meta_data: dict) -> list:
    """
    펌프 성능 곡선 3점(체절, 정격, 최대)을 안전 보정하여 반환합니다.

    Returns:
        [(Q, H), (Q, H), (Q, H)]  — 체절, 정격, 최대 운전점
    """
    try:
        q_rated = float(meta_data.get("rated_q", 1000.0))
    except (TypeError, ValueError):
        q_rated = 1000.0
    try:
        h_rated = float(meta_data.get("rated_p", 10.0))
    except (TypeError, ValueError):
        h_rated = 10.0
    try:
        h_shutoff = float(meta_data.get("shutoff_p", 12.0))
    except (TypeError, ValueError):
        h_shutoff = 12.0
    try:
        q_peak = float(meta_data.get("peak_q", 1500.0))
    except (TypeError, ValueError):
        q_peak = 1500.0
    try:
        h_peak = float(meta_data.get("peak_p", 6.5))
    except (TypeError, ValueError):
        h_peak = 6.5

    if h_shutoff <= h_rated or h_shutoff <= 0.1:
        h_shutoff = max(12.0, h_rated * 1.2)
    if q_rated <= 1.0:
        q_rated = 1000.0
    if q_peak <= q_rated:
        q_peak = q_rated * 1.5
    if h_peak >= h_rated:
        h_peak = h_rated * 0.65

    return [
        (0.0, h_shutoff),
        (q_rated, h_rated),
        (q_peak, h_peak),
    ]


# ---------------------------------------------------------------------------
# 10점 커브 생성 / 정규화 / 검증 / 엔진 전달 함수
# ---------------------------------------------------------------------------

def calculate_pump_curve_points_10(
    rated_q: float,
    rated_p: float,
) -> list[dict[str, float]]:
    """
    정격유량/정격양정으로부터 자동 10점 펌프 커브를 생성합니다.

    정규화 공식:
        x = Q / rated_q
        y = H / rated_p
        y = 1.20 - 0.20 * x ^ C
        C = log((1.20-1.00)/(1.20-0.65)) / log(1.00/1.50) ≈ 2.4949

    마지막 점(runout)은 power-law가 0 bar에 도달하는 지점입니다.
    저장/표시 데이터에는 head_bar = 0.0 을 그대로 유지합니다.

    Args:
        rated_q: 정격유량 (L/min), > 0 필수
        rated_p: 정격양정 (bar), > 0 필수

    Returns:
        list[dict] — 각 항목에 "flow_lpm", "head_bar" 키 포함 (10개)

    Raises:
        ValueError: rated_q 또는 rated_p 가 0 이하일 때
    """
    if rated_q <= 0:
        raise PumpCurveValidationError(PUMP_ERR_RATED_Q_NONPOSITIVE, {"val": rated_q})
    if rated_p <= 0:
        raise PumpCurveValidationError(PUMP_ERR_RATED_P_NONPOSITIVE, {"val": rated_p})

    C = _PUMP_CURVE_EXPONENT
    points: list[dict[str, float]] = []
    for ratio in PUMP_AUTO_FLOW_RATIOS:
        q = ratio * rated_q
        if ratio == PUMP_AUTO_ZERO_HEAD_FLOW_RATIO:
            h = 0.0
        else:
            y = PUMP_AUTO_SHUTOFF_HEAD_RATIO - 0.20 * (ratio ** C)
            h = round(y * rated_p, 4)
            if h < 0.0:
                h = 0.0
        points.append({"flow_lpm": round(q, 4), "head_bar": h})
    return points


def normalize_pump_curve_data(raw_points: object) -> list[dict[str, float]]:
    """
    저장 데이터 또는 UI 입력 데이터를 표준 형식으로 변환합니다.

    허용 입력:
        - list[dict]  — {"flow_lpm": ..., "head_bar": ...}
        - list[tuple] — (flow, head) 또는 [flow, head]

    반환:
        list[dict[str, float]]  — 입력 순서 유지

    금지:
        - 유량 자동 정렬 금지
        - 빈 값/잘못된 값의 임의 대체 금지
        - 변환 불가 값은 그대로 pass-through (TypeError 등은 검증 함수에서 처리)
    """
    if not isinstance(raw_points, list):
        return []
    result: list[dict[str, float]] = []
    for item in raw_points:
        if isinstance(item, dict):
            result.append({
                "flow_lpm": item.get("flow_lpm", item.get("flow", float("nan"))),
                "head_bar": item.get("head_bar", item.get("head", float("nan"))),
            })
        elif isinstance(item, (tuple, list)) and len(item) >= 2:
            result.append({"flow_lpm": item[0], "head_bar": item[1]})
        else:
            result.append({"flow_lpm": float("nan"), "head_bar": float("nan")})
    return result


def _interp_head_at_flow(
    points: list[dict[str, float]],
    target_flow: float,
) -> float | None:
    """
    포인트 목록에서 target_flow 지점의 양정을 선형 보간합니다.

    - 정확히 일치하는 포인트가 있으면 그 값을 반환합니다.
    - target_flow 가 커브 범위(마지막 포인트 유량) 밖이면 None 을 반환합니다.
    """
    if not points:
        return None
    flows = [p["flow_lpm"] for p in points]
    heads = [p["head_bar"] for p in points]

    if target_flow > flows[-1]:
        return None

    for i, (q0, q1) in enumerate(zip(flows[:-1], flows[1:])):
        if q0 <= target_flow <= q1:
            if q1 == q0:
                return heads[i]
            t = (target_flow - q0) / (q1 - q0)
            return heads[i] + t * (heads[i + 1] - heads[i])

    # 첫 번째 포인트 이전(target_flow <= flows[0])
    if target_flow <= flows[0]:
        return heads[0]
    return None


def validate_pump_curve_data(
    points: list[dict[str, float]],
    rated_q: float,
    rated_p: float,
) -> PumpCurveValidationError | None:
    """
    6~10점 펌프 커브 데이터의 공학적 유효성을 검증합니다.

    Args:
        points:  normalize_pump_curve_data() 반환값 형식
        rated_q: 정격유량 (L/min)
        rated_p: 정격양정 (bar)

    Returns:
        유효하면 None, 위반 시 오류 코드
    """
    # rated_q / rated_p 기본 검증
    try:
        rated_q = float(rated_q)
        rated_p = float(rated_p)
    except (TypeError, ValueError):
        return PumpCurveValidationError(PUMP_ERR_INVALID_RATED)
    if rated_q <= 0:
        return PumpCurveValidationError(PUMP_ERR_RATED_Q_NONPOSITIVE, {"val": rated_q})
    if rated_p <= 0:
        return PumpCurveValidationError(PUMP_ERR_RATED_P_NONPOSITIVE, {"val": rated_p})

    # 포인트 수
    n = len(points)
    if n < PUMP_CURVE_MIN_POINTS:
        return PumpCurveValidationError(
            PUMP_ERR_TOO_FEW_POINTS,
            {"n": n, "min": PUMP_CURVE_MIN_POINTS},
        )
    if n > PUMP_CURVE_MAX_POINTS:
        return PumpCurveValidationError(
            PUMP_ERR_TOO_MANY_POINTS,
            {"n": n, "max": PUMP_CURVE_MAX_POINTS},
        )

    # 각 포인트 숫자 변환 및 음수 양정 검사
    flows: list[float] = []
    heads: list[float] = []
    for i, pt in enumerate(points):
        try:
            q = float(pt["flow_lpm"])
            h = float(pt["head_bar"])
        except (TypeError, ValueError, KeyError):
            return PumpCurveValidationError(PUMP_ERR_INVALID_POINT_VALUE, {"idx": i + 1})
        if math.isnan(q) or math.isnan(h):
            return PumpCurveValidationError(PUMP_ERR_INVALID_POINT_VALUE, {"idx": i + 1})
        if h < 0:
            return PumpCurveValidationError(
                PUMP_ERR_NEGATIVE_HEAD,
                {"idx": i + 1, "h": h},
            )
        flows.append(q)
        heads.append(h)

    # 첫 유량 = 0
    if flows[0] != 0.0:
        return PumpCurveValidationError(
            PUMP_ERR_FIRST_FLOW_NOT_ZERO,
            {"flow": flows[0]},
        )

    # 0 bar 포인트가 마지막이 아닌 곳에 있으면 거부
    for i in range(len(heads) - 1):
        if heads[i] == 0.0:
            return PumpCurveValidationError(PUMP_ERR_ZERO_HEAD_NOT_LAST, {"idx": i + 1})

    # 유량 strictly increasing
    for i in range(1, len(flows)):
        if flows[i] <= flows[i - 1]:
            return PumpCurveValidationError(
                PUMP_ERR_FLOW_NOT_INCREASING,
                {"idx": i + 1, "flow": flows[i], "prev": flows[i - 1]},
            )

    # 양정 non-increasing (같거나 낮아져야 함)
    for i in range(1, len(heads)):
        if heads[i] > heads[i - 1]:
            return PumpCurveValidationError(
                PUMP_ERR_HEAD_RISING,
                {"idx": i + 1, "head": heads[i], "prev": heads[i - 1]},
            )

    # 체절양정 <= rated_p * 1.40
    shutoff_h = heads[0]
    max_shutoff = rated_p * 1.40
    if shutoff_h > max_shutoff:
        return PumpCurveValidationError(
            PUMP_ERR_SHUTOFF_TOO_HIGH,
            {"shutoff": shutoff_h, "max_shutoff": max_shutoff},
        )

    # 150% rated_q 지점 보간 검증
    target_150 = rated_q * PUMP_AUTO_OVERLOAD_FLOW_RATIO
    norm_pts = [{"flow_lpm": q, "head_bar": h} for q, h in zip(flows, heads)]
    h_at_150 = _interp_head_at_flow(norm_pts, target_150)
    if h_at_150 is None:
        return PumpCurveValidationError(
            PUMP_ERR_150PCT_OUT_OF_RANGE,
            {"target": f"{target_150:.1f}", "last": f"{flows[-1]:.1f}"},
        )
    # 150% 지점은 실무 입력/반올림/보간 미세 차이를 감안해 65% 기준에서 5%p 낮은
    # 60%까지 허용한다. 자동 생성 기준값(65%) 자체는 그대로 유지한다.
    min_h_at_150 = rated_p * (
        PUMP_AUTO_OVERLOAD_HEAD_RATIO - PUMP_OVERLOAD_HEAD_TOLERANCE_RATIO
    )
    if h_at_150 + PUMP_CURVE_EPS < min_h_at_150:
        return PumpCurveValidationError(
            PUMP_ERR_150PCT_HEAD_TOO_LOW,
            {
                "target": f"{target_150:.1f}",
                "h": f"{h_at_150:.4f}",
                "min_h": f"{min_h_at_150:.4f}",
            },
        )

    return None


def extend_curve_to_zero_head(
    points: list[tuple[float, float]],
    rated_q: float,
) -> list[tuple[float, float]]:
    """
    마지막 두 점의 직선을 연장해 양정 0 bar 지점을 산출합니다.

    - 마지막 점이 이미 0 bar 이면 입력을 그대로 반환합니다.
    - 마지막 구간이 수평이면 PUMP_AUTO_ZERO_HEAD_FLOW_RATIO × rated_q 를 사용하고,
      그 값이 마지막 유량 이하이면 마지막 유량 × 1.05 로 대체합니다.
    """
    if not points:
        return []

    result = list(points)
    last_q, last_h = result[-1]
    if last_h <= PUMP_CURVE_EPS:
        return result

    if len(result) >= 2:
        q0, h0 = result[-2]
        q1, h1 = result[-1]
    else:
        q0, h0 = result[-1]
        q1, h1 = result[-1]

    try:
        rq = float(rated_q)
    except (TypeError, ValueError):
        rq = 0.0

    if abs(h1 - h0) < PUMP_CURVE_EPS:
        ext_q = rq * PUMP_AUTO_ZERO_HEAD_FLOW_RATIO if rq > 0 else q1 * 1.05
        if ext_q <= q1 + PUMP_CURVE_EPS:
            ext_q = q1 * 1.05
    elif abs(q1 - q0) < PUMP_CURVE_EPS:
        ext_q = q1 * 1.05
    else:
        slope = (h1 - h0) / (q1 - q0)
        if abs(slope) < PUMP_CURVE_EPS:
            ext_q = rq * PUMP_AUTO_ZERO_HEAD_FLOW_RATIO if rq > 0 else q1 * 1.05
            if ext_q <= q1 + PUMP_CURVE_EPS:
                ext_q = q1 * 1.05
        else:
            ext_q = q1 - h1 / slope

    if ext_q <= q1 + PUMP_CURVE_EPS:
        ext_q = q1 * 1.05

    result.append((ext_q, 0.0))
    return result


def get_pump_curve_points_for_engine(
    meta_data: dict,
) -> list[tuple[float, float]]:
    """
    계산 엔진(EPANET)에 넘길 (flow_lpm, head_bar) 목록을 반환합니다.

    우선순위:
        1. meta_data["pump_curve_data"] 가 존재하고 비어 있지 않으면 반드시 검증 후 사용
           - 검증 실패 시 조용히 fallback하지 않고 ValueError 를 발생시킨다 (하드 거부 정책)
        2. pump_curve_data 가 없거나 빈 경우에만 rated_q / rated_p 로 자동 10점 생성

    Args:
        meta_data: 노드 메타 dict

    Returns:
        list[tuple[float, float]] — (flow_lpm, head_bar)

    Raises:
        ValueError: pump_curve_data 가 존재하지만 검증을 통과하지 못할 때
    """
    raw = meta_data.get("pump_curve_data")
    if raw:  # pump_curve_data 가 존재하고 비어 있지 않음
        pts = normalize_pump_curve_data(raw)
        err = validate_pump_curve_data(
            pts,
            meta_data.get("rated_q", 0),
            meta_data.get("rated_p", 0),
        )
        if err is not None:
            raise err
        rated_q = float(meta_data.get("rated_q", 0) or 0)
        tuples = [(p["flow_lpm"], p["head_bar"]) for p in pts]
        return extend_curve_to_zero_head(tuples, rated_q)

    # pump_curve_data 없음 또는 빈 값 → 자동 10점 fallback
    try:
        q = float(meta_data.get("rated_q", 1000.0))
    except (TypeError, ValueError):
        q = 1000.0
    try:
        p = float(meta_data.get("rated_p", 10.0))
    except (TypeError, ValueError):
        p = 10.0

    if q <= 0:
        q = 1000.0
    if p <= 0:
        p = 10.0

    pts_10 = calculate_pump_curve_points_10(q, p)
    return [(pt["flow_lpm"], pt["head_bar"]) for pt in pts_10]


def get_pump_curve_points_for_plot(
    meta_data: dict,
    fallback_points: list[tuple[float, float]] | None = None,
) -> list[tuple[float, float]]:
    """
    화면/보고서 펌프커브 표시용 point list를 반환합니다.

    엔진 입력과 같은 get_pump_curve_points_for_engine()을 우선 사용하고,
    표시 단계에서만 필요한 방어 fallback을 한 곳에 모읍니다.
    """
    try:
        points = get_pump_curve_points_for_engine(meta_data)
    except Exception:
        if fallback_points is None:
            raise
        points = fallback_points

    clean: list[tuple[float, float]] = []
    for q, h in points:
        qf = float(q)
        hf = float(h)
        if qf >= 0.0:
            clean.append((qf, hf))
    clean.sort(key=lambda p: p[0])

    if len(clean) < 2 and fallback_points is not None:
        clean = [(float(q), float(h)) for q, h in fallback_points if float(q) >= 0.0]
        clean.sort(key=lambda p: p[0])

    return clean


def interpolate_pump_head_at_flow(
    points: list[tuple[float, float]],
    flow_lpm: float,
) -> float:
    """
    EPANET multi-point pump curve와 같은 선형 보간으로 head(bar)를 반환합니다.
    범위 밖은 첫/마지막 점의 head로 고정합니다.
    """
    if not points:
        return 0.0

    q = float(flow_lpm)
    pts = sorted((float(pq), float(ph)) for pq, ph in points)

    if q <= pts[0][0]:
        return max(pts[0][1], 0.0)
    if q >= pts[-1][0]:
        return max(pts[-1][1], 0.0)

    for (q0, h0), (q1, h1) in zip(pts[:-1], pts[1:]):
        if q0 <= q <= q1:
            if q1 == q0:
                return max(h0, 0.0)
            t = (q - q0) / (q1 - q0)
            return max(h0 + t * (h1 - h0), 0.0)

    return max(pts[-1][1], 0.0)


def summarize_pump_curve_fields(
    points: list[dict[str, float]],
    rated_q: float,
    rated_p: float,
) -> dict[str, float]:
    """
    pump_curve_data 저장과 함께 기존 대표 필드를 동기화하기 위한 값을 산출합니다.

    반환:
        {
            "rated_q":  rated_q,
            "rated_p":  rated_p,
            "shutoff_p": points[0]["head_bar"],
            "peak_q":   1.5 * rated_q,
            "peak_p":   150% 지점 선형 보간 양정,
        }

    150% 지점이 커브 범위 밖이면 peak_p = rated_p * 0.65 를 기본값으로 사용합니다.
    """
    shutoff_p = float(points[0]["head_bar"]) if points else rated_p * 1.2
    peak_q = rated_q * PUMP_AUTO_OVERLOAD_FLOW_RATIO
    target_150 = peak_q
    h_at_150 = _interp_head_at_flow(points, target_150)
    peak_p = h_at_150 if h_at_150 is not None else rated_p * PUMP_AUTO_OVERLOAD_HEAD_RATIO

    return {
        "rated_q":  float(rated_q),
        "rated_p":  float(rated_p),
        "shutoff_p": shutoff_p,
        "peak_q":   peak_q,
        "peak_p":   float(peak_p),
    }


def normalize_pipe_type_label(text) -> str:
    """배관 라벨 정규화 (예: 'Sch40 DN25' -> 'Sch40')"""
    if text is None:
        return ""
    s = str(text).strip()
    if not s:
        return ""

    s_up = s.upper()
    cut_pos = -1

    idx = s_up.find(" DN")
    if idx != -1:
        cut_pos = idx
    if cut_pos == -1:
        idx = s_up.find("DN")
        if idx != -1 and idx + 2 < len(s) and s_up[idx + 2].isdigit():
            cut_pos = idx
    if cut_pos == -1:
        idx = s_up.find("(DN")
        if idx != -1:
            cut_pos = idx

    if cut_pos != -1:
        return s[:cut_pos].strip()
    return s
