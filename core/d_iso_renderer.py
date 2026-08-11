# -*- coding: utf-8 -*-
"""모듈 D — D3: ISO 도면을 벡터 PDF 로 그린다.

D1 이 읽은 화면 좌표와 D2 가 결합한 수치를 받아, 표시 항목 2 종(관로·노드)을 골라
A4 한 장을 그린다. 계산은 하지 않는다 — 값이 없으면 빈칸으로 두고 리포트에 올린다
(지시서 7-4).

**페이지** — 참조 코퍼스의 PIPENET 원본 ISO PDF 583 장이 전량 1 쪽 A4 이고 머리글도
"Page 1 of 1" 이다. 즉 분할 규칙이 없다. 망이 아무리 커도 한 장에 맞춰 축척한다.
세로/가로만 고르며, 코퍼스 분포는 세로 505 · 가로 78 이다.

**프리셋** — 같은 코퍼스의 범례를 역산한 결과다 (설비 종류로 완전히 갈리며 교차 0 건).
    유량본            관로 Pipe vol. flow      노드 없음            268 장
    압력본(스프링클러) 관로 Pipe press. diff.   노드 Pressure       134 장
    압력본(옥내소화전) 관로 Pipe bore           노드 Pressure       140 장

**색** — PIPENET 의 색을 재현하지 않는다. SDF 의 스킴 요소는 코퍼스 4690 개 전량이
자식 없는 ``auto-classify="1"`` 뿐이라 밴드 경계도 색도 파일에 없다. 아래 6 단계
등간격 분류와 팔레트는 우리 것이며, 지시서 7-2 가 자체 서식을 요구하므로 그래도 된다.
"""
from __future__ import annotations

import math
import statistics
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Sequence

import matplotlib
from matplotlib import font_manager
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.backends.backend_pdf import FigureCanvasPdf
from matplotlib.collections import EllipseCollection, LineCollection
from matplotlib.figure import Figure
from matplotlib.font_manager import FontProperties
from matplotlib.lines import Line2D
from matplotlib.patches import Polygon, Rectangle

from core.d_display_model import DisplayModel, DisplayNode, DisplayPipe, UnitSpec
from core.d_label_layout import LabelRequest, lay_out
from core.d_result_binder import BoundModel, normalize_label

# A4. 지시서 밖의 종이를 쓰지 않는다 — 코퍼스 583 장이 전량 A4 다.
A4_MM = (210.0, 297.0)
_MM_PER_INCH = 25.4
# 시각 회귀 비교용 고정 해상도. A4 세로가 1240×1754 px 로 떨어진다.
PNG_DPI = 150

# 한 장에 6 단계. 코퍼스 범례가 전량 6 단계였다.
BAND_COUNT = 6
# 등간격 눈금 후보. 밴드 경계를 읽기 좋은 수로 떨어뜨리기 위한 것이고, 값 자체를
# 반올림하지는 않는다.
_NICE_STEPS = (1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0, 7.5)
# 낮은 값 → 높은 값. 참조 도면 40 장 중 범례가 있는 35 장이 전부 이 여섯 색을 이
# 순서로 쓴다. 관로 획에서도 같은 여섯 색만 나온다.
BAND_COLOURS = ("#ff0000", "#ffac00", "#00ff00", "#00ffff", "#0000ff", "#ff00ff")
# 노드 값은 같은 여섯 색을 쓰지 않는다. 압력 도면 40 장 중 39 장이 노드 범례에 이
# 무채색 계단을 쓴다 — 검정에서 시작해 한 칸에 42/255 씩 밝아진다.
NODE_BAND_COLOURS = ("#000000", "#2a2a2a", "#545454", "#7e7e7e", "#a9a9a9", "#d4d4d4")
_BLANK_COLOUR = "#9a9a9a"

# 값 글씨 높이. 종이가 아니라 모델 좌표에 고정이다 — 참조 40 장에서 모델 단위
# 변동계수 0.029, 같은 값을 종이 pt 로 보면 0.172, 종이 비율로 보면 0.286 이다.
# SDF 의 Label-display/@print-font 는 이 높이를 설명하지 못한다: 값이 40 인 한 장과
# 36 인 39 장의 실측 높이가 24.168 로 같다.
_LABEL_UNITS = 24.17

# 도면 주기 높이. 같은 40 장에서 22.56 모델 단위(변동계수 0.011)다. 다만 코퍼스
# 334 개 SDF 의 typesize 가 8189 개 글자 전부 30 이라, typesize 가 달라지면 어떻게
# 되는지는 재지 못했다. 이름대로 크기에 비례한다고 보고 30 을 기준으로 늘리되,
# 근거가 있는 건 30 하나뿐이다 — 우리 손작업 SDF 는 60 을 쓴다.
_NOTE_UNITS = 22.56
_NOTE_TYPESIZE = 30.0

# 지시선. 망보다 연하고 가늘어야 관로로 오독되지 않는다. 값이 없어 회색으로 그린
# 관로(_BLANK_COLOUR)와는 색이 달라야 한다 — 검사가 둘을 색으로 갈라 본다.
_LEADER_COLOUR = "#7f7f7f"
# 굵기는 배관(_PIPE_WIDTH_UNITS = 1.0)의 절반이다. 종이 pt 로 묶여 있던 0.25pt 는
# 참조 배율 폭(0.089~0.167 pt/단위) 어디에서도 배관보다 1.5~2.8 배 굵어서, 바로 위
# 주석이 말하는 "가늘어야 한다" 를 한 번도 지키지 못했다.
_LEADER_WIDTH_UNITS = 0.5

# 기기·특수기기 기호. 지시서 7-2 에 따라 우리 표기라 PIPENET 에 대응하는 실측이 없다.
# 다만 크기를 매다는 기준은 망과 같아야 한다 — 종이 pt 로 묶어 두면 같은 기호가 참조
# 배율 폭에서 노드 점의 2.1 배부터 4.0 배까지 널뛴다. 아래는 크기를 새로 정한 것이
# 아니라 지금 우리 도면의 생김새(3.4/2.6/0.7/1.4 pt @ 0.185 pt/단위)를 옮긴 값이다.
# 라벨이 피해야 할 자리를 잡는 데도 쓰이므로 그리는 쪽과 한 값이어야 한다.
_DEVICE_UNITS = 18.4
_EQUIPMENT_UNITS = 14.1
_MARKER_EDGE_UNITS = 3.8        # 기호 테두리
_DEVICE_LINK_UNITS = 7.6        # 기기를 관로에 잇는 선

# 노즐 스텁. SDF 가 @/n 좌표를 입력노드와 같은 자리에 두면 헤드가 분기점 위에 겹쳐
# 찍힌다. 원본은 고칠 수 없으므로(지시서 7-3) 그릴 때 방향을 유도한다. 근거는 참조
# 코퍼스 SDF 334 개 / 노즐 3437 개 실측이다 — 입사 관로가 있는 3115 개 중 3079 개
# (98.8%)가 관로의 연장선이고 수직은 0 건이었다.
#
# 길이는 도면 크기를 따라가지 않는다. 같은 표본을 네 기준으로 재면 변동계수가
# 모델 단위 0.218 · 관로 길이 대비 0.294 · 헤드 간격 대비 0.493 · 도면 span 대비
# 0.800 이다. 헤드 간격이 3.8 배 벌어져도 스텁은 1.5 배밖에 안 변하니, 축척이 아니라
# 니플 규격처럼 고정 치수에 가깝다. 실제로 값이 26 종뿐이고 55~62 와 86~88 두 무리로
# 갈린다 — 아래 상수는 그중 단일 최빈값(3437 개 중 1087 개)이다.
_STUB_UNITS = 58.0
# 유도한 것은 실측과 같은 잉크로 그리지 않는다 — 점선 + 이 색으로 갈라 보인다.
_DERIVED_COLOUR = "#c2410c"

# 흐름 화살표. PIPENET 이 직접 출력한 ISO PDF 255 장을 SDF 모델 좌표에 맞춰 실측한
# 값이다 — 채운 삼각형이 아니라 획 두 개짜리 열린 갈매기표이고, 크기는 도면 크기가
# 아니라 모델 좌표에 붙어 있다(날개 6.76 단위, 변동계수 0.029). 벌어진각 53.13° 는
# 2·atan(½) 이고, 자리는 관로 호길이의 0.687 지점, 꼭짓점은 4902/4902 가 진행 방향.
_ARROW_WING_UNITS = 6.76
_ARROW_HALF_ANGLE = math.degrees(math.atan(0.5))
_ARROW_AT = 0.687

# 노즐 머리. 같은 참조 PDF 60 장 / 기호 638 개 실측이다 — 채운 삼각형이 아니라 속이
# 빈 삼각형 윤곽이고(채움 0/638), 꼭짓점이 `@` 노드 좌표에 정확히 앉는다(그린 길이 ÷
# 모델 길이 1.0004, 변동계수 0.004). 스텁은 꼭짓점이 아니라 삼각형 밑변에서 끝난다.
# 크기는 화살표와 마찬가지로 모델 좌표에 고정이다(길이 변동계수 0.023, 반폭 0.011).
_HEAD_LENGTH_UNITS = 17.96
_HEAD_HALF_WIDTH_UNITS = 10.01

# 노드 점. 같은 60 장 / 점 3524 개 실측이다. 여기서도 기준은 종이가 아니라 모델
# 좌표다 — 지름을 종이 pt 로 재면 변동계수가 0.420 인데 모델 단위로 재면 0.019 다.
_NODE_DOT_UNITS = 9.59

# 관로 획 굵기. 참조 80 장은 관경을 6~8 종류씩 쓰면서도 획 굵기는 한 장에 한 종류만
# 쓴다(중앙 1 종, 최대 2 종) — PIPENET 은 선 굵기로 관경을 나타내지 않는다. 그 하나의
# 굵기도 종이가 아니라 모델 좌표에 고정이다(모델 단위 변동계수 0.006, 종이 pt 0.405).
_PIPE_WIDTH_UNITS = 1.0

# 판면. 참조 지면(595×842 pt)을 실측한 양식이다. 테두리가 네 변 8mm 안쪽에 있고
# 표제란은 그 테두리의 오른쪽 아래 모서리에 딱 붙는 128×35mm 상자이며, 행은 위에서
# 5/5/5/10/10mm 다. 망은 그 위를 다 쓴다. 아래 값은 종이 모서리에서 잰 pt 이므로
# A4 두 방향에 그대로 얹힌다.
_FRAME_INSET_PT = 22.68         # 테두리, 종이 네 모서리에서 (8mm)
_RULE_PT = 0.06                 # 테두리·표제란 괘선 굵기
_BLOCK_WIDTH_PT = 362.64        # 128mm
_BLOCK_ROW_PT = (14.22, 14.16, 14.16, 28.32, 28.26)     # 위에서부터 5/5/5/10/10mm
_BLOCK_INSET_PT = 2.82          # 괘선에서 글씨까지 (1mm)
_BLOCK_FONT_PT = 8.5
_FOOTNOTE_PT = 4.6              # 각주는 우리 것이라 참조에 대응하는 글줄이 없다
# 망이 채우는 세로 구간. 표제란이 종이 모서리에 pt 로 붙어 있으므로 이쪽도 pt 여야
# 한다 — 종이 비율로 두면 A4 가로에서 망 아래끝이 표제란 위(121.8pt)로 내려온다.
_PLATE_BOTTOM_PT = 131.33       # 종이 아래에서. 표제란 위에서 9.5pt 뜬다
_PLATE_TOP_PT = 58.93           # 종이 위에서

# 범례. 한 벌이 3 칸 × 2 줄 격자이고 왼쪽에 항목 이름이 두 줄로 붙는다. 값이 큰
# 쪽으로 왼쪽 위에서 오른쪽 아래로 읽는다. 노드·관로 두 벌이 함께 나오면 노드가 위다.
_LEGEND_BOTTOM_PT = 28.8        # 맨 아래 벌의 아랫줄 바닥
_LEGEND_LANE_PT = 28.3          # 벌과 벌 사이
_LEGEND_ROW_PT = 11.4
_LEGEND_COL_PT = 90.6
_LEGEND_SWATCH_DX_PT = 87.8     # 표제란 왼쪽 끝에서 첫 칸까지
_LEGEND_SWATCH_PT = 8.5
_LEGEND_LABEL_DX_PT = 11.3      # 칸에서 그 칸 설명까지
_LEGEND_FONT_PT = 7.0           # 참조에 범례가 없어 표제란 글씨 크기를 따르지 않는다

_KOREAN_FONTS =("malgun.ttf", "malgunsl.ttf", "NanumGothic.ttf", "gulim.ttc", "batang.ttc")

# 장치 링크 종류별 기호. PIPENET 기호를 베끼지 않고 우리 표기를 쓴다 (지시서 7-2).
DEVICE_MARKS = {"Pump-fan": "o", "Elastomeric-valve": "D", "Pressure-loss": "s"}


# ── 표시 항목 ───────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class LinkRow:
    """관로 하나에 대해 SDF·XML 이 아는 것을 한자리에 모은 것."""
    pipe: DisplayPipe
    result: Any             # PipeResult | None
    input_row: dict[str, Any]
    result_row: dict[str, Any]


@dataclass(frozen=True)
class NozzleRow:
    nozzle: Any             # DisplayNozzle
    result: Any             # NozzleResult | None


@dataclass(frozen=True)
class NodeRow:
    node: DisplayNode
    result: Any             # NodeResult | None


@dataclass(frozen=True)
class DisplayItem:
    """SDF 드롭다운 한 줄. ``name`` 은 select-name 표기를 그대로 쓴다."""
    name: str
    scope: str                      # 'pipe' | 'nozzle' | 'node'
    quantity: str                   # <Units> 키. '' 이면 무차원·문자
    read: Callable[[Any], Any]
    per_length: bool = False        # 압력구배처럼 길이로 나눈 양
    si_symbol: str = ""             # 파일이 단위를 선언하지 않는 양의 SI 표기


def _bore(row: LinkRow) -> float | None:
    return (row.result.nominal_bore_m if row.result else None) or row.pipe.bore_m


def _length(row: LinkRow) -> float | None:
    return (row.result.length_m if row.result else None) or row.pipe.length_m


def _pressure_drop(row: LinkRow) -> float | None:
    res = row.result
    if res is None or res.inlet_pressure_pa is None or res.outlet_pressure_pa is None:
        return None
    return res.inlet_pressure_pa - res.outlet_pressure_pa


def _gradient(row: LinkRow) -> float | None:
    drop, length = _pressure_drop(row), _length(row)
    return None if drop is None or not length else drop / length


def _mass_flow(row: LinkRow) -> float | None:
    flow = row.result.flow_m3s if row.result else None
    density = row.result_row.get("Density")
    return None if flow is None or not isinstance(density, float) else flow * density


LINK_ITEMS: tuple[DisplayItem, ...] = (
    DisplayItem("None", "pipe", "", lambda r: None),
    DisplayItem("Pipe bore", "pipe", "Diameter", _bore),
    DisplayItem("Pipe length", "pipe", "Length", _length),
    DisplayItem("Pipe type", "pipe", "", lambda r: r.input_row.get("Type")),
    DisplayItem("Pipe volumetric flow", "pipe", "Volumetric-flow",
                lambda r: r.result.flow_m3s if r.result else None),
    DisplayItem("Pipe mass flow", "pipe", "", _mass_flow, si_symbol="kg/s"),
    DisplayItem("Pipe velocity", "pipe", "Velocity",
                lambda r: r.result.velocity_ms if r.result else None),
    DisplayItem("Pipe pressure difference", "pipe", "Pressure", _pressure_drop),
    DisplayItem("Pipe pressure gradient", "pipe", "Pressure", _gradient, per_length=True),
    DisplayItem("Nozzle required flow", "nozzle", "Volumetric-flow",
                lambda r: r.result.required_flow_m3s if r.result else None),
    DisplayItem("Nozzle calculated flow", "nozzle", "Volumetric-flow",
                lambda r: r.result.calculated_flow_m3s if r.result else None),
    DisplayItem("Nozzle calculated pressure", "nozzle", "Pressure",
                lambda r: r.result.inlet_pressure_pa if r.result else None),
    DisplayItem("Nozzle calculated deviation", "nozzle", "",
                lambda r: r.result.deviation_percent if r.result else None, si_symbol="%"),
    DisplayItem("Nozzle type", "nozzle", "",
                lambda r: (r.result.nozzle_type if r.result else "") or r.nozzle.library_item),
    DisplayItem("Tagging", "pipe", "", lambda r: r.pipe.label),
)

NODE_ITEMS: tuple[DisplayItem, ...] = (
    DisplayItem("None", "node", "", lambda r: None),
    DisplayItem("Node elevation", "node", "Length",
                lambda r: (r.result.elevation_m if r.result else None) or r.node.elevation_m),
    DisplayItem("Node pressure", "node", "Pressure",
                lambda r: r.result.pressure_pa if r.result else None),
    DisplayItem("Tagging", "node", "", lambda r: r.node.label),
)

_LINK_BY_NAME = {i.name: i for i in LINK_ITEMS}
_NODE_BY_NAME = {i.name: i for i in NODE_ITEMS}


@dataclass(frozen=True)
class Preset:
    link: str
    node: str


# 참조 코퍼스 542 장의 범례에서 역산한 실사용 조합.
PRESETS: dict[str, Preset] = {
    "유량본": Preset("Pipe volumetric flow", "None"),
    "압력본": Preset("Pipe pressure difference", "Node pressure"),
    "압력본_옥내소화전": Preset("Pipe bore", "Node pressure"),
}


# ── 값 서식 ─────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class ValueFormat:
    """SI → 화면 표기. 계수는 전부 파일의 <Units> 에서 온다 (지시서 7-5)."""
    factor: float | None
    precision: int
    symbol: str

    def display(self, si: Any) -> Any:
        """SI 값을 화면 단위로. 밴드도 라벨도 이 값으로 정한다."""
        if si is None or not isinstance(si, (int, float)) or isinstance(si, bool):
            return si
        return None if self.factor is None else float(si) * self.factor

    def text(self, shown: Any) -> str:
        if shown is None:
            return ""
        if not isinstance(shown, (int, float)) or isinstance(shown, bool):
            return str(shown)
        return f"{shown:.{self.precision}f}"


def _format_for(item: DisplayItem, units: dict[str, UnitSpec],
                warnings: list[str]) -> ValueFormat:
    if not item.quantity:
        return ValueFormat(1.0, 1, item.si_symbol)
    spec = units.get(item.quantity)
    if spec is None or spec.factor is None:
        warnings.append(f"'{item.name}' 의 단위({item.quantity})를 파일이 정하지 않는다 — SI 로 둔다")
        return ValueFormat(1.0, 2, spec.symbol if spec else "")
    if not item.per_length:
        return ValueFormat(spec.factor, spec.precision, spec.symbol)
    length = units.get("Length")
    if length is None or length.factor is None:
        warnings.append(f"'{item.name}' 은 길이 단위가 있어야 한다 — SI 로 둔다")
        return ValueFormat(spec.factor, spec.precision, f"{spec.symbol}/m")
    return ValueFormat(spec.factor / length.factor, spec.precision + 2,
                       f"{spec.symbol}/{length.symbol}")


# ── 밴드 분류 ───────────────────────────────────────────────────────────────


def _nice_step(raw: float) -> float:
    """등간격 폭을 읽기 좋은 눈금으로 올린다."""
    if raw <= 0 or not math.isfinite(raw):
        return 1.0
    decade = 10.0 ** math.floor(math.log10(raw))
    for step in _NICE_STEPS:
        if step * decade >= raw - 1e-12:
            return step * decade
    return 10.0 * decade


def band_edges(values: Sequence[float]) -> tuple[float, ...]:
    """6 단계를 가르는 경계 5 개. 값이 하나뿐이면 빈 튜플(단일 밴드)."""
    usable = [v for v in values if isinstance(v, (int, float)) and math.isfinite(v)]
    if not usable:
        return ()
    lo, hi = min(usable), max(usable)
    if hi - lo <= 0:
        return ()
    step = _nice_step((hi - lo) / BAND_COUNT)
    first = math.ceil(lo / step + 1e-9) * step
    edges = [first + k * step for k in range(BAND_COUNT - 1)]
    return tuple(e for e in edges if e < hi)


def _band_of(value: Any, edges: Sequence[float]) -> int | None:
    if not isinstance(value, (int, float)) or not math.isfinite(value):
        return None
    for i, edge in enumerate(edges):
        if value < edge:
            return i
    return len(edges)


# ── 기하 ────────────────────────────────────────────────────────────────────


def _polyline(pipe: DisplayPipe, coords: dict[str, tuple[float, float]]
              ) -> list[tuple[float, float]] | None:
    a, b = coords.get(pipe.input_node), coords.get(pipe.output_node)
    if a is None or b is None:
        return None
    return [a, *pipe.waypoints, b]


def _point_at(path: Sequence[tuple[float, float]], t: float) -> tuple[float, float]:
    """폴리라인의 호길이 비율 t 지점 — 관로에 달린 특수기기 자리."""
    spans = [math.dist(path[i], path[i + 1]) for i in range(len(path) - 1)]
    total = sum(spans)
    if total <= 0:
        return path[0]
    want = max(0.0, min(1.0, t)) * total
    for i, span in enumerate(spans):
        if want <= span or i == len(spans) - 1:
            k = (want / span) if span else 0.0
            (x0, y0), (x1, y1) = path[i], path[i + 1]
            return x0 + (x1 - x0) * k, y0 + (y1 - y0) * k
        want -= span
    return path[-1]


def _incident_dirs(pipes: Sequence[DisplayPipe], coords: dict[str, tuple[float, float]],
                   base: str) -> list[tuple[float, float]]:
    """base 노드에 붙은 관로가 base 에서 뻗어 나가는 단위벡터들."""
    origin = coords[base]
    out: list[tuple[float, float]] = []
    for pipe in pipes:
        if pipe.input_node == base:
            ahead = [*pipe.waypoints, coords.get(pipe.output_node)]
        elif pipe.output_node == base:
            ahead = [*reversed(pipe.waypoints), coords.get(pipe.input_node)]
        else:
            continue
        for point in ahead:
            if point is None:
                continue
            dx, dy = point[0] - origin[0], point[1] - origin[1]
            dist = math.hypot(dx, dy)
            if dist > 1e-9:
                out.append((dx / dist, dy / dist))
                break
    return out


def _measured_stub(nozzles: dict[str, Any],
                   coords: dict[str, tuple[float, float]]) -> float | None:
    """같은 도면에서 방향이 있는 헤드가 실제로 얼마나 나와 있는가. 없으면 None."""
    measured = []
    for row in nozzles.values():
        base = coords.get(row.nozzle.input_node)
        tip = coords.get(row.nozzle.output_node)
        if base is not None and tip is not None and tip != base:
            measured.append(math.dist(base, tip))
    return statistics.median(measured) if measured else None


def _nozzle_tips(nozzles: dict[str, Any], pipes: Sequence[DisplayPipe],
                 coords: dict[str, tuple[float, float]], stub: float
                 ) -> tuple[dict[str, tuple[float, float]], list[str], list[str]]:
    """헤드 삼각형의 꼭짓점 자리. 원본이 방향을 준 것은 그대로 쓰고, 출력노드가
    입력노드와 겹쳐 방향이 없는 것만 유도한다. 어느 쪽인지는 갈라서 돌려준다 —
    유도한 것을 실측인 척 그리지 않기 위해서다."""
    tips: dict[str, tuple[float, float]] = {}
    derived: list[str] = []
    undirected: list[str] = []
    for label, row in nozzles.items():
        base = coords.get(row.nozzle.input_node)
        if base is None:
            continue
        tip = coords.get(row.nozzle.output_node)
        if tip is not None and tip != base:
            tips[label] = tip
            continue
        dirs = _incident_dirs(pipes, coords, row.nozzle.input_node)
        sx = sum(d[0] for d in dirs)
        sy = sum(d[1] for d in dirs)
        mag = math.hypot(sx, sy)
        if mag <= 1e-9:
            # 입사 관로가 없거나 서로 상쇄된다 — 방향을 정할 근거가 없다.
            # 자리를 지어내지 않고 입력노드에 둔 채 리포트에 올린다.
            undirected.append(label)
            tips[label] = base
            continue
        tips[label] = (base[0] - sx / mag * stub, base[1] - sy / mag * stub)
        derived.append(label)
    return tips, derived, undirected


def _longest_segment(path: Sequence[tuple[float, float]]
                     ) -> tuple[tuple[float, float], float]:
    """가장 긴 구간의 중점과 방향각 — 값 라벨을 놓을 자리다."""
    best, mid, angle = -1.0, path[0], 0.0
    for i in range(len(path) - 1):
        (x0, y0), (x1, y1) = path[i], path[i + 1]
        span = math.hypot(x1 - x0, y1 - y0)
        if span > best:
            best, mid = span, ((x0 + x1) / 2, (y0 + y1) / 2)
            angle = math.degrees(math.atan2(y1 - y0, x1 - x0))
    return mid, angle


def _along(path: Sequence[tuple[float, float]], fraction: float
           ) -> tuple[tuple[float, float], float] | None:
    """폴리라인 호길이의 `fraction` 지점과 그 자리의 진행각."""
    spans = [math.hypot(b[0] - a[0], b[1] - a[1]) for a, b in zip(path, path[1:])]
    total = sum(spans)
    if total <= 0:
        return None
    want, acc = total * fraction, 0.0
    for (a, b), span in zip(zip(path, path[1:]), spans):
        if span and acc + span >= want:
            t = (want - acc) / span
            return ((a[0] + t * (b[0] - a[0]), a[1] + t * (b[1] - a[1])),
                    math.degrees(math.atan2(b[1] - a[1], b[0] - a[0])))
        acc += span
    return None


def _chevron(tip: tuple[float, float], angle: float, wing: float
             ) -> list[list[tuple[float, float]]]:
    """꼭짓점이 `angle` 쪽을 보는 갈매기표 — 획 두 개. 채우지 않는다."""
    back = [angle + 180 - _ARROW_HALF_ANGLE, angle + 180 + _ARROW_HALF_ANGLE]
    return [[tip, (tip[0] + wing * math.cos(math.radians(a)),
                   tip[1] + wing * math.sin(math.radians(a)))] for a in back]


def _nozzle_head(base: tuple[float, float], tip: tuple[float, float]
                 ) -> tuple[tuple[float, float], list[tuple[float, float]]]:
    """헤드 삼각형의 밑변 가운데와 두 밑각. 꼭짓점은 `@` 노드 자리 그대로다."""
    dx, dy = tip[0] - base[0], tip[1] - base[1]
    reach = math.hypot(dx, dy)
    # 방향이 없는 노즐은 tip 이 입력노드에 얹혀 있다. 아래로 매달아 자리는 보이되
    # 방향을 지어내지 않는다 — 유도 여부는 리포트가 따로 싣는다.
    ux, uy = (dx / reach, dy / reach) if reach > 0 else (0.0, -1.0)
    # 스텁보다 긴 머리는 분기점을 덮는다. 짧으면 그만큼 줄여 밑변이 관로를 넘지 않게.
    length = min(_HEAD_LENGTH_UNITS, reach) if reach > 0 else _HEAD_LENGTH_UNITS
    half = _HEAD_HALF_WIDTH_UNITS * length / _HEAD_LENGTH_UNITS
    back = (tip[0] - ux * length, tip[1] - uy * length)
    return back, [(back[0] - uy * half, back[1] + ux * half),
                  (back[0] + uy * half, back[1] - ux * half)]


# ── 렌더 ────────────────────────────────────────────────────────────────────


@dataclass
class RenderReport:
    output: Path
    link_item: str
    node_item: str
    orientation: str
    # 실제로 쓴 스위치. None 으로 부르면 원본 SDF 설정이 들어오므로 되읽을 길이 필요하다.
    link_labels: bool = True
    node_labels: bool = True
    flow_arrows: bool = True
    pipes_drawn: int = 0
    pipes_unplaced: tuple[str, ...] = ()
    devices_drawn: int = 0
    nozzles_drawn: int = 0
    # 원본이 자리를 주지 않아 유도해 그린 헤드 / 유도할 근거조차 없던 헤드. 전량 나열한다.
    nozzles_derived: tuple[str, ...] = ()
    nozzles_undirected: tuple[str, ...] = ()
    blank_link_values: tuple[str, ...] = ()
    blank_node_values: tuple[str, ...] = ()
    # 결과에는 있는데 도면에 형상이 없어 그리지 못한 라벨. 전량 나열한다.
    undrawn_result_pipes: tuple[str, ...] = ()
    undrawn_result_nodes: tuple[str, ...] = ()
    link_bands: tuple[float, ...] = ()
    node_bands: tuple[float, ...] = ()
    # 라벨 배치 (D4). 자리를 못 찾아 그리지 않은 것은 조용히 사라지지 않는다.
    labels_drawn: int = 0
    labels_with_leader: int = 0
    labels_dropped: tuple[str, ...] = ()
    label_seconds: float = 0.0
    warnings: list[str] = field(default_factory=list)


_FONT_CACHE: dict[str, FontProperties | None] = {}


def _korean_font() -> FontProperties | None:
    """한글이 그려지는 글꼴. 없으면 None — 조용히 네모로 찍지 않고 경고를 올린다."""
    if "font" in _FONT_CACHE:
        return _FONT_CACHE["font"]
    found: FontProperties | None = None
    for name in _KOREAN_FONTS:
        path = Path("C:/Windows/Fonts") / name
        if path.exists():
            found = FontProperties(fname=str(path))
            break
    if found is None:
        for entry in font_manager.fontManager.ttflist:
            if Path(entry.fname).name.lower() in _KOREAN_FONTS:
                found = FontProperties(fname=entry.fname)
                break
    _FONT_CACHE["font"] = found
    return found


def _markers(ax, groups: dict[str, list[tuple[float, float]]], *,
             size: float, edge: float, face: str, zorder: int) -> None:
    """같은 모양끼리 한 번에 찍는다 — 점 하나마다 부르면 PDF 안 표현이 개수에 따라
    달라져, 표시 항목만 바꿔도 도형 지문이 흔들린다."""
    for mark, points in groups.items():
        if not points:
            continue
        ax.plot([p[0] for p in points], [p[1] for p in points], linestyle="none",
                marker=mark, markersize=size, markerfacecolor=face,
                markeredgecolor="#222222", markeredgewidth=edge, zorder=zorder)


def _note_size(item, per_unit: float) -> float:
    """도면 주기 글자 크기(pt). 모델 단위 높이를 종이 pt 로 환산한다."""
    units = _NOTE_UNITS * (item.typesize or _NOTE_TYPESIZE) / _NOTE_TYPESIZE
    return units * per_unit


def _text_measurer(fig: Figure, ax, font: FontProperties | None,
                   fontsize: float) -> Callable[[str], tuple[float, float]]:
    """글자 한 줄이 도면 좌표로 몇 칸을 차지하는지 재는 자를 만든다.

    폭은 글자 수로 어림하지 않는다 — 한글은 라틴 문자의 두 배라 어림하면 밀집
    구간에서 겹침 판정이 그대로 틀린다. 실제 글꼴로 재고 문자열마다 캐시한다.

    높이는 재지 않고 글꼴 한 줄 높이로 잡는다. Agg 는 4.6pt 글자를 정수 픽셀로
    반올림해 재기 때문에 실측값이 잉크보다 16% 낮게 나오고, ``va="center"`` 가
    맞추는 것도 잉크가 아니라 이 한 줄 상자다.
    """
    renderer = FigureCanvasAgg(fig).get_renderer()
    props = font.copy() if font is not None else FontProperties()
    props.set_size(fontsize)
    # 등축척은 그릴 때 축 상자를 줄이면서 걸린다. 먼저 적용하지 않으면 여기서 읽는
    # 자가 가로와 세로에 서로 다른 눈금을 갖고, 세로로 세운 라벨만 상자가 틀어진다.
    ax.apply_aspect()
    inverse = ax.transData.inverted()
    origin = inverse.transform((0.0, 0.0))
    line_px = fontsize * fig.dpi / 72.0
    cache: dict[str, tuple[float, float]] = {}

    def measure(text: str) -> tuple[float, float]:
        got = cache.get(text)
        if got is None:
            width, _, _ = renderer.get_text_width_height_descent(text, props, False)
            far = inverse.transform((width, line_px))
            got = cache[text] = (abs(far[0] - origin[0]), abs(far[1] - origin[1]))
        return got

    return measure


def _pt_per_unit(fig: Figure, ax) -> float:
    """모델 좌표 한 칸이 종이에서 몇 pt 인지. 모델 단위에 고정된 기호를 pt 로만 받는
    자리(획 굵기)에 넘기려면 이 자가 필요하다. 등축척은 축 상자를 줄이면서 걸리므로
    먼저 적용해야 가로·세로가 같은 눈금을 갖는다."""
    ax.apply_aspect()
    span = ax.transData.transform((1.0, 0.0))[0] - ax.transData.transform((0.0, 0.0))[0]
    return span * 72.0 / fig.dpi


def _rows(bound: BoundModel | None, model: DisplayModel) -> tuple[
        dict[str, LinkRow], dict[str, NozzleRow], dict[str, NodeRow]]:
    pin = bound.table("Pipes-input") if bound else None
    pres = bound.table("Pipes-results") if bound else None
    input_rows = pin.indexed() if pin else {}
    result_rows = pres.indexed() if pres else {}

    links = {
        p.label: LinkRow(
            pipe=p,
            result=bound.pipes.get(normalize_label(p.label)) if bound else None,
            input_row=input_rows.get(normalize_label(p.label), {}),
            result_row=result_rows.get(normalize_label(p.label), {}),
        )
        for p in model.pipes
    }
    nozzles = {
        z.label: NozzleRow(z, bound.nozzles.get(normalize_label(z.label)) if bound else None)
        for z in model.nozzles
    }
    nodes = {
        n.label: NodeRow(n, bound.nodes.get(normalize_label(n.label)) if bound else None)
        for n in model.real_nodes
    }
    return links, nozzles, nodes


def _block_grid(page_pt: tuple[float, float]) -> tuple[float, float, list[float]]:
    """표제란 격자를 종이 pt 로. 돌려주는 y 는 종이 아래에서 잰 행 경계이고
    ys[0] 이 맨 아래(테두리와 겹친다), ys[-1] 이 표제란 위다."""
    right = page_pt[0] - _FRAME_INSET_PT
    ys = [_FRAME_INSET_PT]
    for height in reversed(_BLOCK_ROW_PT):
        ys.append(ys[-1] + height)
    return right - _BLOCK_WIDTH_PT, right, ys


def _draw_furniture(fig: Figure, page_pt: tuple[float, float],
                    cells: Sequence[tuple[int, float, str]],
                    font: FontProperties | None) -> None:
    """테두리와 표제란 괘선을 긋고 칸에 글씨를 넣는다.

    cells 는 (행 번호(위에서 0), 칸 왼쪽(표제란 폭에 대한 비율), 글) 이다.
    """
    w_pt, h_pt = page_pt
    left, right, ys = _block_grid(page_pt)
    span = right - left

    def rule(x0: float, y0: float, x1: float, y1: float) -> None:
        fig.add_artist(Line2D((x0 / w_pt, x1 / w_pt), (y0 / h_pt, y1 / h_pt),
                              transform=fig.transFigure, color="#000000", linewidth=_RULE_PT))

    far_x, far_y = w_pt - _FRAME_INSET_PT, h_pt - _FRAME_INSET_PT
    rule(_FRAME_INSET_PT, _FRAME_INSET_PT, far_x, _FRAME_INSET_PT)
    rule(_FRAME_INSET_PT, far_y, far_x, far_y)
    rule(_FRAME_INSET_PT, _FRAME_INSET_PT, _FRAME_INSET_PT, far_y)
    rule(far_x, _FRAME_INSET_PT, far_x, far_y)

    for y in ys[1:]:
        rule(left, y, right, y)
    rule(left, ys[0], left, ys[-1])
    # 2 행은 절반, 3 행은 1/4·3/4 에서 갈린다.
    rule(left + span / 2, ys[3], left + span / 2, ys[4])
    rule(left + span / 4, ys[2], left + span / 4, ys[3])
    rule(left + span * 3 / 4, ys[2], left + span * 3 / 4, ys[3])

    top = len(_BLOCK_ROW_PT)
    for row, frac, text in cells:
        if not text:
            continue
        fig.text((left + span * frac + _BLOCK_INSET_PT) / w_pt,
                 (ys[top - row - 1] + ys[top - row]) / 2 / h_pt,
                 text, fontsize=_BLOCK_FONT_PT, va="center", ha="left", fontproperties=font)


def _draw_legend(fig: Figure, page_pt: tuple[float, float], lane: int, title: str,
                 edges: Sequence[float], fmt: ValueFormat, palette: Sequence[str],
                 font: FontProperties | None) -> None:
    """범례 한 벌 — 왼쪽에 항목 이름, 오른쪽에 3 칸 2 줄 격자. 참조가 비워 둔
    표제란 아래 두 행(10mm 씩)이 한 벌씩 딱 들어가는 자리다."""
    w_pt, h_pt = page_pt
    left = (_block_grid(page_pt)[0] + _BLOCK_INSET_PT) / w_pt
    base = (_LEGEND_BOTTOM_PT + lane * _LEGEND_LANE_PT) / h_pt
    fig.text(left, base, title, fontsize=_LEGEND_FONT_PT, va="bottom", ha="left",
             fontproperties=font)
    count = len(edges) + 1 if edges else 1
    for i in range(count):
        row, col = divmod(i, 3)
        x = left + (_LEGEND_SWATCH_DX_PT + col * _LEGEND_COL_PT) / w_pt
        y = base + (1 - row) * _LEGEND_ROW_PT / h_pt
        fig.add_artist(Rectangle((x, y), _LEGEND_SWATCH_PT / w_pt, _LEGEND_SWATCH_PT / h_pt,
                                 transform=fig.transFigure,
                                 facecolor=palette[i], edgecolor="none"))
        if not edges:
            label = "전량 동일"
        elif i < len(edges):
            label = f"< {fmt.text(edges[i])}"
        else:
            label = f"≥ {fmt.text(edges[-1])}"
        fig.text(x + _LEGEND_LABEL_DX_PT / w_pt, y, label, fontsize=_LEGEND_FONT_PT,
                 va="bottom", ha="left", fontproperties=font)


def render_iso(
    source: BoundModel | DisplayModel,
    output: str | Path,
    *,
    preset: str | None = None,
    link_item: str | None = None,
    node_item: str | None = None,
    show_link_labels: bool | None = None,
    show_node_labels: bool | None = None,
    show_arrows: bool | None = None,
    section: str = "",
    also_png: str | Path | None = None,
) -> RenderReport:
    """ISO 도면 한 장을 벡터 PDF 로 쓴다.

    ``source`` 는 D2 의 BoundModel 이거나, 결과 XML 없이 형상만 그릴 때의
    DisplayModel 이다. 후자면 SDF 만으로 되는 항목(관경·길이·라벨)만 값이 붙는다.

    ``also_png`` 를 주면 같은 도형에서 미리보기 PNG 를 한 장 더 뽑는다. 두 번
    호출하면 도형을 두 번 짓는다 — 짓는 값은 같으니 한 번만 짓는다.

    이름표·화살표 스위치를 ``None`` 으로 두면 원본 SDF 의 ``<Display-options>``,
    즉 PIPENET 이 자기 화면에 쓰던 설정을 그대로 따른다. 원본에 그 항목이 없을
    때만 켠다.
    """
    bound = source if isinstance(source, BoundModel) else None
    model = bound.model if bound else source
    out = Path(output)

    seen = model.source_display
    def follow(given: bool | None, stored: bool | None) -> bool:
        return given if given is not None else (True if stored is None else stored)
    link_labels = follow(show_link_labels, seen.link_labels)
    node_labels = follow(show_node_labels, seen.node_labels)
    arrows = follow(show_arrows, seen.flow_arrows)

    chosen = PRESETS[preset] if preset else None
    link_name = link_item or (chosen.link if chosen else "None")
    node_name = node_item or (chosen.node if chosen else "None")
    if link_name not in _LINK_BY_NAME:
        raise ValueError(f"모르는 관로 표시 항목 '{link_name}' — {sorted(_LINK_BY_NAME)}")
    if node_name not in _NODE_BY_NAME:
        raise ValueError(f"모르는 노드 표시 항목 '{node_name}' — {sorted(_NODE_BY_NAME)}")
    link = _LINK_BY_NAME[link_name]
    node = _NODE_BY_NAME[node_name]

    warnings: list[str] = []
    font = _korean_font()
    if font is None:
        warnings.append("한글 글꼴을 찾지 못했다 — 한글이 네모로 찍힌다")

    link_fmt = _format_for(link, model.units, warnings)
    node_fmt = _format_for(node, model.units, warnings)

    links, nozzles, nodes = _rows(bound, model)
    coords = {n.label: (n.x, n.y) for n in model.nodes}

    # ── 값 — 여기서 화면 단위로 옮기고, 밴드도 라벨도 그 값으로 정한다 ──
    subjects = links if link.scope == "pipe" else nozzles
    link_values = {label: link_fmt.display(link.read(row)) for label, row in subjects.items()}
    node_values = {label: node_fmt.display(node.read(row)) for label, row in nodes.items()}

    def _numbers(values: dict[str, Any]) -> list[float]:
        return [v for v in values.values() if isinstance(v, float) and math.isfinite(v)]

    link_edges = band_edges(_numbers(link_values)) if link.name != "None" else ()
    node_edges = band_edges(_numbers(node_values)) if node.name != "None" else ()

    # ── 종이 ──
    bounds = model.bounds()
    if bounds is None:
        raise ValueError("그릴 좌표가 없다")
    minx, miny, maxx, maxy = bounds
    span_x, span_y = max(maxx - minx, 1e-6), max(maxy - miny, 1e-6)
    landscape = span_x / span_y > A4_MM[1] / A4_MM[0]
    page = (A4_MM[1], A4_MM[0]) if landscape else A4_MM
    fig = Figure(figsize=(page[0] / _MM_PER_INCH, page[1] / _MM_PER_INCH))

    page_pt = (page[0] / _MM_PER_INCH * 72.0, page[1] / _MM_PER_INCH * 72.0)
    margin_x = 12.0 / page[0]
    # 그림틀은 표시 항목과 무관하게 고정이다. 범례가 한 벌이냐 두 벌이냐에 따라
    # 커졌다 작아지면 같은 망이 다른 축척으로 나와 두 장을 겹쳐 볼 수 없다 (지시서 6).
    ax_bottom = _PLATE_BOTTOM_PT / page_pt[1]
    ax_top = 1.0 - _PLATE_TOP_PT / page_pt[1]
    ax = fig.add_axes((margin_x, ax_bottom, 1.0 - 2 * margin_x, ax_top - ax_bottom))
    ax.set_aspect("equal")
    ax.set_axis_off()
    pad = 0.04 * max(span_x, span_y)
    ax.set_xlim(minx - pad, maxx + pad)
    ax.set_ylim(miny - pad, maxy + pad)

    # ── 관로 ──
    paths = {label: _polyline(row.pipe, coords) for label, row in links.items()}
    unplaced = [label for label, path in paths.items() if path is None]
    placed = {label: path for label, path in paths.items() if path is not None}

    per_unit = _pt_per_unit(fig, ax)
    pipe_width = _PIPE_WIDTH_UNITS * per_unit
    segments, colours = [], []
    blank_links: list[str] = []
    banded = link.scope == "pipe" and link.name != "None"
    for label, path in placed.items():
        segments.append(path)
        band = _band_of(link_values.get(label), link_edges) if banded else None
        colours.append(BAND_COLOURS[band] if band is not None else
                       (_BLANK_COLOUR if banded else "#333333"))
    if segments:
        ax.add_collection(LineCollection(segments, colors=colours, linewidths=pipe_width,
                                         capstyle="round", joinstyle="round", zorder=2))

    # ── 장치 링크 (펌프·감압밸브·고정손실) ──
    device_segments = []
    device_points: dict[str, list[tuple[float, float]]] = {}
    for dev in model.devices:
        a, b = coords.get(dev.input_node), coords.get(dev.output_node)
        if a is None or b is None:
            continue
        device_segments.append([a, b])
        device_points.setdefault(DEVICE_MARKS.get(dev.kind, "s"), []).append(
            ((a[0] + b[0]) / 2, (a[1] + b[1]) / 2))
    if device_segments:
        ax.add_collection(LineCollection(device_segments, colors="#222222",
                                         linewidths=_DEVICE_LINK_UNITS * per_unit, zorder=3))
    device_pt = _DEVICE_UNITS * per_unit
    equipment_pt = _EQUIPMENT_UNITS * per_unit
    edge_pt = _MARKER_EDGE_UNITS * per_unit
    _markers(ax, device_points, size=device_pt, edge=edge_pt, face="white", zorder=5)

    # ── 특수기기 (A/V, FLEX) ──
    equipment_points: dict[str, list[tuple[float, float]]] = {}
    for label, path in placed.items():
        for eq in links[label].pipe.equipment:
            # 신축배관은 가위표, 밸브류는 네모. 삼각은 쓰지 않는다 — 노즐 헤드가 같은
            # 자리에 검은 삼각으로 앉아 기기 삼각과 구별되지 않는다.
            mark = "X" if eq.description.upper().startswith(("FX", "FLEX")) else "s"
            equipment_points.setdefault(mark, []).append(
                _point_at(path, eq.rel_position if eq.rel_position is not None else 0.5))
    _markers(ax, {"X": equipment_points.get("X", [])}, size=equipment_pt, edge=edge_pt,
             face="#ffffff", zorder=5)
    _markers(ax, {"s": equipment_points.get("s", [])}, size=equipment_pt, edge=edge_pt,
             face="#222222", zorder=5)

    # ── 노즐 ──
    measured_stub = _measured_stub(nozzles, coords)
    tips, derived_nozzles, undirected_nozzles = _nozzle_tips(
        nozzles, model.pipes, coords, measured_stub or _STUB_UNITS)
    is_derived = set(derived_nozzles)
    drawn_nozzles = 0
    nozzle_tips: list[tuple[float, float]] = []
    # 노즐도 입력노드와 출력노드를 잇는 링크다. 그 선을 그리지 않으면 헤드가
    # 배관에서 떨어져 떠 있는 것처럼 보인다. 실측 자리는 관로와 같은 실선으로,
    # 유도한 자리는 주황 점선으로 — 지어낸 것을 실측인 척 그리지 않는다.
    stubs: list[list[tuple[float, float]]] = []
    guessed_stubs: list[list[tuple[float, float]]] = []
    for label, row in nozzles.items():
        base = coords.get(row.nozzle.input_node)
        tip = tips.get(label)
        if base is None or tip is None:
            continue
        drawn_nozzles += 1
        nozzle_tips.append(tip)
        estimated = label in is_derived
        back, corners = _nozzle_head(base, tip)
        (guessed_stubs if estimated else stubs).append([base, back])
        ax.add_patch(Polygon(
            [tip, *corners], closed=True, facecolor="none",
            edgecolor=_DERIVED_COLOUR if estimated else "#222222",
            linewidth=pipe_width, zorder=6))
    if stubs:
        ax.add_collection(LineCollection(stubs, colors="#333333", linewidths=pipe_width,
                                         capstyle="round", zorder=2))
    if guessed_stubs:
        ax.add_collection(LineCollection(guessed_stubs, colors=_DERIVED_COLOUR,
                                         linewidths=pipe_width,
                                         linestyles=[(0, (2.0, 1.5))], zorder=5))

    # ── 노드 ──
    # 밴드마다 따로 그리지 않고 한 컬렉션에 색만 나눠 담는다 — 그래야 표시 항목을
    # 바꿔도 점이 놓이는 자리가 한 톨도 달라지지 않는다 (지시서 6).
    blank_nodes: list[str] = []
    node_points, node_colours = [], []
    for label in nodes:
        point = coords.get(label)
        if point is None:
            continue
        value = node_values.get(label)
        band = _band_of(value, node_edges) if node.name != "None" else None
        if node.name != "None" and value is None:
            blank_nodes.append(label)
        node_points.append(point)
        node_colours.append(NODE_BAND_COLOURS[band] if band is not None else "#333333")
    if node_points:
        ax.add_collection(EllipseCollection(
            [_NODE_DOT_UNITS] * len(node_points), [_NODE_DOT_UNITS] * len(node_points),
            [0.0] * len(node_points), units="xy", offsets=node_points,
            offset_transform=ax.transData, facecolors=node_colours, linewidths=0,
            zorder=4))

    # ── 흐름 화살표 ──
    # 라벨보다 먼저 자리를 잡는다. D4 에 피해야 할 자리로 넘겨야 값 글자가 화살표
    # 위에 앉아 방향을 가리지 않는다.
    arrow_strokes: list[list[tuple[float, float]]] = []
    arrow_colours: list[str] = []
    arrow_boxes: list[tuple[float, float, float, float]] = []
    if arrows:
        for (label, path), colour in zip(placed.items(), colours):
            result = links[label].result
            flow = result.flow_m3s if result else None
            if not flow:
                continue
            spot = _along(path, _ARROW_AT)
            if spot is None:
                continue
            tip, angle = spot
            wings = _chevron(tip, angle + (180 if flow < 0 else 0), _ARROW_WING_UNITS)
            arrow_strokes.extend(wings)
            arrow_colours.extend([colour] * len(wings))
            xs = [x for wing in wings for x, _ in wing]
            ys = [y for wing in wings for _, y in wing]
            arrow_boxes.append((min(xs), min(ys), max(xs), max(ys)))
    if arrow_strokes:
        ax.add_collection(LineCollection(arrow_strokes, colors=arrow_colours,
                                         linewidths=pipe_width, capstyle="butt",
                                         zorder=3))

    # ── 글자 (D4 가 자리를 정한다) ──
    requests: list[LabelRequest] = []
    for label, path in placed.items():
        (mx, my), angle = _longest_segment(path)
        if angle > 90 or angle < -90:
            angle += 180
        parts = []
        if link_labels:
            parts.append(label)
        if banded:
            text = link_fmt.text(link_values.get(label))
            if text:
                parts.append(text)
            else:
                blank_links.append(label)
        if parts:
            requests.append(LabelRequest(label, "  ".join(parts), (mx, my), angle, "link"))

    if link.scope == "nozzle" and link.name != "None":
        for label, row in nozzles.items():
            tip = tips.get(label)
            text = link_fmt.text(link_values.get(label))
            if tip is None:
                continue
            if not text:
                blank_links.append(label)
                continue
            requests.append(LabelRequest(label, text, tip, 0.0, "nozzle"))

    if node_labels or node.name != "None":
        for label, row in nodes.items():
            point = coords.get(label)
            if point is None:
                continue
            parts = [label] if node_labels else []
            if node.name != "None":
                parts.append(node_fmt.text(node_values.get(label)))
            text = " ".join(p for p in parts if p)
            if text:
                requests.append(LabelRequest(label, text, point, 0.0, "node"))

    label_pt = _LABEL_UNITS * per_unit
    measure = _text_measurer(fig, ax, font, label_pt)
    # 도면 주기는 SDF 가 정해 둔 자리다. 옮기지 않고 피해야 할 자리로만 넘긴다.
    fixed = []
    for item in model.texts:
        scale = _note_size(item, per_unit) / label_pt
        w, h = measure(item.text)
        fixed.append((item.x, item.y, item.x + w * scale, item.y + h * scale))
    # 기호도 피한다. 값이 밸브나 노즐 위에 얹히면 겹친 라벨이 없어도 읽을 수 없다.
    for points, size in ((device_points, _DEVICE_UNITS), (equipment_points, _EQUIPMENT_UNITS)):
        half = size / 2
        for group in points.values():
            fixed.extend((x - half, y - half, x + half, y + half) for x, y in group)
    head = _HEAD_LENGTH_UNITS
    fixed.extend((x - head, y - head, x + head, y + head) for x, y in nozzle_tips)
    fixed.extend(arrow_boxes)
    labels, layout = lay_out(requests, measure, obstacles=fixed)

    leaders = [lab.leader for lab in labels if lab.leader]
    if leaders:
        ax.add_collection(LineCollection(leaders, colors=_LEADER_COLOUR,
                                         linewidths=_LEADER_WIDTH_UNITS * per_unit, zorder=6))
    for lab in labels:
        # anchor 모드라야 (x, y) 가 회전 전 상자의 한가운데로 고정된다. 기본 모드는
        # 회전한 뒤의 상자를 다시 맞춰서 D4 가 잡아 둔 자리와 어긋난다.
        ax.text(lab.x, lab.y, lab.text, rotation=lab.angle, rotation_mode="anchor",
                ha="center", va="center", color="#111111", fontproperties=font,
                fontsize=label_pt, zorder=7)

    # ── 도면 주기 ──
    for text in model.texts:
        ax.text(text.x, text.y, text.text, fontsize=_note_size(text, per_unit),
                color=text.colour or "#000000", fontproperties=font,
                ha="left", va="bottom", zorder=8)

    # ── 테두리·표제란 (지시서 7-2: 참조의 양식만 따르고 칸 내용은 우리 것) ──
    titles = list(bound.title) if bound else []
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    origin = f"원본 {model.source.name}" + (f" + {bound.document.source.name}" if bound else "")
    name = " · ".join(titles) if titles else model.source.stem
    # 표시 항목은 범례가 이름째 적어 주므로 여기서 되풀이하지 않는다. 원본 파일명은
    # 이 칸에 들어가기엔 길어 각주로 뺐다.
    _draw_furniture(fig, page_pt, (
        (0, 0.0, f"{name} — {section}" if section else name),
        (1, 0.5, "FNCADnet 모듈 D 생성"),
        (2, 0.0, "ISO 계통도"),
        (2, 0.25, stamp),
        (2, 0.75, "Page 1 of 1"),
    ), font)

    lane = 0
    if link.name != "None":
        _draw_legend(fig, page_pt, lane, f"{link.name}\n({link_fmt.symbol})",
                     link_edges, link_fmt, BAND_COLOURS, font)
        lane += 1
    if node.name != "None":
        _draw_legend(fig, page_pt, lane, f"{node.name}\n({node_fmt.symbol})",
                     node_edges, node_fmt, NODE_BAND_COLOURS, font)

    # 각주는 테두리 안, 표제란 왼쪽에 남는 자리에만 든다 — 한 줄로 흘리면 범례를 덮는다.
    fig.text((_FRAME_INSET_PT + _BLOCK_INSET_PT) / page_pt[0],
             (_FRAME_INSET_PT + _BLOCK_INSET_PT) / page_pt[1],
             f"{origin}\n"
             "PIPENET 이 계산한 결과를 옮겨 그린 도면이다.\n"
             "값은 결과 XML 원문이며 이 프로그램이\n"
             "계산하거나 보간하지 않는다.",
             fontsize=_FOOTNOTE_PT, va="bottom", ha="left", color="#555555",
             fontproperties=font)

    out.parent.mkdir(parents=True, exist_ok=True)
    if out.suffix.lower() == ".png":
        # 시각 회귀용. 해상도를 고정하지 않으면 픽셀 비교가 성립하지 않는다 (지시서 8).
        fig.set_dpi(PNG_DPI)
        FigureCanvasAgg(fig).print_png(str(out))
    else:
        # fonttype 42 = TrueType 원본 임베드. 기본값 3(Type-3)은 글자를 곡선으로 흩어
        # 놓아 PDF 에서 텍스트를 다시 뽑을 수 없다 — 지시서 8 의 골든 대조가 막힌다.
        with matplotlib.rc_context({"pdf.fonttype": 42}):
            FigureCanvasPdf(fig).print_pdf(str(out), metadata={
                "Title": " / ".join(titles) or model.source.stem,
                "Creator": "FNCADnet module D",
                "Subject": f"link={link.name}; node={node.name}",
            })
    if also_png is not None:
        extra = Path(also_png)
        extra.parent.mkdir(parents=True, exist_ok=True)
        # dpi 를 올린 뒤에는 되돌리지 않는다 — 벡터 출력이 이미 끝난 뒤라야 한다.
        fig.set_dpi(PNG_DPI)
        FigureCanvasAgg(fig).print_png(str(extra))

    if unplaced:
        warnings.append(f"양끝 노드를 못 찾은 관로 {len(unplaced)}개 — 그리지 않았다")
    if blank_links or blank_nodes:
        warnings.append(
            f"값이 없어 빈칸으로 둔 라벨 관로 {len(blank_links)}개 / 노드 {len(blank_nodes)}개"
        )
    if layout.dropped:
        warnings.append(
            f"겹치지 않는 자리를 못 찾은 라벨 {len(layout.dropped)}개 — 겹친 채 두지 않고 "
            f"그리지 않았다: {', '.join(layout.dropped[:12])}"
        )
    if derived_nozzles:
        warnings.append(
            f"원본이 헤드 자리를 주지 않아 유도해 그린 것 {len(derived_nozzles)}개 — 입사 "
            f"관로의 연장선에 {measured_stub or _STUB_UNITS:.1f} 만큼 내고 점선·주황 윤곽으로 "
            f"갈라 표시했다"
            + (" (같은 도면의 다른 헤드에서 잰 길이)" if measured_stub
               else " (이 도면에는 잰 길이가 없어 참조 코퍼스 최빈값)")
            + f": {', '.join(derived_nozzles[:12])}"
        )
    if undirected_nozzles:
        warnings.append(
            f"입사 관로가 없어 방향을 정할 근거가 없는 헤드 {len(undirected_nozzles)}개 — "
            f"자리를 지어내지 않고 입력노드에 겹쳐 두었다: {', '.join(undirected_nozzles[:12])}"
        )
    undrawn_pipes = bound.report.xml_only_pipes if bound else ()
    undrawn_nodes = bound.report.xml_only_nodes if bound else ()
    if undrawn_pipes or undrawn_nodes:
        warnings.append(
            f"결과에만 있고 도면에 형상이 없는 라벨 관로 {len(undrawn_pipes)}개 / "
            f"노드 {len(undrawn_nodes)}개 — 그리지 않았다"
        )
    return RenderReport(
        output=out,
        link_item=link.name,
        node_item=node.name,
        orientation="landscape" if landscape else "portrait",
        link_labels=link_labels,
        node_labels=node_labels,
        flow_arrows=arrows,
        pipes_drawn=len(segments),
        pipes_unplaced=tuple(unplaced),
        devices_drawn=len(device_segments),
        nozzles_drawn=drawn_nozzles,
        nozzles_derived=tuple(derived_nozzles),
        nozzles_undirected=tuple(undirected_nozzles),
        blank_link_values=tuple(blank_links),
        blank_node_values=tuple(blank_nodes),
        undrawn_result_pipes=undrawn_pipes,
        undrawn_result_nodes=undrawn_nodes,
        link_bands=tuple(link_edges),
        node_bands=tuple(node_edges),
        labels_drawn=layout.placed,
        labels_with_leader=layout.leaders,
        labels_dropped=layout.dropped,
        label_seconds=layout.seconds,
        warnings=warnings,
    )
