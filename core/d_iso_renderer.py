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
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Sequence

import matplotlib
from matplotlib import font_manager
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.backends.backend_pdf import FigureCanvasPdf
from matplotlib.collections import LineCollection
from matplotlib.figure import Figure
from matplotlib.font_manager import FontProperties
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
# 낮은 값 → 높은 값. 흑백 인쇄에서도 순서가 보이도록 명도를 단조로 깔았다.
BAND_COLOURS = ("#1f4e9c", "#3d8bcd", "#59b4a8", "#c9a227", "#d9663d", "#a01f1f")
_BLANK_COLOUR = "#9a9a9a"

# 라벨 글꼴 크기(pt)의 기준. SDF 의 typesize 단위는 파일이 밝히지 않으므로 절대 크기로
# 환산하지 않고, Label-display 의 print-font 를 1 로 삼아 상대비만 지킨다.
_LABEL_PT = 4.6
_REFERENCE_PRINT_FONT = 48.0

# 지시선. 망보다 연하고 가늘어야 관로로 오독되지 않는다. 값이 없어 회색으로 그린
# 관로(_BLANK_COLOUR)와는 색이 달라야 한다 — 검사가 둘을 색으로 갈라 본다.
_LEADER_COLOUR = "#7f7f7f"

# 기호 크기(pt). 라벨이 피해야 할 자리를 잡는 데도 쓰이므로 그리는 쪽과 한 값이어야 한다.
_DEVICE_PT = 3.4
_EQUIPMENT_PT = 2.6

# 노즐 스텁. SDF 가 @/n 좌표를 입력노드와 같은 자리에 두면 헤드가 분기점 위에 겹쳐
# 찍힌다. 원본은 고칠 수 없으므로(지시서 7-3) 그릴 때 방향을 유도한다. 근거는 참조
# 코퍼스 SDF 334 개 / 노즐 3437 개 실측이다 — 입사 관로가 있는 3115 개 중 3079 개
# (98.8%)가 관로의 연장선이고 수직은 0 건, 길이는 도면 span 의 중앙값 1.208% 였다.
_STUB_SPAN_RATIO = 0.01208
# 유도한 것은 실측과 같은 잉크로 그리지 않는다 — 점선 + 이 색으로 갈라 보인다.
_DERIVED_COLOUR = "#c2410c"

_KOREAN_FONTS = ("malgun.ttf", "malgunsl.ttf", "NanumGothic.ttf", "gulim.ttc", "batang.ttc")

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


def _nozzle_tips(nozzles: dict[str, Any], pipes: Sequence[DisplayPipe],
                 coords: dict[str, tuple[float, float]], span: float
                 ) -> tuple[dict[str, tuple[float, float]], list[str], list[str]]:
    """헤드 삼각형의 꼭짓점 자리. 원본이 방향을 준 것은 그대로 쓰고, 출력노드가
    입력노드와 겹쳐 방향이 없는 것만 유도한다. 어느 쪽인지는 갈라서 돌려준다 —
    유도한 것을 실측인 척 그리지 않기 위해서다."""
    stub = span * _STUB_SPAN_RATIO
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


# ── 렌더 ────────────────────────────────────────────────────────────────────


@dataclass
class RenderReport:
    output: Path
    link_item: str
    node_item: str
    orientation: str
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
             size: float, face: str, zorder: int) -> None:
    """같은 모양끼리 한 번에 찍는다 — 점 하나마다 부르면 PDF 안 표현이 개수에 따라
    달라져, 표시 항목만 바꿔도 도형 지문이 흔들린다."""
    for mark, points in groups.items():
        if not points:
            continue
        ax.plot([p[0] for p in points], [p[1] for p in points], linestyle="none",
                marker=mark, markersize=size, markerfacecolor=face,
                markeredgecolor="#222222", markeredgewidth=0.7, zorder=zorder)


def _note_size(item) -> float:
    """도면 주기 글자 크기(pt). SDF typesize 를 라벨 크기 기준으로 환산한다."""
    size = _LABEL_PT * (item.typesize or _REFERENCE_PRINT_FONT) / _REFERENCE_PRINT_FONT
    return max(3.0, min(size, 12.0))


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


def _line_widths(bores: Sequence[float | None]) -> Callable[[float | None], float]:
    """선 굵기를 호칭경에 비례시킨다. 관경을 모르는 관로는 가장 가늘게."""
    known = [b for b in bores if b]
    top = max(known) if known else 0.0
    def width(bore: float | None) -> float:
        if not bore or top <= 0:
            return 0.4
        return 0.4 + 1.8 * (bore / top)
    return width


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


def _draw_legend(fig: Figure, rect: tuple[float, float, float, float], title: str,
                 edges: Sequence[float], fmt: ValueFormat, font: FontProperties | None) -> None:
    x, y, w, h = rect
    fig.text(x, y + h, title, fontsize=5.2, va="top", ha="left",
             fontproperties=font, weight="bold")
    count = len(edges) + 1 if edges else 1
    swatch_w, swatch_h = w / (count + 1.2), h * 0.34
    top = y + h * 0.42
    for i in range(count):
        left = x + i * swatch_w * 1.15
        fig.add_artist(Rectangle((left, top), swatch_w * 0.9, swatch_h,
                                 transform=fig.transFigure,
                                 facecolor=BAND_COLOURS[i], edgecolor="none"))
        if not edges:
            label = "전량 동일"
        elif i < len(edges):
            label = f"< {fmt.text(edges[i])}"
        else:
            label = f"≥ {fmt.text(edges[-1])}"
        fig.text(left, top - h * 0.12, label, fontsize=4.4, va="top", ha="left",
                 fontproperties=font)


def render_iso(
    source: BoundModel | DisplayModel,
    output: str | Path,
    *,
    preset: str | None = None,
    link_item: str | None = None,
    node_item: str | None = None,
    show_labels: bool = True,
    show_arrows: bool = True,
    section: str = "",
    also_png: str | Path | None = None,
) -> RenderReport:
    """ISO 도면 한 장을 벡터 PDF 로 쓴다.

    ``source`` 는 D2 의 BoundModel 이거나, 결과 XML 없이 형상만 그릴 때의
    DisplayModel 이다. 후자면 SDF 만으로 되는 항목(관경·길이·라벨)만 값이 붙는다.

    ``also_png`` 를 주면 같은 도형에서 미리보기 PNG 를 한 장 더 뽑는다. 두 번
    호출하면 도형을 두 번 짓는다 — 짓는 값은 같으니 한 번만 짓는다.
    """
    bound = source if isinstance(source, BoundModel) else None
    model = bound.model if bound else source
    out = Path(output)

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

    margin_x, margin_y = 12.0 / page[0], 12.0 / page[1]
    header_h, legend_h, footer_h = 16.0 / page[1], 9.0 / page[1], 6.0 / page[1]
    # 범례 칸은 쓰든 안 쓰든 두 줄을 비워 둔다. 표시 항목에 따라 그림틀이 커졌다
    # 작아지면 같은 망이 다른 축척으로 나와 두 장을 겹쳐 볼 수 없다 (지시서 6).
    legend_total = legend_h * 2
    ax_bottom = margin_y + footer_h
    ax_top = 1.0 - margin_y - header_h - legend_total
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

    width_of = _line_widths([p.bore_m for p in model.pipes])
    segments, colours, widths = [], [], []
    blank_links: list[str] = []
    banded = link.scope == "pipe" and link.name != "None"
    for label, path in placed.items():
        segments.append(path)
        widths.append(width_of(links[label].pipe.bore_m))
        band = _band_of(link_values.get(label), link_edges) if banded else None
        colours.append(BAND_COLOURS[band] if band is not None else
                       (_BLANK_COLOUR if banded else "#333333"))
    if segments:
        ax.add_collection(LineCollection(segments, colors=colours, linewidths=widths,
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
        ax.add_collection(LineCollection(device_segments, colors="#222222", linewidths=1.4,
                                         zorder=3))
    _markers(ax, device_points, size=_DEVICE_PT, face="white", zorder=5)

    # ── 특수기기 (A/V, FLEX) ──
    equipment_points: dict[str, list[tuple[float, float]]] = {}
    for label, path in placed.items():
        for eq in links[label].pipe.equipment:
            # 신축배관은 형상이 휘므로 삼각으로, 밸브류는 네모로 구분한다.
            mark = "v" if eq.description.upper().startswith(("FX", "FLEX")) else "s"
            equipment_points.setdefault(mark, []).append(
                _point_at(path, eq.rel_position if eq.rel_position is not None else 0.5))
    _markers(ax, {"v": equipment_points.get("v", [])}, size=_EQUIPMENT_PT, face="#ffffff",
             zorder=5)
    _markers(ax, {"s": equipment_points.get("s", [])}, size=_EQUIPMENT_PT, face="#222222",
             zorder=5)

    # ── 노즐 ──
    tri = max(span_x, span_y) * 0.006
    tips, derived_nozzles, undirected_nozzles = _nozzle_tips(
        nozzles, model.pipes, coords, max(span_x, span_y))
    is_derived = set(derived_nozzles)
    drawn_nozzles = 0
    nozzle_tips: list[tuple[float, float]] = []
    stubs: list[list[tuple[float, float]]] = []
    for label, row in nozzles.items():
        base = coords.get(row.nozzle.input_node)
        tip = tips.get(label)
        if base is None or tip is None:
            continue
        drawn_nozzles += 1
        nozzle_tips.append(tip)
        angle = math.atan2(tip[1] - base[1], tip[0] - base[0]) if tip != base else -math.pi / 2
        estimated = label in is_derived
        if estimated:
            stubs.append([base, tip])
        ax.add_patch(Polygon(
            [(tip[0], tip[1]),
             (tip[0] - tri * math.cos(angle - 0.4), tip[1] - tri * math.sin(angle - 0.4)),
             (tip[0] - tri * math.cos(angle + 0.4), tip[1] - tri * math.sin(angle + 0.4))],
            closed=True, facecolor="#ffffff" if estimated else "#222222",
            edgecolor=_DERIVED_COLOUR if estimated else "none",
            linewidth=0.5 if estimated else 0.0, zorder=6))
    if stubs:
        ax.add_collection(LineCollection(stubs, colors=_DERIVED_COLOUR, linewidths=0.4,
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
        node_colours.append(BAND_COLOURS[band] if band is not None else "#333333")
    if node_points:
        ax.scatter([p[0] for p in node_points], [p[1] for p in node_points],
                   s=2.5, c=node_colours, marker="o", linewidths=0, zorder=4)

    # ── 글자 (D4 가 자리를 정한다) ──
    requests: list[LabelRequest] = []
    for label, path in placed.items():
        (mx, my), angle = _longest_segment(path)
        if angle > 90 or angle < -90:
            angle += 180
        parts = []
        if show_labels:
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

    if show_labels or node.name != "None":
        for label, row in nodes.items():
            point = coords.get(label)
            if point is None:
                continue
            parts = [label] if show_labels else []
            if node.name != "None":
                parts.append(node_fmt.text(node_values.get(label)))
            text = " ".join(p for p in parts if p)
            if text:
                requests.append(LabelRequest(label, text, point, 0.0, "node"))

    measure = _text_measurer(fig, ax, font, _LABEL_PT)
    # 도면 주기는 SDF 가 정해 둔 자리다. 옮기지 않고 피해야 할 자리로만 넘긴다.
    fixed = []
    for item in model.texts:
        scale = _note_size(item) / _LABEL_PT
        w, h = measure(item.text)
        fixed.append((item.x, item.y, item.x + w * scale, item.y + h * scale))
    # 기호도 피한다. 값이 밸브나 노즐 위에 얹히면 겹친 라벨이 없어도 읽을 수 없다.
    per_pt = measure("0")[1] / _LABEL_PT
    for points, size_pt in ((device_points, _DEVICE_PT), (equipment_points, _EQUIPMENT_PT)):
        half = size_pt * per_pt / 2
        for group in points.values():
            fixed.extend((x - half, y - half, x + half, y + half) for x, y in group)
    fixed.extend((x - tri, y - tri, x + tri, y + tri) for x, y in nozzle_tips)
    labels, layout = lay_out(requests, measure, obstacles=fixed)

    leaders = [lab.leader for lab in labels if lab.leader]
    if leaders:
        ax.add_collection(LineCollection(leaders, colors=_LEADER_COLOUR, linewidths=0.25,
                                         zorder=6))
    for lab in labels:
        # anchor 모드라야 (x, y) 가 회전 전 상자의 한가운데로 고정된다. 기본 모드는
        # 회전한 뒤의 상자를 다시 맞춰서 D4 가 잡아 둔 자리와 어긋난다.
        ax.text(lab.x, lab.y, lab.text, rotation=lab.angle, rotation_mode="anchor",
                ha="center", va="center", color="#111111", fontproperties=font,
                fontsize=_LABEL_PT, zorder=7)

    # ── 흐름 화살표 ──
    if show_arrows:
        for label, path in placed.items():
            result = links[label].result
            flow = result.flow_m3s if result else None
            if not flow:
                continue
            (mx, my), angle = _longest_segment(path)
            if flow < 0:
                angle += 180
            ax.plot([mx], [my], marker=(3, 0, angle - 90), markersize=2.2,
                    color="#111111", zorder=6)

    # ── 도면 주기 ──
    for text in model.texts:
        ax.text(text.x, text.y, text.text, fontsize=_note_size(text),
                color=text.colour or "#000000", fontproperties=font,
                ha="left", va="bottom", zorder=8)

    # ── 표제란 (지시서 7-2: 자체 서식, 생성 출처 명시) ──
    titles = list(bound.title) if bound else []
    head_y = 1.0 - margin_y
    fig.text(margin_x, head_y, titles[0] if titles else model.source.stem,
             fontsize=9, weight="bold", va="top", ha="left", fontproperties=font)
    subtitle = " · ".join(titles[1:]) if len(titles) > 1 else ""
    fig.text(margin_x, head_y - 5.0 / page[1], subtitle, fontsize=5.6, va="top",
             ha="left", fontproperties=font)
    fig.text(margin_x, head_y - 9.5 / page[1], section or model.source.stem,
             fontsize=5.6, va="top", ha="left", fontproperties=font)
    right = 1.0 - margin_x
    fig.text(right, head_y, "ISO 계통도 — FNCADnet 모듈 D 생성",
             fontsize=6, va="top", ha="right", fontproperties=font)
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    origin = f"원본 {model.source.name}" + (f" + {bound.document.source.name}" if bound else "")
    fig.text(right, head_y - 5.0 / page[1], f"{origin} · {stamp}",
             fontsize=5.0, va="top", ha="right", fontproperties=font)
    fig.text(right, head_y - 9.5 / page[1],
             f"관로 {link.name} / 노드 {node.name}",
             fontsize=5.0, va="top", ha="right", fontproperties=font)
    fig.add_artist(Rectangle((margin_x, ax_top + legend_total), 1.0 - 2 * margin_x, 0.0006,
                             transform=fig.transFigure,
                             facecolor="#333333", edgecolor="none"))

    lane = ax_top
    if link.name != "None":
        _draw_legend(fig, (margin_x, lane, (1.0 - 2 * margin_x) * 0.46, legend_h),
                     f"{link.name} ({link_fmt.symbol})", link_edges, link_fmt, font)
        lane += legend_h
    if node.name != "None":
        _draw_legend(fig, (margin_x, lane, (1.0 - 2 * margin_x) * 0.46, legend_h),
                     f"{node.name} ({node_fmt.symbol})", node_edges, node_fmt, font)

    fig.text(margin_x, margin_y, "PIPENET 이 계산한 결과를 옮겨 그린 도면이다. "
             "값은 결과 XML 원문이며 이 프로그램이 계산하거나 보간하지 않는다.",
             fontsize=4.6, va="bottom", ha="left", color="#555555", fontproperties=font)
    fig.text(right, margin_y, "Page 1 of 1", fontsize=4.6, va="bottom", ha="right",
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
            f"관로의 연장선에 도면 span 의 {_STUB_SPAN_RATIO * 100:.3f}% 만큼 내고 점선·"
            f"흰 삼각으로 갈라 표시했다: {', '.join(derived_nozzles[:12])}"
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
