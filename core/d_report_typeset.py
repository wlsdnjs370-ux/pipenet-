# -*- coding: utf-8 -*-
"""모듈 D5 — 수리계산서 PDF.

결과 XML 이 적어 둔 수치를 옮겨 고정폭으로 짜고, 워드를 거치지 않고 바로 PDF 로
찍는다 (지시서 D5·7-6). 절의 순서와 열 이름은 PIPENET 이 내는 계산서를 따르되,
표제란과 고지 문구는 우리 것으로 쓴다 (7-2). 환산 계수는 파일의 <Units> 선언에서만
오고 (7-5), XML 에 없는 값은 빈칸으로 두고 리포트에 남긴다 (7-4).
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence

import matplotlib
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.figure import Figure
from matplotlib.font_manager import FontProperties

from core.d_result_binder import BoundModel, XTable, bind_results

matplotlib.rcParams["pdf.fonttype"] = 42

# 고정폭이 무너지면 표가 무너진다. 한글은 뒤 글꼴로 흘려보낸다 (matplotlib 글꼴 대체).
_MONO_STACK = ("DejaVu Sans Mono", "Malgun Gothic", "Gulim", "Batang")

# A4 가로. 워드 계산서가 120 칸이라 세로로는 표가 잘린다.
PAGE_INCHES = (11.69, 8.27)
MARGIN_INCHES = 0.5
_MAX_PT = 8.5
# DejaVu Sans Mono 의 글자 폭은 em 의 0.602 배다 (advanceWidth 1233 / unitsPerEm 2048).
_ADVANCE = 1233.0 / 2048.0
_LINE_SPACING = 1.25

# FLOW IN PIPES 끝 칸의 표식. 워드 원문에서 이 표식이 붙은 관로 31 개가 SPECIAL EQUIPMENT
# 표의 관로 31 개와 정확히 같다 — 유속 초과 표식이 아니다 (유속 초과는 WARNINGS 절에만 있다).
_FLAG_EQUIPMENT = "E"


# ── Fortran 수치 서식 ───────────────────────────────────────────────────────


def _quantize(value: Decimal, decimals: int) -> Decimal:
    """소수 자리에서 반올림한다. 5 는 올린다 — Fortran 서식이 그렇게 찍는다.

    파이썬의 % 서식은 짝수로 맞춘다 (banker's). 워드 원문과 마지막 자리가 갈린다:
    81.8153586 을 두 자리로 줄이면 Fortran 은 81.82, 파이썬은 81.81 을 낸다.
    """
    return value.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)


def fortran_g(value: Any, sig: int = 4) -> str:
    """Fortran G 편집. 유효숫자를 고정하고, 자릿수를 벗어나면 지수로 넘어간다."""
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return "" if value is None else str(value)
    if not math.isfinite(value):
        return ""
    if value == 0.0:
        return f"{0.0:.{sig - 1}f}"
    exponent = math.floor(math.log10(abs(value))) + 1
    # 0.1 이상 10^sig 미만이면 소수점 표기. 그 밖은 지수로 넘어간다.
    if 0 <= exponent <= sig:
        shown = _quantize(Decimal(value), sig - exponent)
        # 반올림이 자릿수를 하나 늘리면 (9.9999 → 10.00) 소수 자리를 다시 센다.
        if abs(shown) >= 10 ** exponent:
            exponent += 1
            if exponent > sig:
                return _fortran_sci(value, sig)
            shown = _quantize(Decimal(value), sig - exponent)
        decimals = sig - exponent
        # 소수점 이하가 없어도 점은 찍는다 — 정수처럼 보이면 유효숫자를 오해한다.
        return f"{shown:.{decimals}f}" if decimals else f"{shown:.0f}."
    return _fortran_sci(value, sig)


def _fortran_sci(value: float, digits: int) -> str:
    """가수를 1 이상 10 미만으로 두는 지수 표기 (9.8397E-02)."""
    exponent = math.floor(math.log10(abs(value)))
    mant = _quantize(Decimal(value).scaleb(-exponent), digits)
    if abs(mant) >= 10:
        exponent += 1
        mant = _quantize(Decimal(value).scaleb(-exponent), digits)
    return f"{mant:.{digits}f}E{exponent:+03d}"


def fortran_e(value: Any, digits: int = 5) -> str:
    """Fortran E 편집. 가수를 0.1 이상 1 미만으로 두는 그 표기다 (0.10000E+01)."""
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return "" if value is None else str(value)
    if not math.isfinite(value):
        return ""
    if value == 0.0:
        return f"{0.0:.{digits}f}E+00"
    exponent = math.floor(math.log10(abs(value))) + 1
    mant = _quantize(Decimal(value).scaleb(-exponent), digits)
    if abs(mant) >= 1:
        exponent += 1
        mant = _quantize(Decimal(value).scaleb(-exponent), digits)
    return f"{mant:.{digits}f}E{exponent:+03d}"


def fixed(value: Any, decimals: int = 4) -> str:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return "" if value is None else str(value)
    if not math.isfinite(value):
        return ""
    return f"{_quantize(Decimal(value), decimals):.{decimals}f}"


def integer(value: Any) -> str:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        return "" if value is None else str(value)
    if not math.isfinite(value):
        return ""
    return f"{_quantize(Decimal(value), 0):.0f}"


def text(value: Any) -> str:
    return "" if value is None else str(value)


FORMATS: dict[str, Callable[[Any], str]] = {
    "g4": fortran_g,
    "e5": fortran_e,
    "f4": lambda v: fixed(v, 4),
    "f2": lambda v: fixed(v, 2),
    "int": integer,
    "text": text,
}


# ── 절 구성 ─────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class Column:
    """열 하나. 이름은 PIPENET 계산서의 문구 그대로, 단위는 파일 선언 그대로."""
    title: str
    kind: str = ""          # 단위 종류. 빈 문자열이면 환산도 단위 표기도 없다
    fmt: str = "text"
    width: int = 12
    align: str = ">"
    unit_note: str = ""     # 선언에 없는 복합 단위를 직접 적어야 할 때만
    absolute: bool = False  # 값이 절대압 기준면에 있다 — 표시 전에 기준면 차이를 뺀다


@dataclass(frozen=True)
class Block:
    """쪽에 앉히는 최소 단위. 넘치면 title/header 를 다시 찍고 body 만 이어 간다."""
    title: str
    header: tuple[str, ...] = ()
    body: tuple[str, ...] = ()
    tail: tuple[str, ...] = ()


@dataclass
class ReportResult:
    path: Path
    pages: int = 0
    blocks: int = 0
    rows: int = 0
    font_pt: float = 0.0
    velocity_flags: tuple[str, ...] = ()
    nozzle_flags: tuple[str, ...] = ()
    warnings: list[str] = field(default_factory=list)


# ── 값 꺼내기 ───────────────────────────────────────────────────────────────


class Sheet:
    """결합 모델 위에 얹는 표기 계층. SI 를 파일이 선언한 단위로 되돌린다."""

    def __init__(self, bound: BoundModel) -> None:
        self.bound = bound
        self.warnings: list[str] = []
        self._factors: dict[str, float | None] = {}
        for kind, scale in bound.scales.items():
            self._factors[kind] = scale.declared_factor

    def factor(self, kind: str) -> float | None:
        if not kind:
            return 1.0
        if kind not in self._factors:
            self.warnings.append(f"'{kind}' 단위 선언이 없다 — SI 그대로 싣는다")
            self._factors[kind] = 1.0
        factor = self._factors[kind]
        if factor is None:
            self.warnings.append(f"'{kind}' 의 환산 계수를 파일이 정하지 않는다 — 값을 그대로 싣는다")
            self._factors[kind] = 1.0
            return 1.0
        return factor

    def show(self, value: Any, kind: str, absolute: bool = False) -> Any:
        if not isinstance(value, (int, float)) or isinstance(value, bool):
            return value
        factor = self.factor(kind)
        return self.gauge(value, absolute) * (factor if factor is not None else 1.0)

    def gauge(self, value: float, absolute: bool = True) -> float:
        """절대압을 결과표의 기준면으로 되돌린다.

        한 파일 안에 기준면이 둘 있다. 노드 압력표와 노즐의 최소·최대 작동압은 절대압으로
        적히고, 관로·노즐 결과는 이미 게이지압이다 (실측: 노드 66 은 184829 Pa 인데 그 노드에
        달린 노즐 1 의 결과 입구압은 83504 Pa — 차이가 D2 가 잰 기준면 차와 같다).
        빼지 않으면 대기압 한 겹만큼 부풀어 최소압 미달 판정까지 뒤집힌다.
        """
        offset = self.bound.datum_offset_pa
        if not absolute or offset is None:
            return value
        return value - offset

    def unit_token(self, kind: str) -> str:
        return self.bound.document.unit_declarations.get(kind, "")

    def table(self, key: str) -> XTable | None:
        return self.bound.tables.get(key)


def _head_lines(sheet: Sheet, columns: Sequence[Column]) -> tuple[str, str, str]:
    names, units = [], []
    for col in columns:
        names.append(f"{col.title:{col.align}{col.width}}")
        token = col.unit_note or (f"({sheet.unit_token(col.kind)})" if col.kind else "")
        units.append(f"{token:{col.align}{col.width}}")
    first, second = " ".join(names).rstrip(), " ".join(units).rstrip()
    return first, second, "-" * max(len(first), len(second))


def _row_line(sheet: Sheet, columns: Sequence[Column], values: Sequence[Any]) -> str:
    cells = []
    for col, raw in zip(columns, values):
        shown = sheet.show(raw, col.kind, col.absolute) if col.fmt != "text" else raw
        cells.append(f"{FORMATS[col.fmt](shown):{col.align}{col.width}}")
    return " ".join(cells).rstrip()


def _table_block(sheet: Sheet, title: str, columns: Sequence[Column],
                 rows: Iterable[Sequence[Any]], tail: Sequence[str] = ()) -> Block:
    names, units, rule = _head_lines(sheet, columns)
    return Block(
        title=title,
        header=(names, units, rule),
        body=tuple(_row_line(sheet, columns, row) for row in rows),
        tail=tuple(tail),
    )


def _fields(table: XTable | None, names: Sequence[str]) -> Iterable[tuple[Any, ...]]:
    if table is None:
        return ()
    return (tuple(row.get(name) for name in names) for row in table.rows)


# ── 절별 조립 ───────────────────────────────────────────────────────────────


def _control(sheet: Sheet) -> Block:
    settings = sheet.bound.document.settings
    control = settings.get("Control", {})
    velocity = settings.get("Velocity pressure option", {})
    pairs = [
        ("Convergence accuracy", control.get("Convergence accuracy", "")),
        ("Maximum no. of iterations", control.get("Max iterations", "")),
        ("Elevation Check Tolerance",
         f"{control.get('Elevation check tolerance', '')} {sheet.unit_token('length')}".strip()),
        ("Design standard", control.get("Design standard", "")),
        ("Pressure model", control.get("Pressure model", "")),
        ("Velocity pressure model", velocity.get("VP model", "")),
    ]
    return Block("CONTROL INFORMATION",
                 body=tuple(f"    {name:<28}= {value}" for name, value in pairs))


def _fluid(sheet: Sheet) -> Block:
    fluid = sheet.bound.document.settings.get("Fluid", {})
    pairs = [
        ("Fluid Class", f"{fluid.get('Fluid class', '')} ({fluid.get('Fluid name', '')})"),
        ("Density", f"{fluid.get('Density', '')} {sheet.unit_token('density')}".strip()),
        ("Viscosity", f"{fluid.get('Viscosity', '')} {sheet.unit_token('viscosity')}".strip()),
    ]
    return Block("FLUID SYSTEM",
                 body=tuple(f"    {name:<28}= {value}" for name, value in pairs))


def _materials(sheet: Sheet) -> list[tuple[str, dict[str, str]]]:
    return [(key, params) for key, params in sheet.bound.document.settings.items()
            if key.rsplit("/", 1)[-1].split("[")[0] == "Pipe-type"]


def _design(sheet: Sheet) -> Block:
    settings = sheet.bound.document.settings
    control = settings.get("Control", {})
    velocity = settings.get("Velocity pressure option", {})
    lines = [f"    {sheet.bound.document.info.get('Calculator', '')} System", ""]
    lines.append("    Pipe Materials are :")
    lines.append(f"      {'Pipe Type':<24}{'Lining Type':<20}Thickness({sheet.unit_token('diameter')})")
    for _key, params in _materials(sheet):
        thickness = params.get("Thickness", "")
        if thickness.strip().lower() == "nan":
            thickness = ""      # 파일이 값을 적지 않았다 — 채우지 않는다
        lines.append(f"      {params.get('Pipe type', ''):<24}"
                     f"{params.get('Lining type', ''):<20}{thickness}")
    lines.append("")
    lines.append(f"    Design to {control.get('Design standard', '')} Rules")
    lines.append(f"    Using the {control.get('Pressure model', '')} Equation")
    lines.append(f"    Velocity Pressure Model: {velocity.get('VP model', '')}")
    entrance = velocity.get("Pressure loss include entrance", "")
    exit_ = velocity.get("Pressure loss include exit", "")
    lines.append(f"    Pressure loss at entrance: {'Include' if entrance == '1' else 'Ignore'}")
    lines.append(f"    Pressure loss at exit: {'Include' if exit_ == '1' else 'Ignore'}")
    return Block("DESIGN INFORMATION", body=tuple(lines))


# 호칭경은 환산하는 값이 아니라 규격 번호다. 파일도 열 이름에 단위를 박아 둔다
# ('Nominal bore - mm'). 선언 단위를 씌우지 않고 있는 그대로 싣는다.
_NOMINAL_UNIT = "(milli.m.)"

_SIZE_COLUMNS = (
    Column("Nom.Bore", "", "f4", 11, unit_note=_NOMINAL_UNIT),
    Column("Act.Diam", "diameter", "f4", 11),
    Column("Max.Vel.", "velocity", "f4", 11),
)


def _sizes(sheet: Sheet) -> list[Block]:
    blocks: list[Block] = []
    for key, params in _materials(sheet):
        table = sheet.table(f"{key}/Sizes")
        if table is None:
            sheet.warnings.append(f"'{params.get('Pipe type', key)}' 의 규격표가 없다")
            continue
        rows = [r for r in _fields(table, ("Nominal bore", "Actual diameter", "Maximum velocity"))
                if any(v is not None for v in r[1:])]
        block = _table_block(
            sheet,
            f"AVAILABLE PIPE SIZES AND MAXIMUM VELOCITIES — {params.get('Pipe type', '')}"
            f" / {params.get('Lining type', '')}",
            _SIZE_COLUMNS, rows)
        blocks.append(block)
    return blocks


_PIPE_COLUMNS = (
    Column("Pipe Label", "", "text", 12, "<"),
    Column("Input Node", "", "text", 12, "<"),
    Column("Output Node", "", "text", 12, "<"),
    Column("Nom.Bore", "diameter", "g4", 11),
    Column("Length", "length", "g4", 12),
    Column("Elevation", "length", "g4", 12),
    Column("C Factor", "", "g4", 9),
    Column("Fitt.eq.lnth", "length", "g4", 12),
)


def _pipe_configuration(sheet: Sheet) -> Block:
    rows = _fields(sheet.table("Pipes-input/Pipes-input"),
                   ("Label", "Input node", "Output node", "Nominal bore", "Length",
                    "Elevation", "C-factor", "Fittings - equ. length"))
    return _table_block(sheet, "PIPE CONFIGURATION", _PIPE_COLUMNS, rows)


def _fitting_types(table: XTable) -> list[str]:
    """GROUP 이름이 곧 부속 종류다 ('Type 4' → 4)."""
    seen: list[str] = []
    for fld in table.fields:
        if " / " not in fld.name:
            continue
        group = fld.name.split(" / ", 1)[0]
        if group not in seen:
            seen.append(group)
    return seen


def _pipe_fittings(sheet: Sheet) -> Block:
    table = sheet.table("Pipes-input/Pipe-fittings")
    if table is None:
        return Block("PIPE FITTINGS", body=("    결과 XML 에 부속 표가 없다",))
    groups = _fitting_types(table)
    factor = sheet.factor("length") or 1.0
    lines: list[str] = []
    for row in table.rows:
        parts = []
        for group in groups:
            count = row.get(f"{group} / Count")
            if not count:
                continue
            number = group.rsplit(" ", 1)[-1]
            equivalent = row.get(f"{group} / Equiv. length")
            shown = fortran_g(equivalent * factor) if isinstance(equivalent, float) else ""
            parts.append(f"{integer(count):>3} x {number:<3}{shown:>11}")
        if parts:
            lines.append(f" {text(row.get('Label')):<12}" + "   ".join(parts))
    header = (f" {'Pipe':<12}Number x Type  Equivalent Length",
              f" {'Label':<12}               ({sheet.unit_token('length')})",
              "-" * 100)
    return Block("PIPE FITTINGS", header=header, body=tuple(lines))


_NOZZLE_COLUMNS = (
    Column("Nozzle Label", "", "text", 13, "<"),
    Column("Input Node", "", "text", 12, "<"),
    Column("Nozzle Type", "", "text", 12, "<"),
    Column("K-Factor", "", "f4", 11),
    Column("Req Flow", "flowrate", "f4", 13),
    Column("Min Press", "pressure", "e5", 14, absolute=True),
    Column("Max Press", "pressure", "e5", 14, absolute=True),
)


def _nozzle_configuration(sheet: Sheet) -> Block:
    rows = _fields(sheet.table("Spray-nozzles-input/Spray-nozzles-input"),
                   ("Label", "Input node", "Nozzle type", "K-factor", "Required flow",
                    "Minimum pressure", "Maximum pressure"))
    return _table_block(sheet, "NOZZLE CONFIGURATION", _NOZZLE_COLUMNS, rows)


_EQUIPMENT_COLUMNS = (
    Column("Equipment Label", "", "text", 16, "<"),
    Column("Pipe Label", "", "text", 12, "<"),
    Column("Equivalent Length", "length", "e5", 18),
    Column("Description", "", "text", 24, "<"),
)


def _equipment(sheet: Sheet) -> Block:
    rows = _fields(sheet.table("Equipment/Equipment"),
                   ("Label", "Pipe label", "Equivalent length", "Decription"))
    return _table_block(sheet, "SPECIAL EQUIPMENT", _EQUIPMENT_COLUMNS, rows)


def _size_lookup(sheet: Sheet) -> dict[tuple[str, float], tuple[float | None, float | None]]:
    """(재질, 호칭경 mm) → (실지름 SI, 최대유속 SI). 규격표에 적힌 것만 담는다."""
    out: dict[tuple[str, float], tuple[float | None, float | None]] = {}
    for key, params in _materials(sheet):
        table = sheet.table(f"{key}/Sizes")
        if table is None:
            continue
        material = params.get("Pipe type", "")
        for row in table.rows:
            nominal = row.get("Nominal bore")
            if isinstance(nominal, float):
                out[(material, round(nominal, 3))] = (row.get("Actual diameter"),
                                                      row.get("Maximum velocity"))
    return out


_DESIGNED_COLUMNS = (
    Column("Pipe Label", "", "text", 12, "<"),
    Column("Input Node", "", "text", 12, "<"),
    Column("Output Node", "", "text", 12, "<"),
    Column("Flowrate", "flowrate", "f4", 13),
    Column("Pipe Type", "", "text", 12, "<"),
    Column("Act. Bore", "diameter", "f4", 11),
    Column("Nom. Size", "diameter", "f4", 11),
)


def _designed(sheet: Sheet) -> Block:
    pipes = sheet.table("Pipes-input/Pipes-input")
    results = sheet.table("Pipes-results/Pipes-results")
    if pipes is None:
        return Block("DESIGNED DIAMETERS & FLOWRATES", body=("    관로 입력표가 없다",))
    flows = {r.get("Label"): r.get("Flow") for r in (results.rows if results else ())}
    sizes = _size_lookup(sheet)
    nominal_factor = sheet.factor("diameter") or 1.0
    rows = []
    for row in pipes.rows:
        nominal = row.get("Nominal bore")
        material = text(row.get("Type"))
        actual = None
        if isinstance(nominal, float):
            actual = sizes.get((material, round(nominal * nominal_factor, 3)), (None, None))[0]
        rows.append((row.get("Label"), row.get("Input node"), row.get("Output node"),
                     flows.get(row.get("Label")), material, actual, nominal))
    return _table_block(sheet, "DESIGNED DIAMETERS & FLOWRATES", _DESIGNED_COLUMNS, rows,
                        tail=("  Act. Bore 는 재질별 규격표에서 호칭경으로 찾은 값이다.",
                              "  Flowrate 는 관로 결과표의 유량이다 — 결과 XML 에는 관경 결정 단계의",
                              "  유량이 남아 있지 않아, 수렴된 해의 유량을 그대로 싣는다.",
                              "  결과 XML 에 관로 그룹 항이 없어 Pipe Group 열은 싣지 않는다.",))


_FLOW_COLUMNS = (
    Column("Pipe Label", "", "text", 12, "<"),
    Column("Input Node", "", "text", 12, "<"),
    Column("Output Node", "", "text", 12, "<"),
    Column("Nom.Bore", "diameter", "f2", 10),
    Column("Inlet Pr.", "pressure", "g4", 11),
    Column("Outlet Pr.", "pressure", "g4", 11),
    Column("Drop in Pr.", "pressureDifference", "g4", 11),
    Column("Frict. Loss", "pressureDifference", "g4", 11),
    Column("Flowrate", "flowrate", "g4", 11),
    Column("Velocity", "velocity", "g4", 11),
    Column("", "", "text", 1, "<"),
)


def _flow_in_pipes(sheet: Sheet) -> tuple[Block, tuple[str, ...]]:
    results = sheet.table("Pipes-results/Pipes-results")
    pipes = sheet.table("Pipes-input/Pipes-input")
    equipment = sheet.table("Equipment/Equipment")
    if results is None:
        return Block("FLOW IN PIPES", body=("    관로 결과표가 없다",)), ()
    inputs = {r.get("Label"): r for r in (pipes.rows if pipes else ())}
    equipped = {r.get("Pipe label") for r in (equipment.rows if equipment else ())}
    sizes = _size_lookup(sheet)
    nominal_factor = sheet.factor("diameter") or 1.0
    rows, over = [], []
    for row in results.rows:
        label = row.get("Label")
        source = inputs.get(label, {})
        nominal = source.get("Nominal bore")
        inlet, outlet = row.get("Inlet pressure"), row.get("Outlet pressure")
        drop = inlet - outlet if isinstance(inlet, float) and isinstance(outlet, float) else None
        velocity = row.get("Velocity")
        limit = None
        if isinstance(nominal, float):
            limit = sizes.get((text(source.get("Type")), round(nominal * nominal_factor, 3)),
                              (None, None))[1]
        exceeded = (isinstance(velocity, float) and isinstance(limit, float)
                    and abs(velocity) > limit)
        if exceeded:
            over.append(text(label))
        rows.append((label, row.get("Input node"), row.get("Output node"), nominal,
                     inlet, outlet, drop, row.get("Friction loss"), row.get("Flow"),
                     velocity, _FLAG_EQUIPMENT if label in equipped else ""))
    block = _table_block(
        sheet, "FLOW IN PIPES", _FLOW_COLUMNS, rows,
        tail=("  Drop in Pr. 는 Inlet Pr. 에서 Outlet Pr. 을 뺀 값이다.",
              f"  끝의 {_FLAG_EQUIPMENT} 는 부속이 달린 관로를 가리킨다 (SPECIAL EQUIPMENT 참조).",
              "  최대유속을 넘은 관로는 WARNINGS 절에 모아 적는다."))
    return block, tuple(over)


_NOZZLE_FLOW_COLUMNS = (
    Column("Nozzle Label", "", "text", 13, "<"),
    Column("Input Label", "", "text", 13, "<"),
    Column("Inlet Press", "pressure", "e5", 14),
    Column("Req. Flow", "flowrate", "f4", 13),
    Column("Flowrate", "flowrate", "f4", 13),
    Column("% Deviation", "", "f2", 12),
)


def _nozzle_flows(sheet: Sheet) -> tuple[Block, tuple[str, ...]]:
    results = sheet.table("Spray-nozzles-results/Spray-nozzles-results")
    inputs = sheet.table("Spray-nozzles-input/Spray-nozzles-input")
    if results is None:
        return Block("FLOW THROUGH NOZZLES", body=("    노즐 결과표가 없다",)), ()
    minimum = {r.get("Label"): r.get("Minimum pressure") for r in (inputs.rows if inputs else ())}
    rows, under = [], []
    for row in results.rows:
        label = row.get("Label")
        pressure, floor_ = row.get("Inlet pressure"), minimum.get(label)
        # 결과 입구압은 게이지압, 최소 작동압은 절대압이다. 기준면을 맞추지 않고 견주면
        # 대기압 한 겹 차이로 멀쩡한 노즐까지 미달로 잡힌다.
        low = (isinstance(pressure, float) and isinstance(floor_, float)
               and pressure < sheet.gauge(floor_))
        if low:
            under.append(text(label))
        rows.append((label, row.get("Input label"), pressure, row.get("Required flow"),
                     row.get("Calculated flow"), row.get("Deviation")))
    block = _table_block(sheet, "FLOW THROUGH NOZZLES", _NOZZLE_FLOW_COLUMNS, rows,
                         tail=("  최소 작동압에 못 미친 노즐은 WARNINGS 절에 모아 적는다.",
                               "  결과 XML 에 살수밀도 항이 없어 Req. FlowDens / FlowDens 열은 싣지 않는다.",))
    return block, tuple(under)


def _inlets(sheet: Sheet) -> Block:
    table = sheet.table("Input-nodes/Input-nodes")
    columns = (
        Column("Inlet Node", "", "text", 13, "<"),
        Column("Pressure", "pressure", "g4", 11),
        Column("", "", "text", 1, "<"),
        Column("Flowrate", "flowrate", "g4", 11),
        Column("", "", "text", 1, "<"),
        Column("Equivalent K-factor", "K-factor", "f4", 19,
               unit_note=f"({sheet.unit_token('flowrate')}, {sheet.unit_token('pressure')})"),
    )
    rows = []
    for row in (table.rows if table else ()):
        rows.append((row.get("Label"), row.get("Pressure"),
                     "*" if row.get("Pressure spec,") else "",
                     row.get("Flow"), "*" if row.get("Flow spec,") else "",
                     row.get("K-factor")))
    return _table_block(sheet, "FLOW AT INLETS", columns, rows,
                        tail=("  값 뒤의 * 는 그 값이 지정값(specification)임을 뜻한다.",))


def _materials_takeoff(sheet: Sheet) -> list[Block]:
    blocks: list[Block] = []
    lengths = sheet.table("Materials/Schedule/Pipe lengths")
    if lengths is not None:
        columns = (Column("Nom. Size", "", "f4", 11, unit_note=_NOMINAL_UNIT),
                   Column("Tot. Length", "length", "g4", 13))
        blocks.append(_table_block(sheet, "MATERIALS TAKE-OFF / PIPE LENGTHS", columns,
                                   _fields(lengths, ("Nominal bore - mm", "Total length"))))
    nozzles = sheet.table("Materials/Nozzles")
    if nozzles is not None:
        columns = (Column("Type", "", "text", 24, "<"), Column("K-Factor", "", "f4", 11),
                   Column("Number", "", "int", 8))
        blocks.append(_table_block(sheet, "MATERIALS TAKE-OFF / NOZZLES", columns,
                                   _fields(nozzles, ("Type", "K-Factor", "Number"))))
    fittings = sheet.table("Materials/Fittings")
    if fittings is not None:
        kinds = [f.name for f in fittings.fields if f.name != "Fitting Nominal size"]
        columns = [Column("Fitting Nominal Size", "diameter", "f4", 20)]
        columns += [Column(str(n + 1), "", "int", 5) for n in range(len(kinds))]
        legend = tuple(f"  {n + 1} -- {name.strip()}" for n, name in enumerate(kinds))
        blocks.append(_table_block(sheet, "MATERIALS TAKE-OFF / FITTINGS", tuple(columns),
                                   _fields(fittings, ["Fitting Nominal size"] + kinds),
                                   tail=("", "  Fitting Types are :") + legend))
    other = sheet.table("Materials/Other Equipment")
    if other is not None:
        columns = (Column("Type", "", "text", 28, "<"), Column("Number", "", "int", 8))
        blocks.append(_table_block(sheet, "MATERIALS TAKE-OFF / OTHER EQUIPMENT", columns,
                                   _fields(other, ("Type", "Number"))))
    return blocks


_NOTICE = (
    "  이 문서는 PIPENET 이 만든 문서가 아니다. PIPENET 이 남긴 결과 파일에 적힌 수치를",
    "  옮겨 적어 자체 서식으로 짠 것이며, 해석 자체를 다시 하지 않았다. 따라서 여기 실린",
    "  값은 아래 원본 파일이 담고 있는 값과 같고, 그 이상도 이하도 아니다.",
    "",
    "  원본 파일의 해석이 유효한지, 그 결과를 이 현장에 적용해도 되는지는 자격 있는",
    "  기술자가 따로 확인해야 한다.",
)


def _provenance(sheet: Sheet, generated: datetime) -> Block:
    info = sheet.bound.document.info
    lines = [
        f"  결과 파일   : {sheet.bound.document.source.name}",
        f"  도면 파일   : {sheet.bound.model.source.name}",
        f"  해석 시각   : {info.get('Date', '')} {info.get('Time', '')}".rstrip(),
        f"  해석기      : {info.get('Calculator', '')} {info.get('Version', '')}".rstrip(),
        f"  문서 생성   : {generated:%Y-%m-%d %H:%M} (모듈 D5)",
        "",
        *_NOTICE,
    ]
    return Block("생성 출처와 고지", body=tuple(lines))


def _comments(sheet: Sheet, over: Sequence[str], under: Sequence[str]) -> Block:
    lines: list[str] = []
    if over:
        lines.append(f"  최대유속을 넘은 관로 {len(over)}개")
        lines.extend(f"    *** 경고 - 관로 {label} 의 유속이 규격표의 최대유속을 넘는다"
                     for label in over)
        lines.append("")
    if under:
        lines.append(f"  최소 작동압에 못 미친 노즐 {len(under)}개")
        lines.extend(f"    *** 경고 - 노즐 {label} 의 입구압이 최소 작동압보다 낮다"
                     for label in under)
        lines.append("")
    if not lines:
        lines.append("  유속 초과와 최소압 미달 모두 없다.")
    report = sheet.bound.report
    if not report.matched:
        lines.append("  도면과 결과 파일이 어긋난 항목이 있다 — 교집합만 실었다:")
        lines.extend(f"    {w}" for w in sheet.bound.warnings)
    elif sheet.bound.warnings or sheet.warnings:
        lines.append("  읽으면서 남긴 기록:")
        lines.extend(f"    {w}" for w in list(sheet.bound.warnings) + sheet.warnings)
    return Block("WARNINGS", body=tuple(lines))


def build_blocks(bound: BoundModel, generated: datetime) -> tuple[list[Block], Sheet,
                                                                 tuple[str, ...], tuple[str, ...]]:
    """결합 모델 → 계산서의 절들. 순서는 PIPENET 계산서를 따른다."""
    sheet = Sheet(bound)
    flow_block, over = _flow_in_pipes(sheet)
    nozzle_block, under = _nozzle_flows(sheet)
    blocks = [
        _control(sheet),
        _fluid(sheet),
        _design(sheet),
        *_sizes(sheet),
        _pipe_configuration(sheet),
        _pipe_fittings(sheet),
        _nozzle_configuration(sheet),
        _equipment(sheet),
        _designed(sheet),
        flow_block,
        nozzle_block,
        _inlets(sheet),
        *_materials_takeoff(sheet),
        _provenance(sheet, generated),
    ]
    # 경고는 표를 다 채운 뒤에 만든다 — 그 사이에 쌓인 기록까지 싣기 위해서다.
    blocks.append(_comments(sheet, over, under))
    return blocks, sheet, over, under


# ── 쪽 나누기 ───────────────────────────────────────────────────────────────


def paginate(blocks: Sequence[Block], lines_per_page: int) -> list[list[str]]:
    """절이 쪽을 넘으면 제목과 열 머리글을 다시 찍는다 — 이어진 표도 혼자 읽혀야 한다."""
    pages: list[list[str]] = []
    page: list[str] = []

    def flush() -> None:
        nonlocal page
        if page:
            pages.append(page)
            page = []

    for block in blocks:
        rule = "=" * max(len(block.title), 18)
        opening = [block.title, rule, *block.header]
        # 제목만 남기고 쪽이 끝나면 읽을 수 없다. 본문 한 줄은 붙여 둔다.
        if page and len(page) + len(opening) + 1 > lines_per_page:
            flush()
        page.extend(opening)
        for line in block.body:
            if len(page) >= lines_per_page:
                flush()
                page.extend([f"{block.title} (이어짐)", rule, *block.header])
            page.append(line)
        for line in block.tail:
            if len(page) >= lines_per_page:
                flush()
            page.append(line)
        page.append("")
    flush()
    return pages


# ── PDF ─────────────────────────────────────────────────────────────────────


def _type_size(blocks: Sequence[Block]) -> float:
    """가장 긴 줄이 폭 안에 들어오는 글자 크기. 표를 접지 않으려면 이쪽이 양보한다."""
    width_pt = (PAGE_INCHES[0] - 2 * MARGIN_INCHES) * 72.0
    longest = max((len(line) for block in blocks
                   for line in (block.title, *block.header, *block.body, *block.tail)),
                  default=1)
    return min(_MAX_PT, width_pt / (max(longest, 1) * _ADVANCE))


def _lines_per_page(size: float) -> int:
    height_pt = (PAGE_INCHES[1] - 2 * MARGIN_INCHES) * 72.0
    # 맨 윗줄은 표제가 쓴다.
    return max(1, int(height_pt / (size * _LINE_SPACING)) - 2)


def _font() -> FontProperties:
    return FontProperties(family=list(_MONO_STACK))


def render_report(sdf_path: str | Path, xml_path: str | Path, out_pdf: str | Path,
                  *, generated: datetime | None = None) -> ReportResult:
    """SDF + 결과 XML → 수리계산서 PDF. PIPENET 을 띄우지 않는다."""
    bound = bind_results(sdf_path, xml_path)
    stamp = generated or datetime.now()
    blocks, sheet, over, under = build_blocks(bound, stamp)

    size = _type_size(blocks)
    pages = paginate(blocks, _lines_per_page(size))

    out = Path(out_pdf)
    out.parent.mkdir(parents=True, exist_ok=True)
    font = _font()
    left = MARGIN_INCHES / PAGE_INCHES[0]
    top = 1.0 - MARGIN_INCHES / PAGE_INCHES[1]
    title = " · ".join(v for v in bound.title if v)
    with PdfPages(str(out)) as pdf:
        for number, page in enumerate(pages, start=1):
            fig = Figure(figsize=PAGE_INCHES)
            head = (f"{title}    생성 {stamp:%Y-%m-%d}    "
                    f"Page {number} of {len(pages)}")
            fig.text(left, top, head, fontproperties=font, fontsize=size,
                     va="bottom", ha="left")
            fig.text(1.0 - left, top, bound.document.source.name, fontproperties=font,
                     fontsize=size * 0.85, va="bottom", ha="right")
            fig.text(left, top - size * _LINE_SPACING / (PAGE_INCHES[1] * 72.0),
                     "\n".join(page), fontproperties=font, fontsize=size,
                     va="top", ha="left", linespacing=_LINE_SPACING)
            pdf.savefig(fig)

    return ReportResult(path=out, pages=len(pages), blocks=len(blocks),
                        rows=sum(len(b.body) for b in blocks), font_pt=size,
                        velocity_flags=over, nozzle_flags=under,
                        warnings=list(bound.warnings) + sheet.warnings)
