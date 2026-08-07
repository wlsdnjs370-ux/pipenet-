# -*- coding: utf-8 -*-
"""모듈 D — D2: PIPENET 결과 XML(XDSET)을 표시 모델에 결합한다.

D1 이 읽은 SDF(화면·형상)와 결과 XML(수치)을 라벨로 이어 붙여, D3 가 도면에 얹을
값과 D5 가 조판할 표를 한 자리에서 내놓는다.

**단위** — 작업지시서 3.5-가는 "결과 XML 수치는 전부 SI 원단위" 라고 했지만 실물은
다르다. 코퍼스 262 짝 중 259 개는 압력·유량이 이미 표시 단위(kg/cm², L/min)이고
3 개만 SI(Pa, m³/s)다. 관경은 어느 쪽이든 mm 다. 파일 안의 ``RESOURCE name="Units"``
선언은 262 개 전부 ``pressure='kg/cm2 G'`` 로 같아서 이 차이를 구별하지 못한다 —
즉 선언이 거짓말을 한다.

그래서 선언을 믿지 않고 SDF(항상 SI)를 자로 삼아 배율을 파일마다 실측한다. 실측이
안 되는 양만 선언으로 되돌아가고, 그것도 없으면 환산하지 않고 리포트에 올린다
(지시서 7-4·7-5).

**세트 불일치** — SDF 와 XML 이 다른 모델 버전일 수 있다. 세트 전체를 버리지 않고
교집합만 값으로 싣되, 어느 쪽에만 있는 라벨인지 리포트에 전량 나열한다.
"""
from __future__ import annotations

import math
import statistics
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from core.d_display_model import (
    UNIT_FACTORS,
    DisplayModel,
    DisplayPipe,
    load_display_model,
    restore_cp949,
)

# 표준중력 (kgf 정의값). 표고차를 압력으로 환산하는 앵커에만 쓴다.
G0 = 9.80665

# XML 의 단위 표기를 SDF 토큰으로 옮긴다. 환산 계수 자체는 D1 의 UNIT_FACTORS 한 곳에만
# 둔다 — 같은 수를 두 파일에 적어 두지 않기 위해서다.
XML_UNIT_ALIASES: dict[str, str] = {
    "m": "metres", "metres": "metres",
    "mm": "millimetres", "milli.m.": "millimetres",
    "kg/cm2 G": "kg-f-cm2-g", "kg/cm2": "kg-f-cm2-g", "kg/cm²": "kg-f-cm2-g",
    "bar G": "bar-g", "bar": "bar-g",
    "Pa": "pascals", "N/m2": "pascals",
    "l/min": "l-min", "lit/min": "l-min",
    "l/s": "l-s", "lit/sec": "l-s",
    "cu.m/hr": "m3-h",
    "m/s": "m-s", "m/sec": "m-s",
    "kg/cu.m": "kg-m3", "lb/cu.ft": "lb-ft3",
    "K": "kelvin", "Kelvin": "kelvin",
    "Pa.s": "pa-s",
    "W": "watts", "Watts": "watts",
}

# FIELD 의 unit 속성 → <Units> 선언 키. 이름이 다른 것만 적는다.
_DECLARATION_KEY = {"positiveVolFlow": "flowrate"}

# 스칼라 배율로 환산할 수 없는 복합 단위 — 손대지 않고 파일 표기 그대로 싣는다.
# (K-factor 는 l/min·kg/cm²^½, resistanceFactor 는 kg/cm²/m³/sec^exp 같은 꼴이다.)
COMPOUND_KINDS = frozenset({
    "K-factor", "flowCoefficient", "resistanceFactor",
    "pumpConstant", "pumpLinear", "pumpQuadratic", "pumpCubic",
})

# 실측 배율의 허용 산포. 표시 단위 파일은 값이 반올림돼 있어 완전히 같지는 않다.
_SPREAD_TOL = 1e-3
# 배율에 동의하지 않아도 되는 표본의 비율. 세트 안에서 관로 한둘만 다르게 편집돼 있는
# 일이 실제로 있다 (2. Pipenet_hand 짝의 관로 '52'). 그 하나에 배율이 무너지면 안 된다.
_OUTLIER_TOL = 0.1
# 배율을 정하는 데 필요한 최소 표본. 짝이 어긋난 세트는 라벨이 한둘만 우연히 겹치는데,
# 그 하나로 배율을 정하면 어긋남이 배율에 흡수돼 조용히 틀린 수치가 나온다.
_MIN_SAMPLES = 3
# 다만 압력 앵커는 XML 안에서만 잰다 (표고 ↔ 정수두손실). 우연한 라벨 겹침이라는
# 실패 자체가 없고, 표본 수는 망의 모양이 정한다 — 층을 오르내리지 않는 평면망은
# 라이저가 하나뿐이라 표본도 하나다 (2. Pipenet_auto: 136 관로 중 134 개가 표고 0).
# 여기에 최소 표본을 걸면 거짓말하는 선언으로 물러나 압력이 98066 배 부푼다.
INTERNAL_ANCHORS = frozenset({"pressure", "pressureDifference"})
# 기준면 편차의 허용 산포(Pa). kg/cm² 를 소수 4 자리로 찍으면 눈금이 ~10 Pa 다.
_DATUM_TOL_PA = 50.0


def normalize_label(raw: str) -> str:
    """SDF 와 XML 의 라벨 표기를 맞춘다 — XML 은 대문자에 `/0` 접미를 붙인다."""
    return raw.split("/")[0].strip().upper()


# ── XDSET 읽기 ──────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class XField:
    name: str
    datatype: str
    unit: str       # 'pressure', 'length', … 빈 문자열이면 무차원
    arraysize: str


@dataclass(frozen=True)
class XTable:
    path: str                       # 'Materials/Schedule[1]/Pipe lengths'
    name: str
    fields: tuple[XField, ...]
    rows: tuple[dict[str, Any], ...]

    def indexed(self, key: str = "Label") -> dict[str, dict[str, Any]]:
        """정규화 라벨 → 행. 같은 라벨이 겹치면 첫 행을 남긴다."""
        out: dict[str, dict[str, Any]] = {}
        for row in self.rows:
            label = row.get(key)
            if isinstance(label, str):
                out.setdefault(normalize_label(label), row)
        return out


@dataclass
class XdsetDocument:
    """결과 XML 원문 — 값은 파일이 쓴 단위 그대로다."""
    source: Path
    info: dict[str, str] = field(default_factory=dict)
    unit_declarations: dict[str, str] = field(default_factory=dict)
    settings: dict[str, dict[str, str]] = field(default_factory=dict)
    tables: dict[str, XTable] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    def table(self, name: str) -> XTable | None:
        """이름이 같은 표 중 첫 번째. 없으면 None."""
        return _by_name(self.tables, name)


def _by_name(tables: dict[str, XTable], name: str) -> XTable | None:
    for tbl in tables.values():
        if tbl.name == name:
            return tbl
    return None


def _cell(text: str | None, datatype: str) -> Any:
    """TD 한 칸. 값이 없으면 None — 빈칸으로 남기고 추정하지 않는다."""
    raw = (text or "").strip()
    if not raw:
        return None
    if datatype == "char":
        return restore_cp949(raw)
    if datatype == "boolean":
        return raw not in ("0", "false", "False")
    try:
        value = float(raw)
    except ValueError:
        return restore_cp949(raw)
    if math.isnan(value) or math.isinf(value):
        return None
    return value


def _read_table(el: ET.Element, path: str) -> XTable:
    fields = tuple(
        XField(
            name=f.get("name", ""),
            datatype=f.get("datatype", "char"),
            unit=f.get("unit", ""),
            arraysize=f.get("arraysize", "1"),
        )
        for f in el.findall("FIELD")
    )
    rows = tuple(
        {fld.name: _cell(td.text, fld.datatype) for fld, td in zip(fields, tr)}
        for tr in el.iter("TR")
    )
    return XTable(path=path, name=el.get("name", ""), fields=fields, rows=rows)


def read_xdset(path: str | Path) -> XdsetDocument:
    """결과 XML → XdsetDocument. 단위 환산은 하지 않는다."""
    src = Path(path)
    root = ET.parse(str(src)).getroot()
    doc = XdsetDocument(source=src)
    if root.tag != "XDSET":
        doc.warnings.append(f"XDSET 이 아니다 — 루트가 <{root.tag}>")

    for info in root.findall("INFO"):
        doc.info[info.get("name", "")] = restore_cp949(info.get("value", ""))

    seen: dict[str, int] = {}

    def walk(el: ET.Element, prefix: list[str]) -> None:
        for child in el:
            if child.tag == "RESOURCE":
                name = child.get("name", "?")
                params = {
                    p.get("name", ""): restore_cp949(p.get("value", ""))
                    for p in child
                    if p.tag in ("INFO", "PARAM")
                }
                if params:
                    doc.settings.setdefault(name, {}).update(params)
                if name == "Units":
                    doc.unit_declarations.update(params)
                walk(child, prefix + [name])
            elif child.tag == "TABLE":
                base = "/".join(prefix + [child.get("name", "?")])
                n = seen.get(base, 0)
                seen[base] = n + 1
                key = base if n == 0 else f"{base}[{n}]"
                doc.tables[key] = _read_table(child, key)

    walk(root, [])
    if not doc.unit_declarations:
        doc.warnings.append('<RESOURCE name="Units"> 없음 — 선언 단위를 알 수 없다')
    return doc


# ── 단위 배율 ───────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class UnitScale:
    """SI → 이 파일의 표기로 가는 배율. 값 환산은 raw / factor 다."""
    kind: str
    declared_token: str
    declared_factor: float | None
    measured_factor: float | None
    samples: int
    factor: float | None
    source: str          # 'measured' | 'declared' | 'none'


def _measure(ratios: list[float]) -> tuple[float | None, int, int]:
    """배율 표본 → (중앙값, 표본 수, 중앙값에 동의하지 않는 표본 수)."""
    usable = [r for r in ratios if r and math.isfinite(r)]
    if not usable:
        return None, 0, 0
    median = statistics.median(usable)
    if median == 0.0:
        return None, len(usable), len(usable)
    outliers = sum(1 for r in usable if abs(r / median - 1.0) > _SPREAD_TOL)
    return median, len(usable), outliers


def _snap(measured: float) -> float | None:
    """실측을 정의 계수로 되돌린다. 어느 계수도 아니면 None.

    표시 단위 파일은 값이 반올림돼 있어 실측이 1.0 대신 0.99998 로 나온다. 그 값을
    나눗셈에 그대로 쓰면 오차가 모든 수치에 곱해진다 — 계수는 정의값이어야 한다.
    그리고 단위 배율은 정의값 중 하나일 수밖에 없으므로, 어느 것도 아닌 실측은
    단위가 특이한 게 아니라 SDF 와 XML 이 다른 모델이라는 신호다.
    """
    for factor, _symbol in UNIT_FACTORS.values():
        if abs(measured / factor - 1.0) <= _SPREAD_TOL:
            return factor
    return None


def _accept(measured: float | None, samples: int, outliers: int, minimum: int = _MIN_SAMPLES) -> bool:
    return (
        measured is not None
        and samples >= minimum
        and outliers <= _OUTLIER_TOL * samples
        and _snap(measured) is not None
    )


def _declared_factor(token: str) -> float | None:
    known = UNIT_FACTORS.get(XML_UNIT_ALIASES.get(token, token))
    return known[0] if known else None


def _anchor_ratios(
    doc: XdsetDocument, model: DisplayModel, density: float | None
) -> dict[str, list[float]]:
    """SDF(항상 SI)를 자로 삼아 양별 배율 표본을 모은다."""
    ratios: dict[str, list[float]] = {}
    sdf_pipes = {normalize_label(p.label): p for p in model.pipes}
    pin = doc.table("Pipes-input")

    if pin is not None:
        for row in pin.rows:
            pipe = sdf_pipes.get(normalize_label(row.get("Label") or ""))
            if pipe is None:
                continue
            bore, length = row.get("Nominal bore"), row.get("Length")
            if pipe.bore_m and isinstance(bore, float):
                ratios.setdefault("diameter", []).append(bore / pipe.bore_m)
            if pipe.length_m and isinstance(length, float):
                ratios.setdefault("length", []).append(length / pipe.length_m)

    # 압력은 SDF 에 대응하는 값이 없다. 대신 XML 안에서 표고차와 정수두손실을 잇는다 —
    # Static head loss = ρ·g·Δz 이고 Δz(length)는 위에서 이미 잰 양이다.
    pres = doc.table("Pipes-results")
    length_med, length_n, length_bad = _measure(ratios.get("length", []))
    if pres is not None and pin is not None and density and _accept(length_med, length_n, length_bad):
        length_med = _snap(length_med)
        rise_by_label = {
            normalize_label(r.get("Label") or ""): r.get("Elevation")
            for r in pin.rows
        }
        # 부속이 달린 관로는 펌프 등이 압력을 더해 항등식이 깨지므로 앵커에서 뺀다.
        plain = {k for k, p in sdf_pipes.items() if not p.equipment}
        for row in pres.rows:
            label = normalize_label(row.get("Label") or "")
            rise, head = rise_by_label.get(label), row.get("Static head loss")
            if label not in plain or not isinstance(rise, float) or not isinstance(head, float):
                continue
            rise_m = rise / length_med
            if abs(rise_m) < 0.5:      # 짧은 표고차는 반올림 잡음이 배율을 흔든다
                continue
            ratios.setdefault("pressureDifference", []).append(head / (density * G0 * rise_m))

    if pres is not None and density:
        for row in pres.rows:
            value = row.get("Density")
            if isinstance(value, float):
                ratios.setdefault("density", []).append(value / density)

    nin = doc.table("Spray-nozzles-input")
    if nin is not None:
        sdf_nozzles = {normalize_label(z.label): z for z in model.nozzles}
        for row in nin.rows:
            nozzle = sdf_nozzles.get(normalize_label(row.get("Label") or ""))
            flow = row.get("Required flow")
            if nozzle and nozzle.flow_m3s and isinstance(flow, float):
                ratios.setdefault("flowrate", []).append(flow / nozzle.flow_m3s)

    return ratios


def _resolve_scales(
    doc: XdsetDocument, model: DisplayModel, warnings: list[str]
) -> dict[str, UnitScale]:
    """양별로 실측 배율을 먼저 쓰고, 못 재면 파일 선언으로 물러난다."""
    kinds = {f.unit for t in doc.tables.values() for f in t.fields if f.unit}
    ratios = _anchor_ratios(doc, model, model.fluid_density_kgm3)
    # 게이지압과 압력차는 기준면만 다르고 눈금은 같다 — 아래 항등식 검사로 확인한다.
    if "pressureDifference" in ratios:
        ratios.setdefault("pressure", list(ratios["pressureDifference"]))

    scales: dict[str, UnitScale] = {}
    for kind in sorted(kinds):
        token = doc.unit_declarations.get(_DECLARATION_KEY.get(kind, kind), "")
        declared = _declared_factor(token) if token else None
        if kind in COMPOUND_KINDS:
            scales[kind] = UnitScale(kind, token, None, None, 0, None, "none")
            continue

        measured, samples, outliers = _measure(ratios.get(kind, []))
        minimum = 1 if kind in INTERNAL_ANCHORS else _MIN_SAMPLES
        if _accept(measured, samples, outliers, minimum):
            if outliers:
                warnings.append(f"{kind} 배율에 어긋나는 표본 {outliers}/{samples}개 — 중앙값을 쓴다")
            measured = _snap(measured)
        else:
            # 버린 실측이 어차피 선언과 같으면 잃은 것이 없다 — 경고하지 않는다.
            agrees = (
                measured is not None and declared is not None
                and abs(measured / declared - 1.0) <= _SPREAD_TOL
            )
            if measured is not None and not agrees:
                warnings.append(
                    f"{kind} 배율을 실측으로 못 정한다 "
                    f"(표본 {samples}개 중 {outliers}개 불일치, 중앙값 {measured:.6g})"
                )
            measured = None

        if measured is not None:
            if declared is not None and abs(measured / declared - 1.0) > _SPREAD_TOL:
                warnings.append(
                    f"{kind} 선언 '{token}' 과 실측이 어긋난다 "
                    f"(선언 {declared:.6g} / 실측 {measured:.6g}) — 실측을 쓴다"
                )
            scales[kind] = UnitScale(kind, token, declared, measured, samples, measured, "measured")
        elif declared is not None:
            scales[kind] = UnitScale(kind, token, declared, None, samples, declared, "declared")
        else:
            warnings.append(f"{kind} 단위를 못 정했다 (선언 '{token}') — 환산하지 않는다")
            scales[kind] = UnitScale(kind, token, None, None, samples, None, "none")
    return scales


def _to_si(tables: dict[str, XTable], scales: dict[str, UnitScale]) -> dict[str, XTable]:
    factors = {k: s.factor for k, s in scales.items() if s.factor not in (None, 1.0)}
    if not factors:
        return dict(tables)
    out: dict[str, XTable] = {}
    for key, tbl in tables.items():
        scaled = {f.name: factors[f.unit] for f in tbl.fields if f.unit in factors}
        if not scaled:
            out[key] = tbl
            continue
        rows = tuple(
            {
                name: (value / scaled[name] if name in scaled and isinstance(value, float) else value)
                for name, value in row.items()
            }
            for row in tbl.rows
        )
        out[key] = XTable(path=tbl.path, name=tbl.name, fields=tbl.fields, rows=rows)
    return out


# ── 결합 ────────────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class NodeResult:
    label: str
    elevation_m: float | None
    pressure_pa: float | None      # 게이지 — 기준면을 관로 결과에 맞춰 놓았다


@dataclass(frozen=True)
class PipeResult:
    label: str
    input_node: str
    output_node: str
    nominal_bore_m: float | None
    length_m: float | None
    rise_m: float | None
    c_factor: float | None
    fitting_equivalent_length_m: float | None
    inlet_pressure_pa: float | None
    outlet_pressure_pa: float | None
    friction_loss_pa: float | None
    static_head_loss_pa: float | None
    velocity_ms: float | None
    flow_m3s: float | None


@dataclass(frozen=True)
class NozzleResult:
    label: str
    input_node: str
    nozzle_type: str
    k_factor: float | None
    required_flow_m3s: float | None
    calculated_flow_m3s: float | None
    inlet_pressure_pa: float | None
    elevation_m: float | None
    deviation_percent: float | None
    on: bool | None


@dataclass
class JoinReport:
    """어느 라벨이 어느 쪽에만 있었는지 — 전량 나열한다."""
    sdf_only_nodes: tuple[str, ...] = ()
    xml_only_nodes: tuple[str, ...] = ()
    sdf_only_pipes: tuple[str, ...] = ()
    xml_only_pipes: tuple[str, ...] = ()
    sdf_only_nozzles: tuple[str, ...] = ()
    xml_only_nozzles: tuple[str, ...] = ()
    # 라벨은 같은데 관경·길이·표고·끝점이 다른 관로 — 이름만 맞고 내용이 다르다.
    divergent_pipes: tuple[str, ...] = ()
    duplicate_labels: tuple[str, ...] = ()

    @property
    def matched(self) -> bool:
        return not any((
            self.sdf_only_nodes, self.xml_only_nodes,
            self.sdf_only_pipes, self.xml_only_pipes,
            self.sdf_only_nozzles, self.xml_only_nozzles,
            self.divergent_pipes,
        ))

    def summary(self) -> str:
        parts = []
        for name, sdf_only, xml_only in (
            ("노드", self.sdf_only_nodes, self.xml_only_nodes),
            ("관로", self.sdf_only_pipes, self.xml_only_pipes),
            ("노즐", self.sdf_only_nozzles, self.xml_only_nozzles),
        ):
            if sdf_only or xml_only:
                parts.append(f"{name} SDF 만 {len(sdf_only)}개 / XML 만 {len(xml_only)}개")
        if self.divergent_pipes:
            parts.append(f"내용이 다른 관로 {len(self.divergent_pipes)}개")
        return "; ".join(parts) if parts else "라벨 전량 일치"


@dataclass
class BoundModel:
    """표시 모델 + 결과 수치. 수치는 전부 SI 다 (표시 단위는 D1 의 UnitSpec 이 붙인다)."""
    model: DisplayModel
    document: XdsetDocument
    scales: dict[str, UnitScale] = field(default_factory=dict)
    tables: dict[str, XTable] = field(default_factory=dict)
    datum_offset_pa: float | None = None
    nodes: dict[str, NodeResult] = field(default_factory=dict)
    pipes: dict[str, PipeResult] = field(default_factory=dict)
    nozzles: dict[str, NozzleResult] = field(default_factory=dict)
    report: JoinReport = field(default_factory=JoinReport)
    warnings: list[str] = field(default_factory=list)

    @property
    def title(self) -> tuple[str, ...]:
        """표제란 제목 4 줄 — 빈 줄은 뺀다."""
        return tuple(
            v for k in ("Title-1", "Title-2", "Title-3", "Title-4")
            if (v := self.document.info.get(k, "").strip())
        )

    def table(self, name: str) -> XTable | None:
        return _by_name(self.tables, name)


def _diverges(pipe: DisplayPipe, row: dict[str, Any]) -> bool:
    """라벨은 같은데 SDF 와 XML 의 관로 내용이 다른가 — 둘 다 값이 있는 항목만 본다."""
    for shown, key in ((pipe.bore_m, "Nominal bore"), (pipe.length_m, "Length"),
                       (pipe.rise_m, "Elevation")):
        found = row.get(key)
        if shown is None or not isinstance(found, float):
            continue
        if abs(found - shown) > 1e-6 + _SPREAD_TOL * max(abs(found), abs(shown)):
            return True
    for shown, key in ((pipe.input_node, "Input node"), (pipe.output_node, "Output node")):
        found = row.get(key)
        if isinstance(found, str) and normalize_label(found) != normalize_label(shown):
            return True
    return False


def _measure_datum(
    node_rows: dict[str, dict[str, Any]], pipe_rows: tuple[dict[str, Any], ...]
) -> tuple[float | None, float]:
    """노드 압력표와 관로 결과의 기준면 차이. 일정하지 않으면 (None, 산포)."""
    offsets: list[float] = []
    for row in pipe_rows:
        for node_key, pressure_key in (("Input node", "Inlet pressure"),
                                       ("Output node", "Outlet pressure")):
            node = row.get(node_key)
            pressure = row.get(pressure_key)
            if not isinstance(node, str) or not isinstance(pressure, float):
                continue
            at_node = node_rows.get(normalize_label(node), {}).get("Pressure")
            if isinstance(at_node, float):
                offsets.append(at_node - pressure)
    if not offsets:
        return None, math.inf
    median = statistics.median(offsets)
    spread = max(abs(o - median) for o in offsets)
    return (median if spread <= _DATUM_TOL_PA else None), spread


def bind_results(sdf: str | Path | DisplayModel, xml: str | Path) -> BoundModel:
    """SDF 표시 모델과 결과 XML 을 라벨로 잇는다 — 교집합만 값으로 싣는다."""
    model = sdf if isinstance(sdf, DisplayModel) else load_display_model(sdf)
    doc = read_xdset(xml)
    warnings = list(doc.warnings)
    if model.fluid_density_kgm3 is None:
        warnings.append("SDF 에 유체 밀도가 없다 — 압력 배율을 실측하지 못한다")

    scales = _resolve_scales(doc, model, warnings)
    tables = _to_si(doc.tables, scales)

    pin = _by_name(tables, "Pipes-input")
    pres = _by_name(tables, "Pipes-results")
    npr = _by_name(tables, "Node pressures")
    nin = _by_name(tables, "Spray-nozzles-input")
    nres = _by_name(tables, "Spray-nozzles-results")
    for name, tbl in (("Pipes-results", pres), ("Node pressures", npr)):
        if tbl is None:
            warnings.append(f"결과 XML 에 '{name}' 표가 없다")

    node_rows = npr.indexed() if npr else {}
    pipe_rows = pres.rows if pres else ()
    datum, datum_spread = _measure_datum(node_rows, pipe_rows)
    if datum is None and node_rows and pipe_rows:
        warnings.append(
            f"노드 압력 기준면이 일정하지 않다 (산포 {datum_spread:.4g} Pa) — 보정하지 않는다"
        )
    offset = datum or 0.0

    # ── 라벨 조인 ──
    def collect(labels: list[str], what: str) -> dict[str, str]:
        out: dict[str, str] = {}
        for raw in labels:
            key = normalize_label(raw)
            if key in out and out[key] != raw:
                duplicates.append(f"{what} '{out[key]}' / '{raw}' → '{key}'")
            out.setdefault(key, raw)
        return out

    duplicates: list[str] = []
    sdf_nodes = collect([n.label for n in model.real_nodes], "노드")
    sdf_pipes = collect([p.label for p in model.pipes], "관로")
    sdf_nozzles = collect([z.label for z in model.nozzles], "노즐")

    pipe_input_rows = pin.indexed() if pin else {}
    pipe_result_rows = pres.indexed() if pres else {}
    nozzle_input_rows = nin.indexed() if nin else {}
    nozzle_result_rows = nres.indexed() if nres else {}
    xml_pipes = set(pipe_input_rows) | set(pipe_result_rows)
    xml_nozzles = set(nozzle_input_rows) | set(nozzle_result_rows)

    nodes = {
        key: NodeResult(
            label=sdf_nodes[key],
            elevation_m=row.get("Elevation"),
            pressure_pa=(p - offset) if isinstance(p := row.get("Pressure"), float) else None,
        )
        for key, row in node_rows.items()
        if key in sdf_nodes
    }

    sdf_pipe_by_key = {normalize_label(p.label): p for p in model.pipes}
    pipes = {}
    divergent = []
    for key in sorted(set(sdf_pipes) & xml_pipes):
        a = pipe_input_rows.get(key, {})
        b = pipe_result_rows.get(key, {})
        if a and _diverges(sdf_pipe_by_key[key], a):
            divergent.append(key)
        pipes[key] = PipeResult(
            label=sdf_pipes[key],
            input_node=b.get("Input node") or a.get("Input node") or "",
            output_node=b.get("Output node") or a.get("Output node") or "",
            nominal_bore_m=a.get("Nominal bore"),
            length_m=a.get("Length"),
            rise_m=a.get("Elevation"),
            c_factor=a.get("C-factor"),
            fitting_equivalent_length_m=a.get("Fittings - equ. length"),
            inlet_pressure_pa=b.get("Inlet pressure"),
            outlet_pressure_pa=b.get("Outlet pressure"),
            friction_loss_pa=b.get("Friction loss"),
            static_head_loss_pa=b.get("Static head loss"),
            velocity_ms=b.get("Velocity"),
            flow_m3s=b.get("Flow"),
        )

    nozzles = {}
    for key in sorted(set(sdf_nozzles) & xml_nozzles):
        a = nozzle_input_rows.get(key, {})
        b = nozzle_result_rows.get(key, {})
        nozzles[key] = NozzleResult(
            label=sdf_nozzles[key],
            input_node=a.get("Input node") or b.get("Input label") or "",
            nozzle_type=a.get("Nozzle type") or "",
            k_factor=a.get("K-factor"),
            required_flow_m3s=b.get("Required flow", a.get("Required flow")),
            calculated_flow_m3s=b.get("Calculated flow"),
            inlet_pressure_pa=b.get("Inlet pressure"),
            elevation_m=b.get("Elevation"),
            deviation_percent=b.get("Deviation"),
            on=a.get("Nozzle on", b.get("Nozzle on")),
        )

    report = JoinReport(
        sdf_only_nodes=tuple(sorted(set(sdf_nodes) - set(node_rows))),
        xml_only_nodes=tuple(sorted(set(node_rows) - set(sdf_nodes))),
        sdf_only_pipes=tuple(sorted(set(sdf_pipes) - xml_pipes)),
        xml_only_pipes=tuple(sorted(xml_pipes - set(sdf_pipes))),
        sdf_only_nozzles=tuple(sorted(set(sdf_nozzles) - xml_nozzles)),
        xml_only_nozzles=tuple(sorted(xml_nozzles - set(sdf_nozzles))),
        divergent_pipes=tuple(divergent),
        duplicate_labels=tuple(duplicates),
    )
    if duplicates:
        warnings.append(f"정규화 후 겹치는 라벨 {len(duplicates)}건 — 첫 것만 쓴다")
    if not report.matched:
        warnings.append(f"SDF 와 XML 이 어긋난다 ({report.summary()}) — 교집합만 싣는다")

    return BoundModel(
        model=model,
        document=doc,
        scales=scales,
        tables=tables,
        datum_offset_pa=datum,
        nodes=nodes,
        pipes=pipes,
        nozzles=nozzles,
        report=report,
        warnings=warnings,
    )
