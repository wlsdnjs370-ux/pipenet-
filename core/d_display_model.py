# -*- coding: utf-8 -*-
"""모듈 D — D1: PIPENET SDF 의 표시 계열을 손실 없이 읽는다.

`kfp_sdf_converter.parse_sdf` 는 수리 계산 입력을 만드는 경로라 표시 정보를 버린다.
좌표를 ÷1000 하고(`_detect_sdf_coord_unit`), `@` 노드를 건너뛰고, Graphics·Text-element·
Equipment 를 아예 읽지 않는다. 그 경로의 계약은 모듈 A/B 가 의존하므로 바꾸지 않고,
표시 계열은 이 모듈이 따로 읽는다. 모듈 A 의 elevation_m(수리) / display_z_m(표시)
분리와 같은 취지다.

Position x,y 는 PIPENET 이 이미 등각 투영해 둔 화면 좌표다. 여기서는 스케일도
회전도 하지 않는다 — 원본 값 그대로 싣는다.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from pathlib import Path

# ── CP949 모지바케 ──────────────────────────────────────────────────────────
# SDF 는 encoding="UTF-8" 로 선언하지만 한글은 CP949 바이트를 그대로 UTF-8 로 한 번 더
# 감싼 형태로 들어 있다. 그래서 strict UTF-8 디코드는 성공하고(코퍼스 341개 전량 성공)
# 내용만 조용히 깨진다 — '49Ãþ ÃµÀå' 이 실제로는 '49층 천장'.
# 되돌리는 길은 latin-1 로 바이트를 복원해 CP949 로 다시 읽는 것뿐이다.


def restore_cp949(text: str) -> str:
    """모지바케면 한글로 되돌리고, 아니면 원문 그대로 준다."""
    if not text or text.isascii():
        return text
    try:
        return text.encode("latin-1").decode("cp949")
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text


# ── 단위 ────────────────────────────────────────────────────────────────────
# 결과 XML 의 수치는 전부 SI 원단위다. 어떤 단위로 보여줄지는 파일의 <Units> 가
# 정하므로, 계수를 코드에 박지 않고 파일이 지정한 토큰으로 이 표를 찾는다.
# 표에 없는 토큰은 추측하지 않고 경고로 올린다 (환산 없이 SI 그대로 표시).
UNIT_FACTORS: dict[str, tuple[float, str]] = {
    # 길이
    "metres": (1.0, "m"),
    "millimetres": (1000.0, "mm"),
    "centimetres": (100.0, "cm"),
    # 압력 — kgf/cm2 = 98066.5 Pa (표준중력 9.80665 m/s2 정의값)
    "kg-f-cm2-g": (1.0 / 98066.5, "kg/cm²"),
    "bar-g": (1.0 / 1e5, "bar"),
    "pascals": (1.0, "Pa"),
    # 유량 — m3/s → L/min
    "l-min": (60000.0, "L/min"),
    "l-s": (1000.0, "L/s"),
    "m3-h": (3600.0, "m³/h"),
    # 그 밖
    "kelvin": (1.0, "K"),
    "m-s": (1.0, "m/s"),
    "kg-m3": (1.0, "kg/m³"),
    "lb-ft3": (1.0 / 16.018463373960142, "lb/ft³"),
    "pa-s": (1.0, "Pa·s"),
    "watts": (1.0, "W"),
}


# <Links> 아래에서 관로·노즐이 아닌 링크. 이것도 두 노드를 잇는 실제 구간이다.
DEVICE_TAGS = ("Pump-fan", "Elastomeric-valve", "Pressure-loss")


@dataclass(frozen=True)
class UnitSpec:
    quantity: str          # 'Pressure', 'Length', ...
    token: str             # <Pressure-unit unit="kg-f-cm2-g">
    precision: int
    factor: float | None   # SI → 표시. 모르는 토큰이면 None
    symbol: str

    def format(self, si_value: float | None) -> str:
        """SI 값을 이 단위·자릿수로. 값이 없으면 빈칸 — 추정하지 않는다."""
        if si_value is None or self.factor is None:
            return ""
        return f"{si_value * self.factor:.{self.precision}f}"


@dataclass(frozen=True)
class TextElement:
    x: float
    y: float
    text: str
    colour: str
    style: str
    typeface: str
    typesize: float


@dataclass(frozen=True)
class SourceDisplay:
    """PIPENET 이 SDF 에 저장해 둔 자기 화면 설정. 속성이 없으면 None 으로 둔다 —
    기본값을 지어내면 원본이 무엇을 시켰는지 다시 알 수 없다."""
    link_labels: bool | None = None
    node_labels: bool | None = None
    flow_arrows: bool | None = None
    link_values: bool | None = None
    node_values: bool | None = None
    grid: str = ""


@dataclass(frozen=True)
class Equipment:
    label: str
    description: str        # 'A/V', 'FX', 'PUMP' …
    equivalent_length_m: float | None
    rel_position: float | None   # 관로 시작점 기준 0..1 위치


@dataclass(frozen=True)
class DisplayNode:
    label: str
    x: float
    y: float
    elevation_m: float | None
    io_node: str            # 'No' | 'Input'
    virtual: bool           # '@' — 노즐 대기 방출단. 결과 XML 에 없다.


@dataclass(frozen=True)
class DisplayPipe:
    label: str
    input_node: str
    output_node: str
    bore_m: float | None
    length_m: float | None
    rise_m: float | None
    c_factor: float | None
    status: str
    pipe_set: int
    waypoints: tuple[tuple[float, float], ...]
    equipment: tuple[Equipment, ...]
    fittings: tuple[tuple[str, int], ...]


@dataclass(frozen=True)
class DisplayDevice:
    """관로가 아닌 링크 — 펌프·감압밸브·고정손실. 노드를 잇고 있으므로 도면에 그려야 한다.

    코퍼스 336 SDF 에서 Pump-fan 19 · Elastomeric-valve 179 · Pressure-loss 322 개가
    나온다. 관로처럼 bore·length·waypoints 가 없어 두 끝 노드 사이 직선이다.
    """
    kind: str               # 'Pump-fan' | 'Elastomeric-valve' | 'Pressure-loss'
    label: str
    input_node: str
    output_node: str
    description: str        # Library-pump 이름 등. 없으면 빈 문자열
    on: bool
    attributes: dict[str, str]   # 파일 원문 그대로 (target-value, exponent …)


@dataclass(frozen=True)
class DisplayNozzle:
    label: str
    input_node: str
    output_node: str
    on: bool
    flow_m3s: float | None
    library_item: str


@dataclass
class DisplayModel:
    source: Path
    project_version: str
    fluid_density_kgm3: float | None = None
    units: dict[str, UnitSpec] = field(default_factory=dict)
    grid: dict[str, str] = field(default_factory=dict)
    display_options: dict[str, dict[str, str]] = field(default_factory=dict)
    link_scheme: str = ""
    node_scheme: str = ""
    link_schemes: tuple[str, ...] = ()
    node_schemes: tuple[str, ...] = ()
    texts: tuple[TextElement, ...] = ()
    nodes: tuple[DisplayNode, ...] = ()
    pipes: tuple[DisplayPipe, ...] = ()
    nozzles: tuple[DisplayNozzle, ...] = ()
    devices: tuple[DisplayDevice, ...] = ()
    warnings: list[str] = field(default_factory=list)

    @property
    def source_display(self) -> SourceDisplay:
        """원본이 시켜 둔 표시 설정. `label-all` 은 개별 스위치를 덮어쓴다."""
        label = self.display_options.get("Label-display", {})
        result = self.display_options.get("Results-display", {})
        every = _flag(label, "label-all")
        def named(key: str) -> bool | None:
            got = _flag(label, key)
            if every:
                return True
            return got
        return SourceDisplay(
            link_labels=named("link-labels"),
            node_labels=named("node-labels"),
            flow_arrows=_flag(result, "flow-arrows"),
            link_values=_flag(result, "links"),
            node_values=_flag(result, "nodes"),
            grid=self.grid.get("grid", ""),
        )

    @property
    def real_nodes(self) -> tuple[DisplayNode, ...]:
        """결과 XML 과 대응하는 노드 — `@` 가상 노드를 뺀 것."""
        return tuple(n for n in self.nodes if not n.virtual)

    def bounds(self) -> tuple[float, float, float, float] | None:
        """(min_x, min_y, max_x, max_y) — 노드·웨이포인트·텍스트를 모두 포함."""
        xs: list[float] = [n.x for n in self.nodes]
        ys: list[float] = [n.y for n in self.nodes]
        for p in self.pipes:
            xs.extend(w[0] for w in p.waypoints)
            ys.extend(w[1] for w in p.waypoints)
        xs.extend(t.x for t in self.texts)
        ys.extend(t.y for t in self.texts)
        if not xs:
            return None
        return min(xs), min(ys), max(xs), max(ys)


def _flag(attrs: dict[str, str], key: str) -> bool | None:
    raw = attrs.get(key)
    return None if raw is None else raw.strip() not in ("0", "", "false")


def _num(el: ET.Element, key: str) -> float | None:
    raw = el.get(key)
    if raw is None:
        return None
    try:
        return float(raw)
    except ValueError:
        return None


def _scheme_names(container: ET.Element | None) -> tuple[str, ...]:
    """<Pipe-bore-scheme> → 'Pipe bore'. select-name 과 같은 표기다."""
    if container is None:
        return ()
    out = []
    for child in container:
        tag = child.tag
        if tag.endswith("-scheme"):
            tag = tag[: -len("-scheme")]
        out.append(tag.replace("-", " "))
    return tuple(out)


def load_display_model(path: str | Path) -> DisplayModel:
    """SDF → DisplayModel. 좌표는 원본 그대로, 한글은 복원해서 싣는다."""
    src = Path(path)
    root = ET.parse(str(src)).getroot()
    model = DisplayModel(source=src, project_version=root.get("version", ""))

    fluid = root.find(".//Attributes/Fluid-fixed-user")
    if fluid is not None:
        model.fluid_density_kgm3 = _num(fluid, "density")

    units_el = root.find(".//Attributes/Units")
    if units_el is None:
        model.warnings.append("<Units> 없음 — 결과 수치를 SI 원단위로 둔다")
    else:
        for child in units_el:
            quantity = child.tag[: -len("-unit")] if child.tag.endswith("-unit") else child.tag
            token = child.get("unit", "")
            known = UNIT_FACTORS.get(token)
            if known is None:
                model.warnings.append(f"모르는 단위 토큰 '{token}' ({child.tag}) — 환산하지 않는다")
            factor, symbol = known if known else (None, token)
            try:
                precision = int(child.get("precision", "1"))
            except ValueError:
                precision = 1
            model.units[quantity] = UnitSpec(quantity, token, precision, factor, symbol)

    graphics = root.find("Graphics")
    if graphics is None:
        model.warnings.append("<Graphics> 없음 — 표시 스킴과 도면 텍스트를 알 수 없다")
    else:
        opts = graphics.find("Display-options")
        if opts is not None:
            model.display_options = {c.tag: dict(c.attrib) for c in opts}
            model.grid = model.display_options.get("Grid-options", {})
        links = graphics.find("Link-schemes")
        nodes_s = graphics.find("Node-schemes")
        model.link_scheme = links.get("select-name", "") if links is not None else ""
        model.node_scheme = nodes_s.get("select-name", "") if nodes_s is not None else ""
        model.link_schemes = _scheme_names(links)
        model.node_schemes = _scheme_names(nodes_s)

        texts = []
        for te in graphics.findall("Text-element"):
            pos = te.find("Position")
            attrs = te.find("Text-attributes")
            body = te.find("Text")
            if pos is None:
                continue
            a = attrs.attrib if attrs is not None else {}
            texts.append(TextElement(
                x=_num(pos, "x") or 0.0,
                y=_num(pos, "y") or 0.0,
                text=restore_cp949((body.text or "") if body is not None else ""),
                colour=a.get("colour", "#000000"),
                style=a.get("style", ""),
                typeface=restore_cp949(a.get("typeface", "")),
                typesize=float(a.get("typesize", "0") or 0),
            ))
        model.texts = tuple(texts)

    nodes = []
    for nd in root.iterfind(".//Nodes/Node"):
        pos = nd.find("Position")
        label = nd.get("label", "")
        nodes.append(DisplayNode(
            label=label,
            x=(_num(pos, "x") or 0.0) if pos is not None else 0.0,
            y=(_num(pos, "y") or 0.0) if pos is not None else 0.0,
            elevation_m=_num(nd, "elevation"),
            io_node=nd.get("io-node", "No"),
            virtual="@" in label,
        ))
    model.nodes = tuple(nodes)

    pipes = []
    for set_idx, pipe_set in enumerate(root.iterfind(".//Links/Pipe-set")):
        for pp in pipe_set.findall("Pipe"):
            way = pp.find("Waypoints")
            equipment = tuple(
                Equipment(
                    label=eq.get("label", ""),
                    description=restore_cp949(eq.get("description", "")),
                    equivalent_length_m=_num(eq, "equivalent-length"),
                    rel_position=_num(eq, "rel-position"),
                )
                for eq in pp.iterfind("Components/Equipment")
            )
            fittings = tuple(
                (restore_cp949(ft.get("type", "")), int(ft.get("count", "0") or 0))
                for ft in pp.iterfind("Fittings/Fitting")
            )
            pipes.append(DisplayPipe(
                label=pp.get("label", ""),
                input_node=pp.get("input", ""),
                output_node=pp.get("output", ""),
                bore_m=_num(pp, "bore"),
                length_m=_num(pp, "length"),
                rise_m=_num(pp, "rise"),
                c_factor=_num(pp, "roughness-or-c"),
                status=pp.get("status", ""),
                pipe_set=set_idx,
                waypoints=tuple(
                    (_num(p, "x") or 0.0, _num(p, "y") or 0.0)
                    for p in (way.findall("Position") if way is not None else [])
                ),
                equipment=equipment,
                fittings=fittings,
            ))
    model.pipes = tuple(pipes)

    nozzles = []
    for nz in root.iterfind(".//Links/Nozzle"):
        flow = nz.find("Flow-define")
        lib = nz.find("Library-item")
        nozzles.append(DisplayNozzle(
            label=nz.get("label", ""),
            input_node=nz.get("input", ""),
            output_node=nz.get("output", ""),
            on=nz.get("status", "1") == "1",
            flow_m3s=_num(flow, "flow") if flow is not None else None,
            library_item=restore_cp949((lib.text or "").strip() if lib is not None else ""),
        ))
    model.nozzles = tuple(nozzles)

    devices = []
    for links in root.iterfind(".//Links"):
        for dev in links:
            if dev.tag not in DEVICE_TAGS:
                continue
            named = [restore_cp949((c.text or "").strip()) for c in dev if (c.text or "").strip()]
            devices.append(DisplayDevice(
                kind=dev.tag,
                label=dev.get("label", ""),
                input_node=dev.get("input", ""),
                output_node=dev.get("output", ""),
                description=named[0] if named else "",
                on=dev.get("status", "1") == "1",
                attributes=dict(dev.attrib),
            ))
    model.devices = tuple(devices)

    known_nodes = {n.label for n in model.nodes}
    linked = list(model.pipes) + list(model.devices)
    dangling = sorted(
        {e for p in linked for e in (p.input_node, p.output_node) if e not in known_nodes}
    )
    if dangling:
        model.warnings.append(f"노드 목록에 없는 링크 끝점 {len(dangling)}개: {dangling[:5]}")

    return model
