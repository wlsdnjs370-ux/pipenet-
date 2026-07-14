"""답안 SDF 파일 → AnswerNetwork dataclass — 검증 하네스용 역방향 파서.

sdf_writer.py 의 출력 구조를 그대로 역해석:
  Project > Network-spray > Title, Nodes/Node, Links/Pipe-set/Pipe, Nozzle, Elastomeric-valve

bore/length/rise/elevation 의 sdf 단위가 m 인 점에 주의 (sdf_writer 가 그렇게 직렬화).
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable


def _strip_ns(tag: str) -> str:
    """ElementTree tag 에서 namespace prefix 제거."""
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _find_children(elem: ET.Element, name: str) -> list[ET.Element]:
    return [c for c in elem if _strip_ns(c.tag) == name]


def _find_one(elem: ET.Element, name: str) -> ET.Element | None:
    for c in elem:
        if _strip_ns(c.tag) == name:
            return c
    return None


def _ffloat(val: str | None, default: float = 0.0) -> float:
    if val is None or val == "":
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        return default


@dataclass
class AnswerNode:
    label: str
    x: float
    y: float
    elevation_m: float
    io_node: str       # "Input" / "No" / "Output" / ...


@dataclass
class AnswerFitting:
    fitting_type: str
    count: int = 1


@dataclass
class AnswerEquipment:
    label: str
    description: str
    eq_length_m: float
    rel_position: float = 0.5


@dataclass
class AnswerPipe:
    label: str
    input_node: str
    output_node: str
    bore_m: float          # sdf 의 bore 단위 = m (예: 0.1 = 100mm)
    length_m: float
    rise_m: float
    c_factor: float        # roughness-or-c
    status: str = "Normal"
    fittings: list[AnswerFitting] = field(default_factory=list)
    equipment: list[AnswerEquipment] = field(default_factory=list)

    @property
    def bore_mm(self) -> int:
        """직경 mm 단위 (정수 반올림)."""
        return int(round(self.bore_m * 1000))


@dataclass
class AnswerNozzle:
    label: str
    input_node: str
    output_node: str
    flow_m3s: float
    library_item: str
    status: str = "Normal"


@dataclass
class AnswerValve:
    label: str
    input_node: str
    output_node: str
    valve_type: str
    target_value: float | None = None
    sdf_tag: str = "Elastomeric-valve"   # 또는 Pump-fan 등


@dataclass
class AnswerNetwork:
    """답안 PIPENET SDF 의 토폴로지/배관망 데이터."""
    title: str
    nodes: list[AnswerNode]
    pipes: list[AnswerPipe]
    nozzles: list[AnswerNozzle] = field(default_factory=list)
    valves: list[AnswerValve] = field(default_factory=list)
    sdf_version: str = ""
    source_path: str = ""

    # ── 비교용 derived 지표 ──
    @property
    def total_length_m(self) -> float:
        return sum(p.length_m for p in self.pipes)

    @property
    def total_rise_m(self) -> float:
        return sum(p.rise_m for p in self.pipes)

    @property
    def elev_range(self) -> tuple[float, float]:
        if not self.nodes:
            return (0.0, 0.0)
        es = [n.elevation_m for n in self.nodes]
        return (min(es), max(es))

    @property
    def bore_distribution_mm(self) -> Counter:
        """직경(mm) 분포 — 노드/파이프 개수가 아니라 직경별 LINE 개수."""
        return Counter(p.bore_mm for p in self.pipes)

    @property
    def bore_length_distribution_mm(self) -> dict[int, float]:
        """직경(mm) 별 총 length(m) — hydraulic 중요도 가중."""
        out: dict[int, float] = {}
        for p in self.pipes:
            out[p.bore_mm] = out.get(p.bore_mm, 0.0) + p.length_m
        return out

    def find_input_node(self) -> AnswerNode | None:
        for n in self.nodes:
            if n.io_node == "Input":
                return n
        return self.nodes[0] if self.nodes else None

    def find_output_nodes(self) -> list[AnswerNode]:
        return [n for n in self.nodes if n.io_node == "Output"]


def read_sdf(path: str | Path) -> AnswerNetwork:
    """SDF XML 파일 → AnswerNetwork."""
    p = Path(path)
    tree = ET.parse(p)
    root = tree.getroot()
    sdf_version = root.attrib.get("version", "")

    network_spray = _find_one(root, "Network-spray")
    if network_spray is None:
        raise ValueError(f"{p}: <Network-spray> 없음 — 유효한 PIPENET SDF 아님")

    title_el = _find_one(network_spray, "Title")
    title = (title_el.text or "").strip() if title_el is not None else ""

    nodes_el = _find_one(network_spray, "Nodes")
    nodes: list[AnswerNode] = []
    if nodes_el is not None:
        for nd in _find_children(nodes_el, "Node"):
            pos = _find_one(nd, "Position")
            x = _ffloat(pos.attrib.get("x")) if pos is not None else 0.0
            y = _ffloat(pos.attrib.get("y")) if pos is not None else 0.0
            nodes.append(AnswerNode(
                label=nd.attrib.get("label", ""),
                x=x, y=y,
                elevation_m=_ffloat(nd.attrib.get("elevation")),
                io_node=nd.attrib.get("io-node", "No"),
            ))

    links_el = _find_one(network_spray, "Links")
    pipes: list[AnswerPipe] = []
    nozzles: list[AnswerNozzle] = []
    valves: list[AnswerValve] = []

    if links_el is not None:
        # Pipe-set / Pipe
        for ps in _find_children(links_el, "Pipe-set"):
            for pi in _find_children(ps, "Pipe"):
                fittings: list[AnswerFitting] = []
                ff_el = _find_one(pi, "Fittings")
                if ff_el is not None:
                    for ft in _find_children(ff_el, "Fitting"):
                        try:
                            cnt = int(ft.attrib.get("count", "1"))
                        except ValueError:
                            cnt = 1
                        fittings.append(AnswerFitting(
                            fitting_type=ft.attrib.get("type", ""),
                            count=cnt,
                        ))
                equipment: list[AnswerEquipment] = []
                comps_el = _find_one(pi, "Components")
                if comps_el is not None:
                    for eq in _find_children(comps_el, "Equipment"):
                        equipment.append(AnswerEquipment(
                            label=eq.attrib.get("label", ""),
                            description=eq.attrib.get("description", ""),
                            eq_length_m=_ffloat(eq.attrib.get("equivalent-length")),
                            rel_position=_ffloat(eq.attrib.get("rel-position"), 0.5),
                        ))
                pipes.append(AnswerPipe(
                    label=pi.attrib.get("label", ""),
                    input_node=pi.attrib.get("input", ""),
                    output_node=pi.attrib.get("output", ""),
                    bore_m=_ffloat(pi.attrib.get("bore")),
                    length_m=_ffloat(pi.attrib.get("length")),
                    rise_m=_ffloat(pi.attrib.get("rise")),
                    c_factor=_ffloat(pi.attrib.get("roughness-or-c"), 120.0),
                    status=pi.attrib.get("status", "Normal"),
                    fittings=fittings,
                    equipment=equipment,
                ))
        # Nozzle (Links 의 직접 자식)
        for nz in _find_children(links_el, "Nozzle"):
            flow_el = _find_one(nz, "Flow-define")
            lib_el = _find_one(nz, "Library-item")
            nozzles.append(AnswerNozzle(
                label=nz.attrib.get("label", ""),
                input_node=nz.attrib.get("input", ""),
                output_node=nz.attrib.get("output", ""),
                flow_m3s=_ffloat(flow_el.attrib.get("flow")) if flow_el is not None else 0.0,
                library_item=(lib_el.text or "").strip() if lib_el is not None else "",
                status=nz.attrib.get("status", "Normal"),
            ))
        # Elastomeric-valve / Pump-fan (Links 직접 자식, tag 가 valve_type 마다 다름)
        for vc in links_el:
            tag = _strip_ns(vc.tag)
            if tag in ("Elastomeric-valve", "Pump-fan"):
                tv = vc.attrib.get("target-value")
                valves.append(AnswerValve(
                    label=vc.attrib.get("label", ""),
                    input_node=vc.attrib.get("input", ""),
                    output_node=vc.attrib.get("output", ""),
                    valve_type=vc.attrib.get("type", ""),
                    target_value=_ffloat(tv) if tv else None,
                    sdf_tag=tag,
                ))

    return AnswerNetwork(
        title=title,
        nodes=nodes,
        pipes=pipes,
        nozzles=nozzles,
        valves=valves,
        sdf_version=sdf_version,
        source_path=str(p),
    )


def summarize(network: AnswerNetwork) -> dict:
    """진단/로그용 dict 요약."""
    return {
        "title": network.title,
        "nodes": len(network.nodes),
        "pipes": len(network.pipes),
        "nozzles": len(network.nozzles),
        "valves": len(network.valves),
        "total_length_m": round(network.total_length_m, 2),
        "total_rise_m": round(network.total_rise_m, 2),
        "elev_range_m": tuple(round(v, 2) for v in network.elev_range),
        "bore_distribution_mm": dict(sorted(network.bore_distribution_mm.items())),
        "bore_length_distribution_mm": {
            k: round(v, 2) for k, v in sorted(network.bore_length_distribution_mm.items())
        },
        "io_node_distribution": dict(Counter(n.io_node for n in network.nodes)),
    }


if __name__ == "__main__":
    import sys, json
    for arg in sys.argv[1:]:
        net = read_sdf(arg)
        print(f"\n=== {Path(arg).name} ===")
        print(json.dumps(summarize(net), ensure_ascii=False, indent=2, default=str))
