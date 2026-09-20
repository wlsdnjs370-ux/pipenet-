"""PIPENET display directions and explicit zero-demand terminal conditions.

These changes leave all pipe/nozzle calculation values and connectivity intact.
Only free, unspecified pipe ends are capped; existing boundary specifications
and component endpoints (including nozzle outlets) are never replaced.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
import xml.etree.ElementTree as ET


@dataclass
class CompatibilityReport:
    """Labels changed by the compatibility pass, for export diagnostics."""

    directed_nozzles: list[str] = field(default_factory=list)
    capped_nodes: list[str] = field(default_factory=list)
    unresolved_nozzles: list[str] = field(default_factory=list)


def prepare_sdf(path: str | Path) -> CompatibilityReport:
    """Write explicit capped ends and noncoincident nozzle display endpoints.

    Vertical direction comes from the real elevation of a terminal head relative
    to its only adjacent pipe node, independent of pipe input/output orientation.
    Horizontal or nonterminal nozzles keep their supplied display direction.
    """
    path = Path(path)
    tree = ET.parse(path)
    root = tree.getroot()
    container = root.find(".//Network-spray/Nodes")
    links = root.find(".//Network-spray/Links")
    if container is None or links is None:
        raise ValueError("SDF에 Nodes 또는 Links가 없습니다.")
    nodes = {n.get("label"): n for n in container.findall("Node")}
    incident: dict[str, list[ET.Element]] = defaultdict(list)
    pipes: dict[str, list[str]] = defaultdict(list)
    for link in links.iter():
        a, b = link.get("input"), link.get("output")
        if a is None or b is None:
            continue
        incident[a].append(link)
        incident[b].append(link)
        if link.tag == "Pipe":
            pipes[a].append(b)
            pipes[b].append(a)
    coords = [n.find("Position") for label, n in nodes.items()
              if label and not label.startswith("@/")]
    xs = [float(p.get("x", 0)) for p in coords if p is not None]
    ys = [float(p.get("y", 0)) for p in coords if p is not None]
    span = max(max(xs, default=0) - min(xs, default=0),
               max(ys, default=0) - min(ys, default=0))
    gap = max(1.0, span * 0.0025)  # display units only, never metres of pipe
    report = CompatibilityReport()
    for nozzle in links.iter("Nozzle"):
        label = nozzle.get("label", "")
        a, b = nozzle.get("input"), nozzle.get("output")
        head = nodes.get(a)
        if head is None or not b:
            raise ValueError(f"노즐 {label}: 연결 절점이 없습니다.")
        outlet = nodes.get(b)
        if outlet is None:
            if not b.startswith("@/"):
                raise ValueError(f"노즐 {label}: 대기 출구 {b}가 없습니다.")
            outlet = ET.SubElement(container, "Node", {
                "label": b, "elevation": head.get("elevation", "0"), "io-node": "No"})
            nodes[b] = outlet
        # A shared or user-specified outlet is not merely a glyph endpoint.
        if not b.startswith("@/") or len(incident[b]) != 1:
            report.unresolved_nozzles.append(label)
            continue
        pos = head.find("Position")
        if pos is None:
            raise ValueError(f"노즐 {label}: 표시 좌표가 없습니다.")
        end = outlet.find("Position")
        if end is None:
            end = ET.SubElement(outlet, "Position", dict(pos.attrib))
        neighbors = pipes[a]
        dz = (float(head.get("elevation", 0)) -
              float(nodes[neighbors[0]].get("elevation", 0))) if len(neighbors) == 1 else 0
        if abs(dz) > 1e-8:
            end.set("x", pos.get("x", "0"))
            end.set("y", f"{float(pos.get('y', 0)) + (gap if dz > 0 else -gap):.12g}")
            report.directed_nozzles.append(label)
        else:
            report.unresolved_nozzles.append(label)
    for label, node in nodes.items():
        connected = incident[label]
        if (node.get("io-node", "No") == "No" and
                node.find("Calculation-spec") is None and
                len(connected) == 1 and connected[0].tag == "Pipe"):
            node.set("io-node", "Output")
            ET.SubElement(node, "Calculation-spec", {"flow": "0"})
            report.capped_nodes.append(label)
    ET.indent(root, space="  ")
    path.write_bytes(b'<?xml version="1.0" encoding="UTF-8"?>\n'
                     b'<!DOCTYPE Project SYSTEM "spray.dtd">\n' +
                     ET.tostring(root, encoding="utf-8"))
    return report
