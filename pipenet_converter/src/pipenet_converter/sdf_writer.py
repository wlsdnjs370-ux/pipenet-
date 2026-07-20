"""PipeNet SDF XML writer."""

from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from pipenet_converter.models import Node, Nozzle, Pipe, PipeNetwork, Valve


def network_to_sdf_element(network: PipeNetwork) -> ET.Element:
    """Convert a ``PipeNetwork`` into a minimal PipeNet-like SDF XML element."""
    project = ET.Element("Project", {"version": "1.6  (0)"})
    network_spray = ET.SubElement(project, "Network-spray")
    title = ET.SubElement(network_spray, "Title")
    title.text = network.title
    network_spray.append(_build_nodes_element(network))
    network_spray.append(_build_links_element(network))
    ET.SubElement(project, "Graphics")
    _indent(project)
    return project


def network_to_sdf_text(network: PipeNetwork) -> str:
    """Serialize a ``PipeNetwork`` to SDF XML text."""
    root = network_to_sdf_element(network)
    return ET.tostring(root, encoding="unicode", xml_declaration=False)


def write_sdf(
    network: PipeNetwork,
    output_path: str | Path,
    template_path: str | Path | None = None,
) -> None:
    """Write a ``PipeNetwork`` to an SDF file.

    When ``template_path`` is provided, only ``<Nodes>`` and ``<Links>`` under
    ``<Network-spray>`` are replaced. Other project-level settings are left in
    place where possible.
    """
    output = Path(output_path)
    if template_path is None:
        root = network_to_sdf_element(network)
    else:
        root = ET.parse(template_path).getroot()
        network_spray = _find_first(root, "Network-spray")
        if network_spray is None:
            raise ValueError("Template does not contain a <Network-spray> element.")
        _replace_direct_child(network_spray, "Nodes", _build_nodes_element(network))
        _replace_direct_child(network_spray, "Links", _build_links_element(network))
        title = _find_direct_child(network_spray, "Title")
        if title is None:
            title = ET.Element("Title")
            network_spray.insert(0, title)
        title.text = network.title
        _indent(root)

    tree = ET.ElementTree(root)
    output.parent.mkdir(parents=True, exist_ok=True)
    tree.write(output, encoding="utf-8", xml_declaration=True)


def _build_nodes_element(network: PipeNetwork) -> ET.Element:
    nodes_element = ET.Element("Nodes")
    for node in network.nodes.values():
        nodes_element.append(_build_node_element(node))
    return nodes_element


def _build_node_element(node: Node) -> ET.Element:
    io_node = str(node.metadata.get("io_node", node.node_type))
    node_element = ET.Element(
        "Node",
        {
            "elevation": _format_number(node.z),
            "io-node": io_node,
            "label": node.node_id,
        },
    )
    pos_attrs = {
        "x": _format_number(node.x),
        "y": _format_number(node.y),
    }
    # Position z = 표시 전용 좌표(라이저 입상관 높이·헤드평면 평탄화). elevation 속성
    # (수리 실표고)과 분리. metadata 에 display_z 가 있을 때만 기록 → 단독망/외부 SDF
    # 는 종전대로 z 미기록(변환기가 elevation 으로 fallback).
    display_z = node.metadata.get("display_z")
    if display_z is not None:
        pos_attrs["z"] = _format_number(display_z)
    ET.SubElement(node_element, "Position", pos_attrs)
    return node_element


def _build_links_element(network: PipeNetwork) -> ET.Element:
    links_element = ET.Element("Links")
    pipe_set = ET.SubElement(links_element, "Pipe-set")
    for pipe in network.pipes.values():
        pipe_set.append(_build_pipe_element(pipe))

    for nozzle in network.nozzles.values():
        links_element.append(_build_nozzle_element(nozzle))

    for valve in network.valves.values():
        links_element.append(_build_valve_element(valve))

    return links_element


def _build_pipe_element(pipe: Pipe) -> ET.Element:
    pipe_element = ET.Element(
        "Pipe",
        {
            "bore": _format_number(pipe.diameter_m),
            "input": pipe.from_node,
            "label": pipe.pipe_id,
            "length": _format_number(pipe.length_m),
            "output": pipe.to_node,
            "rise": _format_number(pipe.rise_m),
            "roughness-or-c": _format_number(pipe.c_factor),
            "status": pipe.status,
        },
    )

    if pipe.fittings:
        fittings_element = ET.SubElement(pipe_element, "Fittings")
        for fitting in pipe.fittings:
            ET.SubElement(
                fittings_element,
                "Fitting",
                {
                    "count": str(fitting.count),
                    "type": fitting.fitting_type,
                },
            )

    if pipe.equipment:
        components_element = ET.SubElement(pipe_element, "Components")
        for equipment in pipe.equipment:
            ET.SubElement(
                components_element,
                "Equipment",
                {
                    "description": equipment.description,
                    "equivalent-length": _format_number(equipment.equivalent_length_m),
                    "label": equipment.equipment_id,
                    "rel-position": _format_number(equipment.rel_position),
                },
            )

    if pipe.waypoints:
        waypoints_element = ET.SubElement(pipe_element, "Waypoints", {"symbol-segment": "0"})
        for x, y in pipe.waypoints:
            ET.SubElement(
                waypoints_element,
                "Position",
                {
                    "x": _format_number(x),
                    "y": _format_number(y),
                },
            )

    return pipe_element


def _build_nozzle_element(nozzle: Nozzle) -> ET.Element:
    nozzle_element = ET.Element(
        "Nozzle",
        {
            "input": nozzle.input_node,
            "label": nozzle.nozzle_id,
            "output": nozzle.output_node,
            "status": str(nozzle.status),
        },
    )
    ET.SubElement(nozzle_element, "Flow-define", {"flow": _format_number(nozzle.flow_m3s)})
    library_item = ET.SubElement(nozzle_element, "Library-item")
    library_item.text = nozzle.library_item
    return nozzle_element


def _build_valve_element(valve: Valve) -> ET.Element:
    tag = str(valve.metadata.get("sdf_tag", "Elastomeric-valve"))
    valve_element = ET.Element(
        tag,
        {
            "input": valve.input_node,
            "label": valve.valve_id,
            "output": valve.output_node,
            "type": valve.valve_type,
        },
    )
    if valve.target_value is not None:
        valve_element.set("target-value", _format_number(valve.target_value))
    return valve_element


def _format_number(value: float) -> str:
    return format(float(value), ".6g")


def _replace_direct_child(parent: ET.Element, child_name: str, replacement: ET.Element) -> None:
    for index, child in enumerate(list(parent)):
        if _local_name(child.tag) == child_name:
            parent.remove(child)
            parent.insert(index, replacement)
            return
    parent.append(replacement)


def _find_first(root: ET.Element, name: str) -> ET.Element | None:
    for element in root.iter():
        if _local_name(element.tag) == name:
            return element
    return None


def _find_direct_child(root: ET.Element, name: str) -> ET.Element | None:
    for child in root:
        if _local_name(child.tag) == name:
            return child
    return None


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _indent(element: ET.Element, level: int = 0) -> None:
    indent_text = "\n" + level * "  "
    child_indent_text = "\n" + (level + 1) * "  "
    if len(element):
        if not element.text or not element.text.strip():
            element.text = child_indent_text
        for child in element:
            _indent(child, level + 1)
        if not element.tail or not element.tail.strip():
            element.tail = indent_text
    elif level and (not element.tail or not element.tail.strip()):
        element.tail = indent_text
