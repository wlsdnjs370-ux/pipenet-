"""PipeNet SDF XML parser.

The parser maps an existing PipeNet ``.sdf`` XML file into the package's
intermediate :class:`~pipenet_converter.models.PipeNetwork` representation. It
does not validate hydraulic correctness; validation is handled by a separate
module.
"""

from __future__ import annotations

from pathlib import Path
import xml.etree.ElementTree as ET

from pipenet_converter.models import (
    Equipment,
    Fitting,
    Node,
    Nozzle,
    Pipe,
    PipeNetwork,
    Valve,
)


def parse_sdf(path: str | Path) -> PipeNetwork:
    """Read a PipeNet SDF file from disk and parse it into a ``PipeNetwork``."""
    sdf_path = Path(path)
    return parse_sdf_text(sdf_path.read_text(encoding="utf-8"))


def parse_sdf_text(xml_text: str) -> PipeNetwork:
    """Parse PipeNet SDF XML text into a ``PipeNetwork``."""
    root = ET.fromstring(xml_text)
    network_spray = _find_first(root, "Network-spray")
    if network_spray is None:
        raise ValueError("SDF XML does not contain a <Network-spray> element.")

    title_element = _find_direct_child(network_spray, "Title")
    title = (title_element.text or "").strip() if title_element is not None else ""
    network = PipeNetwork(title=title or "Untitled SDF")

    for node_element in _iter_named(network_spray, "Node"):
        network.add_node(_parse_node(node_element))

    for pipe_element in _iter_named(network_spray, "Pipe"):
        network.add_pipe(_parse_pipe(pipe_element))

    for nozzle_element in _iter_named(network_spray, "Nozzle"):
        network.add_nozzle(_parse_nozzle(nozzle_element))

    for valve_element in _iter_valve_elements(network_spray):
        network.add_valve(_parse_valve(valve_element))

    return network


def _parse_node(element: ET.Element) -> Node:
    label = _attr(element, "label", "")
    io_node = _attr(element, "io-node", "No")
    position = _find_direct_child(element, "Position")

    return Node(
        node_id=label,
        x=_float_attr(position, "x", 0.0),
        y=_float_attr(position, "y", 0.0),
        z=_float_attr(element, "elevation", 0.0),
        node_type=io_node,
        source="sdf",
        metadata={"io_node": io_node},
    )


def _parse_pipe(element: ET.Element) -> Pipe:
    return Pipe(
        pipe_id=_attr(element, "label", ""),
        from_node=_attr(element, "input", ""),
        to_node=_attr(element, "output", ""),
        diameter_m=_float_attr(element, "bore", 0.0),
        length_m=_float_attr(element, "length", 0.0),
        rise_m=_float_attr(element, "rise", 0.0),
        c_factor=_float_attr(element, "roughness-or-c", 120.0),
        status=_attr(element, "status", "normal"),
        fittings=_parse_fittings(element),
        equipment=_parse_equipment(element),
        waypoints=_parse_waypoints(element),
    )


def _parse_fittings(pipe_element: ET.Element) -> list[Fitting]:
    fittings_parent = _find_direct_child(pipe_element, "Fittings")
    if fittings_parent is None:
        return []

    return [
        Fitting(
            fitting_type=_attr(fitting_element, "type", ""),
            count=_int_attr(fitting_element, "count", 0),
        )
        for fitting_element in _iter_named(fittings_parent, "Fitting")
    ]


def _parse_equipment(pipe_element: ET.Element) -> list[Equipment]:
    components_parent = _find_direct_child(pipe_element, "Components")
    if components_parent is None:
        return []

    return [
        Equipment(
            equipment_id=_attr(equipment_element, "label", ""),
            description=_attr(equipment_element, "description", ""),
            equivalent_length_m=_float_attr(equipment_element, "equivalent-length", 0.0),
            rel_position=_float_attr(equipment_element, "rel-position", 0.5),
        )
        for equipment_element in _iter_named(components_parent, "Equipment")
    ]


def _parse_waypoints(pipe_element: ET.Element) -> list[tuple[float, float]]:
    waypoints_parent = _find_direct_child(pipe_element, "Waypoints")
    if waypoints_parent is None:
        return []

    return [
        (
            _float_attr(position_element, "x", 0.0),
            _float_attr(position_element, "y", 0.0),
        )
        for position_element in _iter_named(waypoints_parent, "Position")
    ]


def _parse_nozzle(element: ET.Element) -> Nozzle:
    flow_define = _find_direct_child(element, "Flow-define")
    library_item = _find_direct_child(element, "Library-item")

    return Nozzle(
        nozzle_id=_attr(element, "label", ""),
        input_node=_attr(element, "input", ""),
        output_node=_attr(element, "output", ""),
        flow_m3s=_float_attr(flow_define, "flow", 0.0),
        status=_int_attr(element, "status", 1),
        library_item=(library_item.text or "").strip() if library_item is not None else "SP-HEAD",
    )


def _iter_valve_elements(network_spray: ET.Element) -> list[ET.Element]:
    valve_tags = {"Elastomeric-valve", "Pressure-drop-valve"}
    return [element for element in network_spray.iter() if _local_name(element.tag) in valve_tags]


def _parse_valve(element: ET.Element) -> Valve:
    label = _attr(element, "label", "")
    valve_type = _attr(element, "type", _local_name(element.tag))
    return Valve(
        valve_id=label,
        input_node=_attr(element, "input", ""),
        output_node=_attr(element, "output", ""),
        valve_type=valve_type,
        target_value=_optional_float_attr(element, "target-value"),
        metadata={"sdf_tag": _local_name(element.tag)},
    )


def _find_first(root: ET.Element, name: str) -> ET.Element | None:
    return next(_iter_named(root, name), None)


def _find_direct_child(root: ET.Element, name: str) -> ET.Element | None:
    for child in root:
        if _local_name(child.tag) == name:
            return child
    return None


def _iter_named(root: ET.Element, name: str) -> iter[ET.Element]:
    return (element for element in root.iter() if _local_name(element.tag) == name)


def _local_name(tag: str) -> str:
    if "}" in tag:
        return tag.rsplit("}", 1)[1]
    return tag


def _attr(element: ET.Element | None, name: str, default: str) -> str:
    if element is None:
        return default
    return str(element.attrib.get(name, default))


def _float_attr(element: ET.Element | None, name: str, default: float) -> float:
    value = _optional_float_attr(element, name)
    return default if value is None else value


def _optional_float_attr(element: ET.Element | None, name: str) -> float | None:
    if element is None or name not in element.attrib:
        return None
    return float(element.attrib[name])


def _int_attr(element: ET.Element | None, name: str, default: int) -> int:
    if element is None or name not in element.attrib:
        return default
    return int(element.attrib[name])
