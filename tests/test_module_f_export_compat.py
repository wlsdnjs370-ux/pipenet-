"""Regression coverage for vertical heads, capped ends and physical KFP export."""
from __future__ import annotations

import json
import math
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "pipenet_converter" / "src"))
sys.path.insert(0, str(ROOT / "core"))
from pipenet_converter.physical_layout import LayoutEdge, physical_layout
from pipenet_converter.sdf_compat import prepare_sdf
from routes.module_f.export_compat import emit_physical_kfp


def _sdf(tmp_path, *, upright=False, reverse=False, outlet=True):
    root = ET.Element("Project")
    network = ET.SubElement(root, "Network-spray")
    nodes = ET.SubElement(network, "Nodes")
    for label, x, y, z, io in [
        ("1", 0, 0, 10, "Input"), ("2", 50, 0, 10, "No"),
        ("3", 50, 0, 10.3 if upright else 9.7, "No"),
        ("4", 55, 0, 10, "No"),
    ]:
        node = ET.SubElement(nodes, "Node", label=label, elevation=str(z), **{"io-node": io})
        ET.SubElement(node, "Position", x=str(x), y=str(y))
        if io == "Input":
            ET.SubElement(node, "Calculation-spec", pressure="201325")
    if outlet:
        n = ET.SubElement(nodes, "Node", label="@/1", elevation="10.3" if upright else "9.7", **{"io-node": "No"})
        ET.SubElement(n, "Position", x="50", y="0")
    links = ET.SubElement(network, "Links")
    ps = ET.SubElement(links, "Pipe-set")
    for label, a, b, length, rise in [("P1", "1", "2", 5, 0),
                                    ("P2", "3" if reverse else "2", "2" if reverse else "3", .3,
                                     (.3 if upright else -.3) * (-1 if reverse else 1)),
                                    ("P3", "2", "4", .5, 0)]:
        ET.SubElement(ps, "Pipe", label=label, input=a, output=b, length=str(length),
                      rise=str(rise), bore="0.025", **{"roughness-or-c": "120"})
    noz = ET.SubElement(links, "Nozzle", label="1", input="3", output="@/1", status="1")
    ET.SubElement(noz, "Flow-define", flow="0.00133333333")
    p = tmp_path / "network.sdf"
    ET.ElementTree(root).write(p, encoding="utf-8")
    return p


def _calculation(root):
    return ([(p.tag, dict(p.attrib), [dict(c.attrib) for c in p])
             for p in root.iter() if p.tag in ("Pipe", "Nozzle")],
            {n.get("label"): n.get("elevation") for n in root.iter("Node")
             if not n.get("label").startswith("@/")})


@pytest.mark.parametrize("upright", [False, True])
@pytest.mark.parametrize("reverse", [False, True])
@pytest.mark.parametrize("outlet", [False, True])
def test_head_direction_and_calculation_preservation(tmp_path, upright, reverse, outlet):
    p = _sdf(tmp_path, upright=upright, reverse=reverse, outlet=outlet)
    before = _calculation(ET.parse(p).getroot())
    report = prepare_sdf(p)
    root = ET.parse(p).getroot()
    nodes = {n.get("label"): n for n in root.iter("Node")}
    a, b = [nodes[k].find("Position") for k in ("3", "@/1")]
    assert float(a.get("x")) == float(b.get("x"))
    assert (float(b.get("y")) > float(a.get("y"))) == upright
    assert float(b.get("y")) != float(a.get("y"))
    assert report.capped_nodes == ["4"]
    assert nodes["4"].get("io-node") == "Output"
    assert nodes["4"].find("Calculation-spec").get("flow") == "0"
    assert nodes["1"].find("Calculation-spec").get("pressure") == "201325"
    assert _calculation(root) == before
    once = p.read_bytes()
    prepare_sdf(p)
    assert p.read_bytes() == once


def test_existing_boundary_and_component_outlet_are_not_capped(tmp_path):
    p = _sdf(tmp_path)
    tree = ET.parse(p)
    n = tree.find(".//Node[@label='4']")
    n.set("io-node", "Output")
    ET.SubElement(n, "Calculation-spec", flow="0.5")
    tree.write(p)
    report = prepare_sdf(p)
    assert report.capped_nodes == []
    tree = ET.parse(p)
    assert tree.find(".//Node[@label='4']/Calculation-spec").get("flow") == "0.5"
    assert tree.find(".//Node[@label='@/1']").get("io-node") == "No"


def test_loop_preserved_and_inconsistent_loop_rejected():
    nodes = {"a": (0, 0, 0), "b": (100, 0, 0), "c": (100, 100, 0), "d": (0, 100, 0)}
    edges = [LayoutEdge(str(i), a, b, 2) for i, (a, b) in enumerate([
        ("a", "b"), ("b", "c"), ("c", "d"), ("d", "a")])]
    coords = physical_layout(nodes, edges)
    assert len(coords) == 4
    assert all(math.dist(coords[e.start], coords[e.end]) == pytest.approx(2) for e in edges)
    edges[-1] = LayoutEdge("bad", "d", "a", 3)
    with pytest.raises(ValueError, match="루프"):
        physical_layout(nodes, edges)


@pytest.mark.parametrize("b,length,message", [((0, 0, 5), 2, "표고차"), ((0, 0, 0), 2, "수평 방향")])
def test_invalid_geometry_is_explicit(b, length, message):
    with pytest.raises(ValueError, match=message):
        physical_layout({"a": (0, 0, 0), "b": b}, [LayoutEdge("P1", "a", "b", length)])


def test_kfp_uses_metres_and_vertical_head_without_rescaling(tmp_path):
    sdf = _sdf(tmp_path)
    prepare_sdf(sdf)
    kfp = tmp_path / "network.kfp"
    emit_physical_kfp(sdf, kfp)
    data = json.loads(kfp.read_text(encoding="utf-8"))
    coords = {key: n["coords"] for key, n in data["nodes_meta_runtime"].items()}
    assert len(coords) == 4
    assert len(data["pipe_data"]) == 3
    for pipe in data["pipe_data"].values():
        assert math.dist(coords[pipe["start"]], coords[pipe["end"]]) == pytest.approx(pipe["length_m"], abs=1e-9)
    assert coords["N2"][:2] == coords["N3"][:2]
    assert coords["N3"][2] == 9.7


def test_pump_becomes_inline_node_without_extra_pipe_or_length(tmp_path):
    sdf = _sdf(tmp_path)
    tree = ET.parse(sdf)
    original = tree.find(".//Node[@label='1']")
    original.set("io-node", "No")
    original.remove(original.find("Calculation-spec"))
    n = ET.SubElement(tree.find(".//Nodes"), "Node", label="9", elevation="10", **{"io-node": "Input"})
    ET.SubElement(n, "Position", x="-20", y="0")
    ET.SubElement(n, "Calculation-spec", pressure="101325")
    ET.SubElement(tree.find(".//Links"), "Pump-fan", input="9", output="1", label="pump", **{
        "rated-p": "5", "rated-q": "1000", "shutoff-p": "7", "peak-p": "3.25", "peak-q": "1500"})
    tree.write(sdf)
    output = tmp_path / "pump.kfp"
    emit_physical_kfp(sdf, output)
    data = json.loads(output.read_text(encoding="utf-8"))
    assert len(data["pipe_data"]) == 3
    assert sum(p["length_m"] for p in data["pipe_data"].values()) == pytest.approx(5.8)
    assert data["nodes_meta_runtime"]["N9"]["type_id"] == "pump"
    assert data["nodes_meta_runtime"]["N9"]["rated_p"] == 5
    assert data["pipe_data"]["P1"]["start"] == "N9"
    assert "N1" not in data["nodes_meta_runtime"]


def test_disconnected_network_is_not_overlaid_or_silently_joined():
    with pytest.raises(ValueError, match="끊겨"):
        physical_layout({"a": (0, 0, 0), "b": (1, 0, 0), "c": (9, 0, 0)},
                        [LayoutEdge("P1", "a", "b", 1)])


def test_merged_iso_kfp_is_same_physical_network(tmp_path):
    from remote30_full_network import CombinedTables
    from routes.module_f.emit import emit_merged
    # All positions are unprojected mm; make the iso view deliberately different.
    combined = CombinedTables(nodes=[
        {"label": "1", "x": 0, "y": 0, "elevation": 1, "io_node": "Input"},
        {"label": "2", "x": 5000, "y": 0, "elevation": 1, "io_node": "No"},
        {"label": "3", "x": 5000, "y": 0, "elevation": .7, "io_node": "No"},
    ], pipes=[
        {"label": "P1", "in": "1", "out": "2", "length": 5, "elev": 0, "dia": 25, "c": 120},
        {"label": "P2", "in": "2", "out": "3", "length": .3, "elev": -.3, "dia": 25, "c": 120},
    ], nozzles=[{"label": "1", "in": "3", "out": "@/1", "status": "1", "lib": "SP-HEAD", "flow_m3s": 80/60000}],
       fittings=[], equipment=[], pumps=[], valves=[], meta=[])
    iso = [dict(n, x=n["x"]*.866, y=n["x"]*.5+n["elevation"]*100) for n in combined.nodes]
    result = emit_merged(combined, tmp_path, iso_nodes=iso)
    assert result.get("kfp_iso"), result["warnings"]
    assert Path(result["kfp"]).read_bytes() == Path(result["kfp_iso"]).read_bytes()
