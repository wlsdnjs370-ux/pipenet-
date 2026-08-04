"""Tests for PipeNet SDF writing."""

from pathlib import Path

from pipenet_converter.models import Equipment, Fitting, Node, Nozzle, Pipe, PipeNetwork, Valve
from pipenet_converter.sdf_parser import parse_sdf, parse_sdf_text
from pipenet_converter.sdf_writer import network_to_sdf_text, write_sdf


def _sample_network() -> PipeNetwork:
    network = PipeNetwork(title="Writer sample")
    network.add_node(Node("10", 6500.0, 692.0, 48.55, "Input"))
    network.add_node(Node("20", 6600.0, 700.0, 48.55, "No"))
    network.add_node(Node("30", 6700.0, 710.0, 48.85, "No"))
    network.add_node(Node("@/1", 6750.0, 730.0, 49.15, "No"))
    network.add_pipe(
        Pipe(
            pipe_id="88",
            from_node="10",
            to_node="20",
            diameter_m=0.15,
            length_m=68.2,
            rise_m=0.0,
            fittings=[Fitting("elbow", 3), Fitting("tee", 1)],
            equipment=[Equipment("EQ1", "AV", 24.0)],
            waypoints=[(6550.0, 696.0)],
        )
    )
    network.add_pipe(
        Pipe(
            pipe_id="89",
            from_node="20",
            to_node="30",
            diameter_m=0.05,
            length_m=3.5,
            rise_m=0.3,
        )
    )
    network.add_pipe(
        Pipe(
            pipe_id="90",
            from_node="30",
            to_node="@/1",
            diameter_m=0.032,
            length_m=0.3,
            rise_m=0.3,
        )
    )
    network.add_nozzle(Nozzle("1", "30", "@/1", 0.00266666667))
    network.add_valve(Valve("V1", "20", "30", "pressure-drop", target_value=0.0))
    return network


def test_network_to_sdf_text_round_trips_through_parser() -> None:
    xml_text = network_to_sdf_text(_sample_network())

    parsed = parse_sdf_text(xml_text)

    assert parsed.title == "Writer sample"
    assert len(parsed.nodes) == 4
    assert len(parsed.pipes) == 3
    assert len(parsed.nozzles) == 1
    assert len(parsed.valves) == 1

    pipe = parsed.pipes["88"]
    assert [(fitting.fitting_type, fitting.count) for fitting in pipe.fittings] == [
        ("elbow", 3),
        ("tee", 1),
    ]
    assert pipe.equipment[0].equipment_id == "EQ1"
    assert pipe.equipment[0].description == "AV"
    assert pipe.equipment[0].equivalent_length_m == 24.0
    assert pipe.waypoints == [(6550.0, 696.0)]


def test_nozzle_flow_survives_the_lmin_round_trip() -> None:
    """설계 유량 80 L/min 이 리포트에 80.0000 으로 찍혀야 한다(79.9998 아님)."""
    network = PipeNetwork(title="flow precision")
    network.add_node(Node("30", 0.0, 0.0, 0.0, "No"))
    network.add_node(Node("@/1", 1.0, 0.0, 0.0, "No"))
    for design_lmin in (80.0, 130.0):
        network.nozzles.clear()
        network.add_nozzle(Nozzle("1", "30", "@/1", design_lmin / 60000.0))

        parsed = parse_sdf_text(network_to_sdf_text(network))

        assert round(parsed.nozzles["1"].flow_m3s * 60000.0, 4) == design_lmin


def test_write_sdf_writes_file(tmp_path: Path) -> None:
    output_path = tmp_path / "generated.sdf"

    write_sdf(_sample_network(), output_path)

    parsed = parse_sdf(output_path)
    assert len(parsed.nodes) == 4
    assert len(parsed.pipes) == 3
    assert len(parsed.nozzles) == 1


def test_write_sdf_template_mode_replaces_nodes_and_links(tmp_path: Path) -> None:
    template_path = tmp_path / "template.sdf"
    output_path = tmp_path / "generated_from_template.sdf"
    template_path.write_text(
        """<?xml version="1.0" encoding="UTF-8"?>
<Project version="1.6  (0)">
  <Network-spray>
    <Title>Old title</Title>
    <Attributes keep="yes"/>
    <Nodes>
      <Node elevation="0" io-node="No" label="OLD">
        <Position x="0" y="0"/>
      </Node>
    </Nodes>
    <Links/>
  </Network-spray>
  <Graphics keep="yes"/>
</Project>
""",
        encoding="utf-8",
    )

    write_sdf(_sample_network(), output_path, template_path=template_path)

    output_text = output_path.read_text(encoding="utf-8")
    parsed = parse_sdf_text(output_text)
    assert parsed.title == "Writer sample"
    assert "OLD" not in parsed.nodes
    assert len(parsed.nodes) == 4
    assert "<Attributes keep=\"yes\" />" in output_text or "<Attributes keep=\"yes\"/>" in output_text
    assert "<Graphics keep=\"yes\" />" in output_text or "<Graphics keep=\"yes\"/>" in output_text
