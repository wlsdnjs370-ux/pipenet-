"""Tests for PipeNet SDF parsing."""

from pathlib import Path

from pipenet_converter.sdf_parser import parse_sdf, parse_sdf_text


SAMPLE_SDF = """<?xml version="1.0" encoding="UTF-8"?>
<Project version="1.6  (0)">
  <Network-spray>
    <Title>Sample sprinkler network</Title>
    <Nodes>
      <Node elevation="48.55" io-node="Input" label="10">
        <Position x="6500" y="692"/>
      </Node>
      <Node elevation="49.15" io-node="No" label="@/4">
        <Position x="5200" y="1182"/>
      </Node>
    </Nodes>
    <Links>
      <Pipe-set>
        <Pipe bore="0.15"
              input="10"
              label="88"
              length="68.2"
              output="128"
              rise="0"
              roughness-or-c="120"
              status="normal">
          <Fittings>
            <Fitting count="3" type="elbow"/>
            <Fitting count="1" type="tee"/>
          </Fittings>
          <Components>
            <Equipment description="AV"
                       equivalent-length="24"
                       label="2"
                       rel-position="0.297400758"/>
          </Components>
          <Waypoints symbol-segment="0">
            <Position x="6425" y="736"/>
            <Position x="6600" y="837"/>
          </Waypoints>
        </Pipe>
      </Pipe-set>
      <Nozzle input="128" label="4" output="@/4" status="1">
        <Flow-define flow="0.00266666667"/>
        <Library-item>SP-HEAD</Library-item>
      </Nozzle>
      <Elastomeric-valve input="13"
                         label="1"
                         output="15"
                         target-value="0"
                         type="pressure-drop"/>
    </Links>
  </Network-spray>
</Project>
"""


def test_parse_sdf_text_parses_title_nodes_pipes_and_links() -> None:
    network = parse_sdf_text(SAMPLE_SDF)

    assert network.title == "Sample sprinkler network"

    assert set(network.nodes) == {"10", "@/4"}
    input_node = network.nodes["10"]
    assert input_node.x == 6500.0
    assert input_node.y == 692.0
    assert input_node.z == 48.55
    assert input_node.node_type == "Input"
    assert input_node.metadata["io_node"] == "Input"

    pipe = network.pipes["88"]
    assert pipe.from_node == "10"
    assert pipe.to_node == "128"
    assert pipe.diameter_m == 0.15
    assert pipe.length_m == 68.2
    assert pipe.rise_m == 0.0
    assert pipe.c_factor == 120.0
    assert pipe.status == "normal"
    assert pipe.waypoints == [(6425.0, 736.0), (6600.0, 837.0)]

    assert [(fitting.fitting_type, fitting.count) for fitting in pipe.fittings] == [
        ("elbow", 3),
        ("tee", 1),
    ]

    equipment = pipe.equipment[0]
    assert equipment.equipment_id == "2"
    assert equipment.description == "AV"
    assert equipment.equivalent_length_m == 24.0
    assert equipment.rel_position == 0.297400758

    nozzle = network.nozzles["4"]
    assert nozzle.input_node == "128"
    assert nozzle.output_node == "@/4"
    assert nozzle.status == 1
    assert nozzle.flow_m3s == 0.00266666667
    assert nozzle.library_item == "SP-HEAD"

    valve = network.valves["1"]
    assert valve.input_node == "13"
    assert valve.output_node == "15"
    assert valve.valve_type == "pressure-drop"
    assert valve.target_value == 0.0
    assert valve.metadata["sdf_tag"] == "Elastomeric-valve"


def test_parse_sdf_reads_file(tmp_path: Path) -> None:
    sdf_path = tmp_path / "sample.sdf"
    sdf_path.write_text(SAMPLE_SDF, encoding="utf-8")

    network = parse_sdf(sdf_path)

    assert network.title == "Sample sprinkler network"
    assert "88" in network.pipes
