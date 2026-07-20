"""Tests for the command-line interface."""

from pathlib import Path

import ezdxf
import pytest

from pipenet_converter.cli import main
from pipenet_converter.export_tables import write_network_tables
from pipenet_converter.models import Fitting, Node, Nozzle, Pipe, PipeNetwork
from pipenet_converter.sdf_parser import parse_sdf
from pipenet_converter.sdf_writer import write_sdf


SAMPLE_SDF = """<?xml version="1.0" encoding="UTF-8"?>
<Project version="1.6  (0)">
  <Network-spray>
    <Title>CLI sample</Title>
    <Nodes>
      <Node elevation="0" io-node="Input" label="INPUT"><Position x="0" y="0"/></Node>
      <Node elevation="0" io-node="No" label="N001"><Position x="1" y="0"/></Node>
      <Node elevation="0" io-node="No" label="@/1"><Position x="1" y="0"/></Node>
    </Nodes>
    <Links>
      <Pipe-set>
        <Pipe bore="0.15" input="INPUT" label="P001" length="1" output="N001" rise="0" roughness-or-c="120" status="normal">
          <Fittings><Fitting count="1" type="gate"/></Fittings>
        </Pipe>
      </Pipe-set>
      <Nozzle input="N001" label="NZ1" output="@/1" status="1">
        <Flow-define flow="0.00266666667"/>
        <Library-item>SP-HEAD</Library-item>
      </Nozzle>
    </Links>
  </Network-spray>
  <Graphics/>
</Project>
"""


def test_cli_help_works(capsys: pytest.CaptureFixture[str]) -> None:
    with pytest.raises(SystemExit) as exc_info:
        main(["--help"])

    assert exc_info.value.code == 0
    assert "parse-sdf" in capsys.readouterr().out


def test_parse_sdf_command_creates_csv_outputs(tmp_path: Path) -> None:
    sdf_path = tmp_path / "input.sdf"
    out_dir = tmp_path / "parsed"
    sdf_path.write_text(SAMPLE_SDF, encoding="utf-8")

    main(["parse-sdf", "--sdf", str(sdf_path), "--out-dir", str(out_dir)])

    assert (out_dir / "network_3d_nodes.csv").exists()
    assert (out_dir / "network_3d_pipes.csv").exists()
    assert (out_dir / "network_3d_nozzles.csv").exists()
    assert (out_dir / "validation_report.csv").exists()


def test_extract_dxf_command_creates_raw_csv_outputs(tmp_path: Path) -> None:
    dxf_path, layer_map, block_map = _write_dxf_fixture(tmp_path)
    out_dir = tmp_path / "extracted"

    main(
        [
            "extract-dxf",
            "--dxf",
            str(dxf_path),
            "--layer-map",
            str(layer_map),
            "--block-map",
            str(block_map),
            "--out-dir",
            str(out_dir),
        ]
    )

    assert (out_dir / "raw_segments.csv").exists()
    assert (out_dir / "raw_blocks.csv").exists()
    assert (out_dir / "raw_texts.csv").exists()


def test_write_sdf_command_creates_output_file(tmp_path: Path) -> None:
    network_dir = tmp_path / "network"
    output_sdf = tmp_path / "generated.sdf"
    write_network_tables(_network_fixture(), network_dir)

    main(["write-sdf", "--network-dir", str(network_dir), "--output", str(output_sdf)])

    assert output_sdf.exists()
    assert output_sdf.stat().st_size > 0
    assert output_sdf.with_suffix(".validation_report.csv").exists()


def test_build_network_command_runs_on_generated_dxf_fixture(tmp_path: Path) -> None:
    dxf_path, layer_map, block_map = _write_dxf_fixture(tmp_path)
    elevation_rules = tmp_path / "elevation_rules.csv"
    elevation_rules.write_text(
        "rule_id,priority,match_field,match_value,z_m,description\n"
        "Z_HEAD,10,node_type,head,49.15,head\n"
        "Z_DEFAULT,999,default,default,48.55,default\n",
        encoding="utf-8",
    )
    out_dir = tmp_path / "network"

    main(
        [
            "build-network",
            "--dxf",
            str(dxf_path),
            "--layer-map",
            str(layer_map),
            "--block-map",
            str(block_map),
            "--elevation-rules",
            str(elevation_rules),
            "--out-dir",
            str(out_dir),
        ]
    )

    assert (out_dir / "network_3d_nodes.csv").exists()
    assert (out_dir / "network_3d_pipes.csv").exists()
    assert (out_dir / "isometric_check.png").exists()


def test_render_iso_and_validate_commands_run(tmp_path: Path) -> None:
    network_dir = tmp_path / "network"
    iso_path = tmp_path / "iso.png"
    report_path = tmp_path / "validation.csv"
    write_network_tables(_network_fixture(), network_dir)

    main(["render-iso", "--network-dir", str(network_dir), "--output", str(iso_path)])
    main(
        [
            "validate",
            "--network-dir",
            str(network_dir),
            "--input-node",
            "INPUT",
            "--out",
            str(report_path),
        ]
    )

    assert iso_path.exists()
    assert report_path.exists()


def test_strict_mode_fails_on_validation_errors(tmp_path: Path) -> None:
    network_dir = tmp_path / "bad_network"
    report_path = tmp_path / "validation.csv"
    network = PipeNetwork(title="Bad network")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input"))
    network.add_pipe(Pipe("P_BAD", "INPUT", "MISSING", 0.15, 1.0, 0.0))
    write_network_tables(network, network_dir)

    with pytest.raises(SystemExit) as exc_info:
        main(
            [
                "validate",
                "--strict",
                "--network-dir",
                str(network_dir),
                "--input-node",
                "INPUT",
                "--out",
                str(report_path),
            ]
        )

    assert exc_info.value.code == 1
    assert report_path.exists()


def test_merge_system_command_runs(tmp_path: Path) -> None:
    network_dir = tmp_path / "network"
    out_dir = tmp_path / "merged"
    system_edges = tmp_path / "system_edges.csv"
    connection_map = tmp_path / "connection_map.json"
    write_network_tables(_network_fixture(), network_dir)
    system_edges.write_text(
        "edge_id,from_node,to_node,from_z_m,to_z_m,diameter_m,length_m,rise_m,c_factor,material,fittings,equipment,description\n"
        "R001,INPUT,SYS_MAIN,0,0,0.15,1,0,120,KSD3507,\"\",\"\",system\n",
        encoding="utf-8",
    )
    connection_map.write_text('{"SYS_MAIN": "INPUT"}', encoding="utf-8")

    main(
        [
            "merge-system",
            "--network-dir",
            str(network_dir),
            "--system-edges",
            str(system_edges),
            "--connection-map",
            str(connection_map),
            "--out-dir",
            str(out_dir),
        ]
    )

    assert (out_dir / "network_3d_nodes.csv").exists()
    assert (out_dir / "validation_report.csv").exists()


def test_compare_sdf_command_creates_report(tmp_path: Path) -> None:
    reference_sdf = tmp_path / "reference.sdf"
    candidate_sdf = tmp_path / "candidate.sdf"
    report_csv = tmp_path / "compare_report.csv"
    reference = _network_fixture()
    candidate = _network_fixture()
    candidate.pipes["P001"].length_m = 2.0
    write_sdf(reference, reference_sdf)
    write_sdf(candidate, candidate_sdf)

    main(
        [
            "compare-sdf",
            "--reference",
            str(reference_sdf),
            "--candidate",
            str(candidate_sdf),
            "--out",
            str(report_csv),
        ]
    )

    assert report_csv.exists()
    assert "PIPE_LENGTH_MISMATCH" in report_csv.read_text(encoding="utf-8")


def test_run_pipeline_completes_and_writes_expected_outputs(tmp_path: Path) -> None:
    dxf_path, layer_map, block_map = _write_dxf_fixture(tmp_path)
    elevation_rules = tmp_path / "elevation_rules.csv"
    system_edges = tmp_path / "system_edges.csv"
    connection_map = tmp_path / "connection_map.json"
    template_sdf = tmp_path / "template.sdf"
    reference_sdf = tmp_path / "reference.sdf"
    out_dir = tmp_path / "run_001"

    elevation_rules.write_text(
        "rule_id,priority,match_field,match_value,z_m,description\n"
        "Z_HEAD,10,node_type,head,49.15,head\n"
        "Z_HEAD_OUT,20,node_type,head_output,49.15,head output\n"
        "Z_DEFAULT,999,default,default,48.55,default\n",
        encoding="utf-8",
    )
    system_edges.write_text(
        "edge_id,from_node,to_node,from_z_m,to_z_m,diameter_m,length_m,rise_m,c_factor,material,fittings,equipment,description\n"
        "R001,INPUT,SYS_MAIN,0,48.55,0.15,48.55,48.55,120,KSD3507,\"gate:1\",\"AV:24\",system\n",
        encoding="utf-8",
    )
    connection_map.write_text('{"SYS_MAIN": "N000001"}', encoding="utf-8")
    template_sdf.write_text(SAMPLE_SDF, encoding="utf-8")
    write_sdf(_network_fixture(), reference_sdf)

    main(
        [
            "run-pipeline",
            "--plan-dxf",
            str(dxf_path),
            "--layer-map",
            str(layer_map),
            "--block-map",
            str(block_map),
            "--elevation-rules",
            str(elevation_rules),
            "--system-edges",
            str(system_edges),
            "--connection-map",
            str(connection_map),
            "--template-sdf",
            str(template_sdf),
            "--reference-sdf",
            str(reference_sdf),
            "--out-dir",
            str(out_dir),
        ]
    )

    expected_outputs = [
        "raw_segments.csv",
        "raw_blocks.csv",
        "raw_texts.csv",
        "extraction_summary.csv",
        "extraction_warnings.csv",
        "network_3d_nodes.csv",
        "network_3d_pipes.csv",
        "network_3d_nozzles.csv",
        "network_3d_fittings.csv",
        "network_3d_equipment.csv",
        "network_3d_valves.csv",
        "validation_report.csv",
        "isometric_check.png",
        "generated_pipenet.sdf",
        "compare_report.csv",
        "redline_items.csv",
        "redline_items.json",
        "isometric_redline.png",
        "run_summary.txt",
    ]
    for filename in expected_outputs:
        assert (out_dir / filename).exists(), filename

    parsed = parse_sdf(out_dir / "generated_pipenet.sdf")
    assert len(parsed.nodes) > 0
    assert len(parsed.pipes) > 0
    assert len(parsed.nozzles) > 0
    summary = (out_dir / "run_summary.txt").read_text(encoding="utf-8")
    assert "generated_sdf_path:" in summary
    assert "isometric_png_path:" in summary


def test_run_config_command_runs_pipeline_from_yaml(tmp_path: Path) -> None:
    dxf_path, layer_map, block_map = _write_dxf_fixture(tmp_path)
    elevation_rules = tmp_path / "elevation_rules.csv"
    system_edges = tmp_path / "system_edges.csv"
    connection_map = tmp_path / "connection_map.json"
    template_sdf = tmp_path / "template.sdf"
    reference_sdf = tmp_path / "reference.sdf"
    out_dir = tmp_path / "config_run"
    config_path = tmp_path / "pipeline.yaml"

    elevation_rules.write_text(
        "rule_id,priority,match_field,match_value,z_m,description\n"
        "Z_HEAD,10,node_type,head,49.15,head\n"
        "Z_HEAD_OUT,20,node_type,head_output,49.15,head output\n"
        "Z_DEFAULT,999,default,default,48.55,default\n",
        encoding="utf-8",
    )
    system_edges.write_text(
        "edge_id,from_node,to_node,from_z_m,to_z_m,diameter_m,length_m,rise_m,c_factor,material,fittings,equipment,description\n"
        "R001,INPUT,SYS_MAIN,0,48.55,0.15,48.55,48.55,120,KSD3507,\"gate:1\",\"AV:24\",system\n",
        encoding="utf-8",
    )
    connection_map.write_text('{"SYS_MAIN": "N000001"}', encoding="utf-8")
    template_sdf.write_text(SAMPLE_SDF, encoding="utf-8")
    write_sdf(_network_fixture(), reference_sdf)
    config_path.write_text(
        f"""
project_name: cli_config_test
plan_dxf: {dxf_path.name}
layer_map: {layer_map.name}
block_map: {block_map.name}
elevation_rules: {elevation_rules.name}
system_edges: {system_edges.name}
connection_map: {connection_map.name}
template_sdf: {template_sdf.name}
reference_sdf: {reference_sdf.name}
output_dir: {out_dir.name}
snap_tolerance: 50
head_attach_tolerance: 600
cad_unit_scale_to_m: 0.001
default_diameter_m: 0.032
input_node_id: INPUT
""".strip(),
        encoding="utf-8",
    )

    main(["run-config", "--config", str(config_path)])

    assert (out_dir / "generated_pipenet.sdf").exists()
    assert (out_dir / "run_summary.txt").exists()
    assert (out_dir / "isometric_check.png").exists()


def _network_fixture() -> PipeNetwork:
    network = PipeNetwork(title="CLI network")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input", source="test"))
    network.add_node(Node("N001", 1.0, 0.0, 0.0, "main_pipe", source="test"))
    network.add_node(Node("@/1", 1.0, 0.0, 0.0, "head_output", source="test"))
    network.add_pipe(
        Pipe(
            "P001",
            "INPUT",
            "N001",
            0.15,
            1.0,
            0.0,
            fittings=[Fitting("gate", 1)],
        )
    )
    network.add_nozzle(Nozzle("NZ1", "N001", "@/1", 0.00266666667, status=1))
    return network


def _write_dxf_fixture(tmp_path: Path) -> tuple[Path, Path, Path]:
    dxf_path = tmp_path / "plan.dxf"
    layer_map_path = tmp_path / "layer_map.csv"
    block_map_path = tmp_path / "block_map.csv"
    layer_map_path.write_text(
        "layer_name,class\nF-SP-PIPE,pipe\nF-SP-HEAD,head\nF-SP-TEXT,diameter_text\n",
        encoding="utf-8",
    )
    block_map_path.write_text("block_name,class\nSP_HEAD,head\n", encoding="utf-8")

    doc = ezdxf.new("R2010")
    for layer in ["F-SP-PIPE", "F-SP-HEAD", "F-SP-TEXT"]:
        doc.layers.add(layer)
    head_block = doc.blocks.new("SP_HEAD")
    head_block.add_circle((0, 0), radius=50)
    modelspace = doc.modelspace()
    modelspace.add_line((0, 0), (1000, 0), dxfattribs={"layer": "F-SP-PIPE"})
    modelspace.add_blockref("SP_HEAD", (1000, 0), dxfattribs={"layer": "F-SP-HEAD"})
    modelspace.add_text("150A", dxfattribs={"layer": "F-SP-TEXT"}).set_placement((500, 0))
    doc.saveas(dxf_path)
    return dxf_path, layer_map_path, block_map_path
