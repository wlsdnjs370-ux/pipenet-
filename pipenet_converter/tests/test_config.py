"""Tests for YAML pipeline configuration loading."""

from pathlib import Path

import pytest

from pipenet_converter.config import load_pipeline_config


def test_load_pipeline_config_resolves_relative_paths(tmp_path: Path) -> None:
    config_dir = tmp_path / "configs"
    config_dir.mkdir()
    config_path = config_dir / "run.yaml"
    config_path.write_text(
        """
project_name: test_project
plan_dxf: ../data/plan.dxf
layer_map: ../data/layer_map.csv
block_map: ../data/block_map.csv
elevation_rules: ../data/elevation_rules.csv
system_edges: ../data/system_edges.csv
connection_map: ../data/connection_map.json
template_sdf: ../data/template.sdf
reference_sdf: ../data/reference.sdf
output_dir: ../outputs/run
snap_tolerance: 25
head_attach_tolerance: 500
cad_unit_scale_to_m: 0.001
default_diameter_m: 0.05
input_node_id: INPUT
""".strip(),
        encoding="utf-8",
    )

    config = load_pipeline_config(config_path)

    assert config.project_name == "test_project"
    assert config.plan_dxf == str(tmp_path / "data" / "plan.dxf")
    assert config.output_dir == str(tmp_path / "outputs" / "run")
    assert config.snap_tolerance == 25.0
    assert config.head_attach_tolerance == 500.0
    assert config.default_diameter_m == 0.05


def test_load_pipeline_config_reports_missing_required_fields(tmp_path: Path) -> None:
    config_path = tmp_path / "bad.yaml"
    config_path.write_text("project_name: bad\n", encoding="utf-8")

    with pytest.raises(ValueError, match="missing required fields"):
        load_pipeline_config(config_path)
