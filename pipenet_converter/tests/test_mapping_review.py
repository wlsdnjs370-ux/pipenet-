"""Tests for real-project mapping review helpers."""

from pathlib import Path

import ezdxf
import pandas as pd

from pipenet_converter.mapping_review import (
    DAISO_TARGET_DXF_NAME,
    prepare_daiso_4f_mapping_review,
)


def test_prepare_daiso_mapping_review_handles_missing_inputs(tmp_path: Path) -> None:
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "outputs"

    report_path = prepare_daiso_4f_mapping_review(input_dir, output_dir)

    assert report_path.exists()
    report = report_path.read_text(encoding="utf-8")
    assert "status: missing" in report
    assert (output_dir / "layer_map_daiso_4f.csv").exists()
    assert (output_dir / "block_map_daiso_4f.csv").exists()
    assert (output_dir / "extraction_summary.csv").exists()


def test_prepare_daiso_mapping_review_drafts_maps_from_dxf(tmp_path: Path) -> None:
    input_dir = tmp_path / "input"
    output_dir = tmp_path / "outputs"
    input_dir.mkdir()
    _write_review_dxf(input_dir / DAISO_TARGET_DXF_NAME)

    prepare_daiso_4f_mapping_review(input_dir, output_dir)

    layer_map = pd.read_csv(output_dir / "layer_map_daiso_4f.csv")
    block_map = pd.read_csv(output_dir / "block_map_daiso_4f.csv")
    report = (output_dir / "daiso_4f_mapping_review.md").read_text(encoding="utf-8")

    assert set(layer_map["layer_name"]) == {"F-SP-PIPE", "F-SP-TEXT", "A-WALL"}
    assert layer_map.loc[layer_map["layer_name"] == "F-SP-PIPE", "class"].iloc[0] == "pipe"
    assert layer_map.loc[layer_map["layer_name"] == "F-SP-TEXT", "class"].iloc[0] == "diameter_text"
    assert block_map.loc[block_map["block_name"] == "SP_HEAD", "class"].iloc[0] == "head"
    assert "Likely Pipe Layers" in report
    assert "`F-SP-PIPE`" in report


def _write_review_dxf(path: Path) -> None:
    doc = ezdxf.new("R2010")
    for layer in ["F-SP-PIPE", "F-SP-TEXT", "A-WALL"]:
        doc.layers.add(layer)
    doc.blocks.new("SP_HEAD").add_circle((0, 0), radius=50)
    modelspace = doc.modelspace()
    modelspace.add_line((0, 0), (1000, 0), dxfattribs={"layer": "F-SP-PIPE"})
    modelspace.add_text("150A", dxfattribs={"layer": "F-SP-TEXT"}).set_placement((500, 0))
    modelspace.add_line((0, 100), (1000, 100), dxfattribs={"layer": "A-WALL"})
    modelspace.add_blockref("SP_HEAD", (1000, 0), dxfattribs={"layer": "F-SP-PIPE"})
    doc.saveas(path)
