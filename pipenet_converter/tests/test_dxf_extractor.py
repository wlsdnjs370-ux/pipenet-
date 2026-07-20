"""Tests for DXF raw entity extraction."""

from pathlib import Path

import ezdxf
import pandas as pd

from pipenet_converter.dxf_extractor import (
    extract_dxf,
    is_diameter_text,
    load_block_map,
    load_layer_map,
    normalize_diameter_text,
    validate_dxf_extraction,
    write_dxf_extraction_tables,
)


def test_extract_dxf_reads_pipe_segments_blocks_and_diameter_text(tmp_path: Path) -> None:
    dxf_path = tmp_path / "sample.dxf"
    layer_map_path = tmp_path / "layer_map.csv"
    block_map_path = tmp_path / "block_map.csv"
    layer_map_path.write_text(
        "layer_name,class\n"
        "F-SP-PIPE,pipe\n"
        "F-SP-HEAD,head\n"
        "F-SP-TEXT,diameter_text\n"
        "F-SP-VALVE,valve\n"
        "F-SP-RISER,riser\n",
        encoding="utf-8",
    )
    block_map_path.write_text(
        "block_name,class\n"
        "SP_HEAD,head\n"
        "GATE_VALVE,gate\n"
        "CHECK_VALVE,check\n"
        "ALARM_VALVE,alarm_valve\n"
        "RISER,riser\n",
        encoding="utf-8",
    )

    doc = ezdxf.new("R2010")
    for layer in ["F-SP-PIPE", "F-SP-HEAD", "F-SP-TEXT"]:
        doc.layers.add(layer)

    head_block = doc.blocks.new("SP_HEAD")
    head_block.add_circle((0, 0), radius=50)

    modelspace = doc.modelspace()
    modelspace.add_line((0, 0), (1000, 0), dxfattribs={"layer": "F-SP-PIPE"})
    modelspace.add_lwpolyline(
        [(1000, 0), (1000, 500), (1500, 500)],
        dxfattribs={"layer": "F-SP-PIPE"},
    )
    modelspace.add_blockref("SP_HEAD", (1500, 500), dxfattribs={"layer": "F-SP-HEAD", "rotation": 45})
    modelspace.add_text("150A", dxfattribs={"layer": "F-SP-TEXT"}).set_placement((1200, 250))
    doc.saveas(dxf_path)

    result = extract_dxf(dxf_path, layer_map_path, block_map_path)

    assert len(result.segments) == 3
    assert result.segments[0].start == (0.0, 0.0)
    assert result.segments[0].end == (1000.0, 0.0)
    assert all(segment.semantic_class == "pipe" for segment in result.segments)

    assert len(result.blocks) == 1
    assert result.blocks[0].block_name == "SP_HEAD"
    assert result.blocks[0].insert == (1500.0, 500.0)
    assert result.blocks[0].semantic_class == "head"
    assert result.blocks[0].rotation == 45.0

    assert len(result.texts) == 1
    assert result.texts[0].text == "150A"
    assert result.texts[0].insert == (1200.0, 250.0)
    assert result.texts[0].semantic_class == "diameter_text"


def test_mapping_loaders_and_diameter_text_helpers(tmp_path: Path) -> None:
    layer_map_path = tmp_path / "layer_map.csv"
    block_map_path = tmp_path / "block_map.csv"
    layer_map_path.write_text("layer_name,class\nF-SP-PIPE,pipe\n", encoding="utf-8")
    block_map_path.write_text("block_name,class\nSP_HEAD,head\n", encoding="utf-8")

    assert load_layer_map(layer_map_path) == {"F-SP-PIPE": "pipe"}
    assert load_block_map(block_map_path) == {"SP_HEAD": "head"}
    assert is_diameter_text("150 A")
    assert normalize_diameter_text("pipe 65a") == "65A"
    assert normalize_diameter_text("DN300") is None


def test_missing_csv_column_raises_clear_error(tmp_path: Path) -> None:
    layer_map_path = tmp_path / "bad_layer_map.csv"
    layer_map_path.write_text("layer_name\nF-SP-PIPE\n", encoding="utf-8")

    import pytest

    with pytest.raises(ValueError, match="missing required columns: class"):
        load_layer_map(layer_map_path)


def test_write_dxf_extraction_tables_creates_human_review_csvs(tmp_path: Path) -> None:
    dxf_path, layer_map_path, block_map_path = _write_basic_dxf_fixture(tmp_path)
    result = extract_dxf(dxf_path, layer_map_path, block_map_path)
    output_dir = tmp_path / "extracted"

    write_dxf_extraction_tables(result, output_dir)

    expected_files = {
        "raw_segments.csv",
        "raw_blocks.csv",
        "raw_texts.csv",
        "extraction_summary.csv",
        "extraction_warnings.csv",
    }
    assert {path.name for path in output_dir.iterdir()} == expected_files

    segments = pd.read_csv(output_dir / "raw_segments.csv")
    blocks = pd.read_csv(output_dir / "raw_blocks.csv")
    texts = pd.read_csv(output_dir / "raw_texts.csv")
    summary = pd.read_csv(output_dir / "extraction_summary.csv")

    assert list(segments.columns) == [
        "segment_id",
        "start_x",
        "start_y",
        "end_x",
        "end_y",
        "layer",
        "semantic_class",
        "length",
    ]
    assert list(blocks.columns) == ["block_id", "block_name", "x", "y", "layer", "semantic_class", "rotation"]
    assert list(texts.columns) == ["text_id", "text", "x", "y", "layer", "semantic_class", "normalized_diameter"]
    assert set(summary["item"]) == {"segments", "blocks", "texts", "unknown_layers", "unknown_blocks"}


def test_unknown_layer_warning_works(tmp_path: Path) -> None:
    dxf_path, layer_map_path, block_map_path = _write_basic_dxf_fixture(tmp_path, include_unknown_layer=True)

    result = extract_dxf(dxf_path, layer_map_path, block_map_path)
    issues = validate_dxf_extraction(result)

    assert "UNKNOWN-LAYER" in result.unknown_layers
    assert any(issue.code == "UNKNOWN_LAYER" and issue.object_id == "UNKNOWN-LAYER" for issue in issues)


def test_empty_extraction_warns() -> None:
    from pipenet_converter.dxf_extractor import DxfExtractionResult

    issues = validate_dxf_extraction(DxfExtractionResult())

    assert {"NO_PIPE_SEGMENTS", "NO_HEAD_BLOCKS", "NO_DIAMETER_TEXT"}.issubset(
        {issue.code for issue in issues}
    )


def test_unknown_block_warning_works(tmp_path: Path) -> None:
    dxf_path, layer_map_path, block_map_path = _write_basic_dxf_fixture(tmp_path, include_unknown_block=True)

    result = extract_dxf(dxf_path, layer_map_path, block_map_path)
    issues = validate_dxf_extraction(result)

    assert "UNKNOWN_HEAD" in result.unknown_blocks
    assert any(issue.code == "UNKNOWN_BLOCK" and issue.object_id == "UNKNOWN_HEAD" for issue in issues)


def _write_basic_dxf_fixture(
    tmp_path: Path,
    include_unknown_layer: bool = False,
    include_unknown_block: bool = False,
) -> tuple[Path, Path, Path]:
    dxf_path = tmp_path / "basic.dxf"
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
    if include_unknown_layer:
        doc.layers.add("UNKNOWN-LAYER")

    doc.blocks.new("SP_HEAD").add_circle((0, 0), radius=50)
    if include_unknown_block:
        doc.blocks.new("UNKNOWN_HEAD").add_circle((0, 0), radius=50)

    modelspace = doc.modelspace()
    modelspace.add_line((0, 0), (1000, 0), dxfattribs={"layer": "F-SP-PIPE"})
    modelspace.add_blockref("SP_HEAD", (1000, 0), dxfattribs={"layer": "F-SP-HEAD"})
    modelspace.add_text("150A", dxfattribs={"layer": "F-SP-TEXT"}).set_placement((500, 0))
    if include_unknown_layer:
        modelspace.add_line((0, 100), (1000, 100), dxfattribs={"layer": "UNKNOWN-LAYER"})
    if include_unknown_block:
        modelspace.add_blockref("UNKNOWN_HEAD", (1500, 0), dxfattribs={"layer": "F-SP-HEAD"})
    doc.saveas(dxf_path)
    return dxf_path, layer_map_path, block_map_path
