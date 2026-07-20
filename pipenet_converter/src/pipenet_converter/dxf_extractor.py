"""DXF extraction utilities.

This module extracts raw 2D CAD entities needed for sprinkler network
generation. It intentionally supports DXF only and does not depend on AutoCAD,
DWG parsing, COM automation, or OCR.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass, field
from math import dist
from pathlib import Path
import re
from typing import Any

import ezdxf
import pandas as pd

from pipenet_converter.models import ValidationIssue


SUPPORTED_DIAMETER_LABELS = {"25A", "32A", "40A", "50A", "65A", "80A", "100A", "125A", "150A", "200A"}
DIAMETER_PATTERN = re.compile(r"\b(25|32|40|50|65|80|100|125|150|200)\s*A\b", re.IGNORECASE)
PIPE_CLASSES = {"pipe", "branch_pipe", "main_pipe", "riser"}


@dataclass(slots=True)
class RawSegment:
    """A raw 2D line segment extracted from pipe-like CAD geometry."""

    segment_id: str
    start: tuple[float, float]
    end: tuple[float, float]
    layer: str
    semantic_class: str | None


@dataclass(slots=True)
class RawBlock:
    """A raw block insert extracted from a DXF file."""

    block_id: str
    block_name: str
    insert: tuple[float, float]
    layer: str
    semantic_class: str | None
    rotation: float | None


@dataclass(slots=True)
class RawText:
    """A raw diameter text annotation extracted from a DXF file."""

    text_id: str
    text: str
    insert: tuple[float, float]
    layer: str
    semantic_class: str | None


@dataclass(slots=True)
class DxfExtractionResult:
    """Raw extraction result from a DXF drawing."""

    segments: list[RawSegment] = field(default_factory=list)
    blocks: list[RawBlock] = field(default_factory=list)
    texts: list[RawText] = field(default_factory=list)
    unknown_layers: list[str] = field(default_factory=list)
    unknown_blocks: list[str] = field(default_factory=list)


def load_layer_map(path: str | Path) -> dict[str, str]:
    """Load ``layer_map.csv`` into a layer-name to semantic-class mapping."""
    return _load_two_column_map(path, key_field="layer_name", value_field="class")


def load_block_map(path: str | Path) -> dict[str, str]:
    """Load ``block_map.csv`` into a block-name to semantic-class mapping."""
    return _load_two_column_map(path, key_field="block_name", value_field="class")


def extract_dxf(
    path: str | Path,
    layer_map_path: str | Path,
    block_map_path: str | Path,
) -> DxfExtractionResult:
    """Extract raw pipe segments, block inserts, and diameter texts from a DXF file."""
    dxf_path = Path(path)
    if not dxf_path.exists():
        raise FileNotFoundError(f"DXF file not found: {dxf_path}")
    layer_map = load_layer_map(layer_map_path)
    block_map = load_block_map(block_map_path)
    doc = ezdxf.readfile(dxf_path)
    modelspace = doc.modelspace()
    result = DxfExtractionResult()

    segment_counter = 1
    block_counter = 1
    text_counter = 1

    for entity in modelspace:
        entity_type = entity.dxftype()
        layer = str(entity.dxf.layer)
        layer_class = layer_map.get(layer)
        if layer_class is None and layer not in result.unknown_layers:
            result.unknown_layers.append(layer)

        if entity_type == "LINE" and layer_class in PIPE_CLASSES:
            result.segments.append(
                RawSegment(
                    segment_id=f"S{segment_counter:06d}",
                    start=_xy(entity.dxf.start),
                    end=_xy(entity.dxf.end),
                    layer=layer,
                    semantic_class=layer_class,
                )
            )
            segment_counter += 1
        elif entity_type == "LWPOLYLINE" and layer_class in PIPE_CLASSES:
            segments = _segments_from_lwpolyline(entity, f"S{segment_counter:06d}", layer, layer_class)
            result.segments.extend(segments)
            segment_counter += len(segments)
        elif entity_type == "POLYLINE" and layer_class in PIPE_CLASSES:
            segments = _segments_from_polyline(entity, f"S{segment_counter:06d}", layer, layer_class)
            result.segments.extend(segments)
            segment_counter += len(segments)
        elif entity_type == "INSERT":
            block_name = str(entity.dxf.name)
            if block_name not in block_map and block_name not in result.unknown_blocks:
                result.unknown_blocks.append(block_name)
            result.blocks.append(
                RawBlock(
                    block_id=f"B{block_counter:06d}",
                    block_name=block_name,
                    insert=_xy(entity.dxf.insert),
                    layer=layer,
                    semantic_class=block_map.get(block_name, layer_class),
                    rotation=_optional_dxf_float(entity, "rotation"),
                )
            )
            block_counter += 1
        elif entity_type in {"TEXT", "MTEXT"}:
            raw_text = _entity_text(entity)
            diameter_text = normalize_diameter_text(raw_text)
            if diameter_text is not None:
                result.texts.append(
                    RawText(
                        text_id=f"T{text_counter:06d}",
                        text=diameter_text,
                        insert=_text_insert(entity),
                        layer=layer,
                        semantic_class=layer_class,
                    )
                )
                text_counter += 1

    return result


def write_dxf_extraction_tables(result: DxfExtractionResult, output_dir: Path) -> None:
    """Write raw DXF extraction tables for human review."""
    output_dir.mkdir(parents=True, exist_ok=True)
    pd.DataFrame(
        [
            {
                "segment_id": segment.segment_id,
                "start_x": segment.start[0],
                "start_y": segment.start[1],
                "end_x": segment.end[0],
                "end_y": segment.end[1],
                "layer": segment.layer,
                "semantic_class": segment.semantic_class,
                "length": dist(segment.start, segment.end),
            }
            for segment in result.segments
        ],
        columns=["segment_id", "start_x", "start_y", "end_x", "end_y", "layer", "semantic_class", "length"],
    ).to_csv(output_dir / "raw_segments.csv", index=False)

    pd.DataFrame(
        [
            {
                "block_id": block.block_id,
                "block_name": block.block_name,
                "x": block.insert[0],
                "y": block.insert[1],
                "layer": block.layer,
                "semantic_class": block.semantic_class,
                "rotation": block.rotation,
            }
            for block in result.blocks
        ],
        columns=["block_id", "block_name", "x", "y", "layer", "semantic_class", "rotation"],
    ).to_csv(output_dir / "raw_blocks.csv", index=False)

    pd.DataFrame(
        [
            {
                "text_id": text.text_id,
                "text": text.text,
                "x": text.insert[0],
                "y": text.insert[1],
                "layer": text.layer,
                "semantic_class": text.semantic_class,
                "normalized_diameter": normalize_diameter_text(text.text),
            }
            for text in result.texts
        ],
        columns=["text_id", "text", "x", "y", "layer", "semantic_class", "normalized_diameter"],
    ).to_csv(output_dir / "raw_texts.csv", index=False)

    pd.DataFrame(
        [
            {"item": "segments", "count": len(result.segments)},
            {"item": "blocks", "count": len(result.blocks)},
            {"item": "texts", "count": len(result.texts)},
            {"item": "unknown_layers", "count": len(result.unknown_layers)},
            {"item": "unknown_blocks", "count": len(result.unknown_blocks)},
        ],
        columns=["item", "count"],
    ).to_csv(output_dir / "extraction_summary.csv", index=False)

    _extraction_warnings_dataframe(validate_dxf_extraction(result)).to_csv(
        output_dir / "extraction_warnings.csv",
        index=False,
    )


def validate_dxf_extraction(result: DxfExtractionResult) -> list[ValidationIssue]:
    """Return human-review warnings for raw DXF extraction results."""
    issues: list[ValidationIssue] = []
    if not result.segments:
        issues.append(
            ValidationIssue("WARNING", "NO_PIPE_SEGMENTS", "No pipe segments were extracted.", "DXF", None)
        )
    if not any(block.semantic_class == "head" for block in result.blocks):
        issues.append(ValidationIssue("WARNING", "NO_HEAD_BLOCKS", "No head blocks were extracted.", "DXF", None))
    if not result.texts:
        issues.append(
            ValidationIssue("WARNING", "NO_DIAMETER_TEXT", "No diameter text was extracted.", "DXF", None)
        )
    for layer in sorted(result.unknown_layers):
        issues.append(
            ValidationIssue(
                "WARNING",
                "UNKNOWN_LAYER",
                f"DXF layer {layer!r} was not found in layer_map.csv.",
                "Layer",
                layer,
            )
        )
    for block_name in sorted(result.unknown_blocks):
        issues.append(
            ValidationIssue(
                "WARNING",
                "UNKNOWN_BLOCK",
                f"DXF block {block_name!r} was not found in block_map.csv.",
                "Block",
                block_name,
            )
        )
    return issues


def _extraction_warnings_dataframe(issues: list[ValidationIssue]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "severity": issue.severity,
                "code": issue.code,
                "message": issue.message,
                "object_type": issue.object_type,
                "object_id": issue.object_id,
            }
            for issue in issues
        ],
        columns=["severity", "code", "message", "object_type", "object_id"],
    )


def is_diameter_text(text: str) -> bool:
    """Return whether text contains a supported nominal diameter label."""
    return normalize_diameter_text(text) is not None


def normalize_diameter_text(text: str) -> str | None:
    """Normalize supported diameter labels such as ``150 A`` to ``150A``."""
    match = DIAMETER_PATTERN.search(text.replace("\\P", " "))
    if match is None:
        return None
    label = f"{match.group(1).upper()}A"
    return label if label in SUPPORTED_DIAMETER_LABELS else None


def _load_two_column_map(path: str | Path, key_field: str, value_field: str) -> dict[str, str]:
    mapping_path = Path(path)
    if not mapping_path.exists():
        raise FileNotFoundError(f"Mapping CSV file not found: {mapping_path}")
    with mapping_path.open(newline="", encoding="utf-8-sig") as csv_file:
        reader = csv.DictReader(csv_file)
        _require_columns(reader.fieldnames, {key_field, value_field}, mapping_path)
        return {
            str(row[key_field]).strip(): str(row[value_field]).strip()
            for row in reader
            if row.get(key_field) and row.get(value_field)
        }


def _require_columns(fieldnames: list[str] | None, required: set[str], path: Path) -> None:
    available = set(fieldnames or [])
    missing = sorted(required - available)
    if missing:
        raise ValueError(
            f"CSV file {path} is missing required columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(sorted(available))}"
        )


def _segments_from_lwpolyline(
    entity: Any,
    first_segment_id: str,
    layer: str,
    semantic_class: str | None,
) -> list[RawSegment]:
    points = [(float(point[0]), float(point[1])) for point in entity.get_points()]
    if entity.closed and points:
        points.append(points[0])
    return _segments_from_points(points, first_segment_id, layer, semantic_class)


def _segments_from_polyline(
    entity: Any,
    first_segment_id: str,
    layer: str,
    semantic_class: str | None,
) -> list[RawSegment]:
    points = [_xy(vertex.dxf.location) for vertex in entity.vertices]
    if entity.is_closed and points:
        points.append(points[0])
    return _segments_from_points(points, first_segment_id, layer, semantic_class)


def _segments_from_points(
    points: list[tuple[float, float]],
    first_segment_id: str,
    layer: str,
    semantic_class: str | None,
) -> list[RawSegment]:
    if len(points) < 2:
        return []

    prefix = first_segment_id[0]
    start_number = int(first_segment_id[1:])
    segments: list[RawSegment] = []
    for offset, (start, end) in enumerate(zip(points, points[1:], strict=False)):
        segments.append(
            RawSegment(
                segment_id=f"{prefix}{start_number + offset:06d}",
                start=start,
                end=end,
                layer=layer,
                semantic_class=semantic_class,
            )
        )
    return segments


def _entity_text(entity: Any) -> str:
    if entity.dxftype() == "MTEXT":
        return str(entity.text)
    return str(entity.dxf.text)


def _text_insert(entity: Any) -> tuple[float, float]:
    if entity.dxftype() == "MTEXT":
        return _xy(entity.dxf.insert)
    return _xy(entity.dxf.insert)


def _xy(point: Any) -> tuple[float, float]:
    return (float(point[0]), float(point[1]))


def _optional_dxf_float(entity: Any, attr_name: str) -> float | None:
    if not entity.dxf.hasattr(attr_name):
        return None
    return float(entity.dxf.get(attr_name))
