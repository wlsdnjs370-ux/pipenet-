"""Mapping review helpers for real CAD project onboarding."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
import re

import ezdxf
import pandas as pd

from pipenet_converter.sdf_parser import parse_sdf


DAISO_SDF_NAME = "1-1. 다이소 세종허브센터 지상4층 창고.sdf"
DAISO_TARGET_DXF_NAME = "MF-110~113 지상4층 상부 소방시설 확대평면도.dxf"


@dataclass(slots=True)
class DxfInventory:
    """Raw DXF layer and block inventory used for mapping review."""

    entity_layer_counts: Counter[str] = field(default_factory=Counter)
    pipe_candidate_layers: list[str] = field(default_factory=list)
    diameter_text_candidate_layers: list[str] = field(default_factory=list)
    block_counts: Counter[str] = field(default_factory=Counter)
    block_layers: dict[str, set[str]] = field(default_factory=dict)
    text_count: int = 0
    diameter_text_count: int = 0


def prepare_daiso_4f_mapping_review(
    input_dir: str | Path,
    output_dir: str | Path,
) -> Path:
    """Create draft mapping files and a markdown review for the Daiso 4F target plan."""
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    sdf_path = input_path / DAISO_SDF_NAME
    dxf_path = input_path / DAISO_TARGET_DXF_NAME

    sdf_summary = _summarize_sdf(sdf_path)
    inventory = _inventory_dxf(dxf_path) if dxf_path.exists() else DxfInventory()

    layer_map = _draft_layer_map(inventory)
    block_map = _draft_block_map(inventory)
    layer_map_path = output_path / "layer_map_daiso_4f.csv"
    block_map_path = output_path / "block_map_daiso_4f.csv"
    layer_map.to_csv(layer_map_path, index=False)
    block_map.to_csv(block_map_path, index=False)

    _write_extraction_summary(inventory, output_path / "extraction_summary.csv")
    report_path = output_path / "daiso_4f_mapping_review.md"
    report_path.write_text(
        _render_report(
            input_path=input_path,
            output_path=output_path,
            sdf_path=sdf_path,
            dxf_path=dxf_path,
            sdf_summary=sdf_summary,
            inventory=inventory,
            layer_map=layer_map,
            block_map=block_map,
        ),
        encoding="utf-8",
    )
    return report_path


def _summarize_sdf(sdf_path: Path) -> dict[str, int | str]:
    if not sdf_path.exists():
        return {"status": "missing", "nodes": 0, "pipes": 0, "nozzles": 0, "active_nozzles": 0}
    network = parse_sdf(sdf_path)
    return {
        "status": "parsed",
        "nodes": len(network.nodes),
        "pipes": len(network.pipes),
        "nozzles": len(network.nozzles),
        "active_nozzles": len(network.active_nozzles()),
    }


def _inventory_dxf(dxf_path: Path) -> DxfInventory:
    doc = ezdxf.readfile(dxf_path)
    modelspace = doc.modelspace()
    inventory = DxfInventory()
    for entity in modelspace:
        entity_type = entity.dxftype()
        layer = str(entity.dxf.layer)
        inventory.entity_layer_counts[layer] += 1
        if entity_type in {"LINE", "LWPOLYLINE", "POLYLINE", "ARC"} and _looks_like_pipe_layer(layer):
            _append_unique(inventory.pipe_candidate_layers, layer)
        if entity_type == "INSERT":
            block_name = str(entity.dxf.name)
            inventory.block_counts[block_name] += 1
            inventory.block_layers.setdefault(block_name, set()).add(layer)
        if entity_type in {"TEXT", "MTEXT"}:
            inventory.text_count += 1
            text = _entity_text(entity)
            if _looks_like_diameter_text(text):
                inventory.diameter_text_count += 1
                _append_unique(inventory.diameter_text_candidate_layers, layer)
    return inventory


def _draft_layer_map(inventory: DxfInventory) -> pd.DataFrame:
    rows = []
    for layer, count in sorted(inventory.entity_layer_counts.items()):
        guessed_class = _guess_layer_class(layer, inventory)
        rows.append(
            {
                "layer_name": layer,
                "class": guessed_class,
                "description": f"guessed from layer name; entity_count={count}",
            }
        )
    return pd.DataFrame(rows, columns=["layer_name", "class", "description"])


def _draft_block_map(inventory: DxfInventory) -> pd.DataFrame:
    rows = []
    for block_name, count in sorted(inventory.block_counts.items()):
        guessed_class = _guess_block_class(block_name)
        layers = ",".join(sorted(inventory.block_layers.get(block_name, set())))
        rows.append(
            {
                "block_name": block_name,
                "class": guessed_class,
                "description": f"guessed from block name; insert_count={count}; layers={layers}",
            }
        )
    return pd.DataFrame(rows, columns=["block_name", "class", "description"])


def _write_extraction_summary(inventory: DxfInventory, output_path: Path) -> None:
    pd.DataFrame(
        [
            {"item": "segments", "count": sum(inventory.entity_layer_counts[layer] for layer in inventory.pipe_candidate_layers)},
            {"item": "blocks", "count": sum(inventory.block_counts.values())},
            {"item": "texts", "count": inventory.text_count},
            {"item": "diameter_texts", "count": inventory.diameter_text_count},
            {"item": "layers", "count": len(inventory.entity_layer_counts)},
            {"item": "block_names", "count": len(inventory.block_counts)},
        ],
        columns=["item", "count"],
    ).to_csv(output_path, index=False)


def _render_report(
    input_path: Path,
    output_path: Path,
    sdf_path: Path,
    dxf_path: Path,
    sdf_summary: dict[str, int | str],
    inventory: DxfInventory,
    layer_map: pd.DataFrame,
    block_map: pd.DataFrame,
) -> str:
    likely_pipe_layers = _rows_by_class(layer_map, {"pipe", "main_pipe", "branch_pipe", "riser"})
    likely_text_layers = _rows_by_class(layer_map, {"diameter_text"})
    likely_head_blocks = _rows_by_class(block_map, {"head"})
    likely_valve_blocks = _rows_by_class(block_map, {"gate", "check", "alarm_valve"})
    unclassified_layers = _rows_by_class(layer_map, {"background", "ignore"})
    unclassified_blocks = _rows_by_class(block_map, {"ignore"})

    return "\n".join(
        [
            "# Daiso 4F Mapping Review",
            "",
            "## Input Files",
            f"- input_dir: `{input_path}`",
            f"- reference_sdf: `{sdf_path}` ({'present' if sdf_path.exists() else 'missing'})",
            f"- target_plan_dxf: `{dxf_path}` ({'present' if dxf_path.exists() else 'missing'})",
            "",
            "## SDF Summary",
            f"- status: {sdf_summary['status']}",
            f"- node count: {sdf_summary['nodes']}",
            f"- pipe count: {sdf_summary['pipes']}",
            f"- nozzle count: {sdf_summary['nozzles']}",
            f"- active nozzle count: {sdf_summary['active_nozzles']}",
            "",
            "## DXF Extraction Summary",
            f"- candidate pipe segment/entity count: {sum(inventory.entity_layer_counts[layer] for layer in inventory.pipe_candidate_layers)}",
            f"- block insert count: {sum(inventory.block_counts.values())}",
            f"- text entity count: {inventory.text_count}",
            f"- diameter text count: {inventory.diameter_text_count}",
            "",
            "## Likely Pipe Layers",
            _markdown_list(likely_pipe_layers, "layer_name"),
            "",
            "## Likely Head Blocks",
            _markdown_list(likely_head_blocks, "block_name"),
            "",
            "## Likely Valve Blocks",
            _markdown_list(likely_valve_blocks, "block_name"),
            "",
            "## Likely Diameter Text Layers",
            _markdown_list(likely_text_layers, "layer_name"),
            "",
            "## Layers Requiring Human Classification",
            _markdown_list(unclassified_layers, "layer_name"),
            "",
            "## Blocks Requiring Human Classification",
            _markdown_list(unclassified_blocks, "block_name"),
            "",
            "## Output Files",
            f"- `{output_path / 'layer_map_daiso_4f.csv'}`",
            f"- `{output_path / 'block_map_daiso_4f.csv'}`",
            f"- `{output_path / 'extraction_summary.csv'}`",
            "",
            "## Notes",
            "- Original drawings were not modified.",
            "- Unknown layers and blocks are intentionally retained for human classification.",
            "- Final SDF generation was not run in this review step.",
            "",
        ]
    )


def _guess_layer_class(layer: str, inventory: DxfInventory) -> str:
    upper = layer.upper()
    if layer in inventory.diameter_text_candidate_layers or any(token in upper for token in ["TEXT", "TXT", "SIZE", "DIA"]):
        return "diameter_text"
    if any(token in upper for token in ["HEAD", "SPHD", "SPRINKLER"]):
        return "head"
    if any(token in upper for token in ["ALARM", "A/V", "AV"]):
        return "alarm_valve"
    if "VALVE" in upper or "VLV" in upper:
        return "valve"
    if "RISER" in upper or "입상" in layer:
        return "riser"
    if any(token in upper for token in ["MAIN", "주배관"]):
        return "main_pipe"
    if any(token in upper for token in ["BRANCH", "가지"]):
        return "branch_pipe"
    if layer in inventory.pipe_candidate_layers or any(token in upper for token in ["PIPE", "F-SP", "SP-", "소화", "SPR"]):
        return "pipe"
    if any(token in upper for token in ["WALL", "GRID", "DIM", "HATCH", "XREF", "XR", "A-"]):
        return "background"
    return "ignore"


def _guess_block_class(block_name: str) -> str:
    upper = block_name.upper()
    if any(token in upper for token in ["HEAD", "SPHD", "SPRINKLER"]):
        return "head"
    if "GATE" in upper:
        return "gate"
    if "CHECK" in upper:
        return "check"
    if any(token in upper for token in ["ALARM", "A/V", "AV"]):
        return "alarm_valve"
    if "RISER" in upper or "입상" in block_name:
        return "riser"
    if "PUMP" in upper:
        return "pump"
    if any(token in upper for token in ["EQUIP", "FLEX", "JOINT"]):
        return "equipment"
    return "ignore"


def _looks_like_pipe_layer(layer: str) -> bool:
    upper = layer.upper()
    return any(token in upper for token in ["PIPE", "F-SP", "SP-", "SPR", "소화"])


def _looks_like_diameter_text(text: str) -> bool:
    return re.search(r"\b(25|32|40|50|65|80|100|125|150|200)\s*A\b", text, re.IGNORECASE) is not None


def _entity_text(entity) -> str:
    if entity.dxftype() == "MTEXT":
        return str(entity.text)
    return str(entity.dxf.text)


def _append_unique(items: list[str], value: str) -> None:
    if value not in items:
        items.append(value)


def _rows_by_class(dataframe: pd.DataFrame, classes: set[str]) -> list[str]:
    if dataframe.empty:
        return []
    name_column = "layer_name" if "layer_name" in dataframe.columns else "block_name"
    return dataframe.loc[dataframe["class"].isin(classes), name_column].astype(str).tolist()


def _markdown_list(items: list[str], column_name: str) -> str:
    if not items:
        return "- none found"
    return "\n".join(f"- `{item}`" for item in items)
