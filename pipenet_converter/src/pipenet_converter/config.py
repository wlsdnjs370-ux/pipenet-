"""YAML configuration loading for full pipeline runs."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import yaml


@dataclass(slots=True)
class PipelineConfig:
    """Configuration for running the full CAD-to-SDF pipeline."""

    project_name: str
    plan_dxf: str
    layer_map: str
    block_map: str
    elevation_rules: str
    system_edges: str
    connection_map: str
    template_sdf: str | None
    reference_sdf: str | None
    output_dir: str
    snap_tolerance: float
    head_attach_tolerance: float
    cad_unit_scale_to_m: float
    default_diameter_m: float
    input_node_id: str


REQUIRED_CONFIG_FIELDS = {
    "project_name",
    "plan_dxf",
    "layer_map",
    "block_map",
    "elevation_rules",
    "system_edges",
    "connection_map",
    "output_dir",
}


def load_pipeline_config(path: str | Path) -> PipelineConfig:
    """Load a pipeline YAML configuration file."""
    config_path = Path(path)
    if not config_path.exists():
        raise FileNotFoundError(f"Pipeline config file not found: {config_path}")

    data = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    if not isinstance(data, dict):
        raise ValueError(f"Pipeline config {config_path} must contain a YAML mapping.")
    missing = sorted(REQUIRED_CONFIG_FIELDS - set(data))
    if missing:
        raise ValueError(f"Pipeline config {config_path} is missing required fields: {', '.join(missing)}")

    base_dir = config_path.parent
    return PipelineConfig(
        project_name=str(data["project_name"]),
        plan_dxf=_resolve_path(data["plan_dxf"], base_dir),
        layer_map=_resolve_path(data["layer_map"], base_dir),
        block_map=_resolve_path(data["block_map"], base_dir),
        elevation_rules=_resolve_path(data["elevation_rules"], base_dir),
        system_edges=_resolve_path(data["system_edges"], base_dir),
        connection_map=_resolve_path(data["connection_map"], base_dir),
        template_sdf=_resolve_optional_path(data.get("template_sdf"), base_dir),
        reference_sdf=_resolve_optional_path(data.get("reference_sdf"), base_dir),
        output_dir=_resolve_path(data["output_dir"], base_dir),
        snap_tolerance=_float_field(data, "snap_tolerance", 50.0),
        head_attach_tolerance=_float_field(data, "head_attach_tolerance", 600.0),
        cad_unit_scale_to_m=_float_field(data, "cad_unit_scale_to_m", 0.001),
        default_diameter_m=_float_field(data, "default_diameter_m", 0.032),
        input_node_id=str(data.get("input_node_id", "INPUT")),
    )


def _resolve_path(value: Any, base_dir: Path) -> str:
    path = Path(str(value))
    if not path.is_absolute():
        path = base_dir / path
    return str(path.resolve(strict=False))


def _resolve_optional_path(value: Any, base_dir: Path) -> str | None:
    if value in {None, ""}:
        return None
    return _resolve_path(value, base_dir)


def _float_field(data: dict[str, Any], key: str, default: float) -> float:
    try:
        return float(data.get(key, default))
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Pipeline config field {key!r} must be numeric.") from exc
