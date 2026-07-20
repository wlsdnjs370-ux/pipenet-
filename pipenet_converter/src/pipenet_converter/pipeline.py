"""End-to-end CAD-to-PipeNet pipeline orchestration."""

from __future__ import annotations

from dataclasses import dataclass
import json
from pathlib import Path

from pipenet_converter.comparator import CompareTolerance, compare_networks, write_comparison_report
from pipenet_converter.config import PipelineConfig
from pipenet_converter.dxf_extractor import extract_dxf, write_dxf_extraction_tables
from pipenet_converter.elevation import apply_elevation_rules, load_elevation_rules, recompute_pipe_rise_and_length
from pipenet_converter.export_tables import write_network_tables
from pipenet_converter.graph_builder import (
    assign_diameters_to_edges,
    attach_blocks_to_graph,
    build_2d_graph_from_segments,
    graph_to_pipenetwork_2d,
)
from pipenet_converter.isometric import render_isometric_png
from pipenet_converter.redline import create_redline_items, write_redline_items
from pipenet_converter.sdf_parser import parse_sdf
from pipenet_converter.sdf_writer import write_sdf
from pipenet_converter.system_graph import merge_networks, parse_system_edges_csv
from pipenet_converter.validator import apply_detected_fittings, validate_network, write_validation_report


@dataclass(slots=True)
class PipelineRunResult:
    """Summary of a full pipeline run."""

    output_dir: Path
    generated_sdf_path: Path
    isometric_path: Path
    validation_issue_count: int
    validation_error_count: int
    comparison_issue_count: int


def run_pipeline_from_config(config: PipelineConfig) -> PipelineRunResult:
    """Run the full pipeline using a loaded configuration."""
    out_dir = Path(config.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    extraction = extract_dxf(config.plan_dxf, config.layer_map, config.block_map)
    write_dxf_extraction_tables(extraction, out_dir)

    graph = build_2d_graph_from_segments(extraction.segments, snap_tolerance=config.snap_tolerance)
    attach_blocks_to_graph(graph, extraction.blocks, max_distance=config.head_attach_tolerance)
    assign_diameters_to_edges(graph, extraction.texts, default_diameter_m=config.default_diameter_m)
    plan_network = graph_to_pipenetwork_2d(graph, title=config.project_name)

    rules = load_elevation_rules(config.elevation_rules)
    apply_elevation_rules(plan_network, rules)
    recompute_pipe_rise_and_length(plan_network)
    apply_detected_fittings(plan_network)

    system_network = parse_system_edges_csv(config.system_edges)
    connection_map = json.loads(Path(config.connection_map).read_text(encoding="utf-8"))
    merged_network = merge_networks(system_network, plan_network, connection_map)

    validation_issues = validate_network(merged_network, input_node_id=config.input_node_id)
    validation_error_count = sum(1 for issue in validation_issues if issue.severity == "ERROR")
    write_validation_report(validation_issues, out_dir / "validation_report.csv")
    write_network_tables(merged_network, out_dir)

    isometric_path = out_dir / "isometric_check.png"
    render_isometric_png(merged_network, isometric_path, issues=validation_issues)

    generated_sdf_path = out_dir / "generated_pipenet.sdf"
    write_sdf(merged_network, generated_sdf_path, template_path=config.template_sdf)
    generated_network = parse_sdf(generated_sdf_path)

    comparison_issues = []
    redline_iso_path: Path | None = None
    if config.reference_sdf:
        reference_network = parse_sdf(config.reference_sdf)
        comparison_issues = compare_networks(reference_network, generated_network, CompareTolerance())
        write_comparison_report(comparison_issues, out_dir / "compare_report.csv")
        redline_items = create_redline_items(comparison_issues, reference_network, generated_network)
        if redline_items:
            write_redline_items(redline_items, out_dir)
            redline_iso_path = out_dir / "isometric_redline.png"
            render_isometric_png(generated_network, redline_iso_path, redline_items=redline_items)

    _write_pipeline_summary(
        output_path=out_dir / "run_summary.txt",
        config=config,
        network=generated_network,
        validation_issue_count=len(validation_issues),
        comparison_issue_count=len(comparison_issues),
        generated_sdf_path=generated_sdf_path,
        isometric_path=isometric_path,
        redline_iso_path=redline_iso_path,
    )
    return PipelineRunResult(
        output_dir=out_dir,
        generated_sdf_path=generated_sdf_path,
        isometric_path=isometric_path,
        validation_issue_count=len(validation_issues),
        validation_error_count=validation_error_count,
        comparison_issue_count=len(comparison_issues),
    )


def _write_pipeline_summary(
    output_path: Path,
    config: PipelineConfig,
    network,
    validation_issue_count: int,
    comparison_issue_count: int,
    generated_sdf_path: Path,
    isometric_path: Path,
    redline_iso_path: Path | None,
) -> None:
    lines = [
        "PipeNet Converter Run Summary",
        "",
        "Input files:",
        f"project_name: {config.project_name}",
        f"plan_dxf: {config.plan_dxf}",
        f"layer_map: {config.layer_map}",
        f"block_map: {config.block_map}",
        f"elevation_rules: {config.elevation_rules}",
        f"system_edges: {config.system_edges}",
        f"connection_map: {config.connection_map}",
        f"template_sdf: {config.template_sdf}",
        f"reference_sdf: {config.reference_sdf}",
        "",
        f"output_directory: {config.output_dir}",
        f"node_count: {len(network.nodes)}",
        f"pipe_count: {len(network.pipes)}",
        f"nozzle_count: {len(network.nozzles)}",
        f"active_nozzle_count: {len(network.active_nozzles())}",
        f"validation_issue_count: {validation_issue_count}",
        f"comparison_issue_count: {comparison_issue_count}",
        f"generated_sdf_path: {generated_sdf_path}",
        f"isometric_png_path: {isometric_path}",
    ]
    if redline_iso_path is not None:
        lines.append(f"redline_isometric_png_path: {redline_iso_path}")
    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
