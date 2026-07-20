"""Command-line interface for the PipeNet converter."""

from __future__ import annotations

import argparse
import logging
from pathlib import Path
from typing import Sequence

from pipenet_converter.comparator import CompareTolerance, compare_networks, write_comparison_report
from pipenet_converter.config import PipelineConfig, load_pipeline_config
from pipenet_converter.dxf_extractor import extract_dxf, write_dxf_extraction_tables
from pipenet_converter.elevation import (
    apply_elevation_rules,
    load_elevation_rules,
    recompute_pipe_rise_and_length,
)
from pipenet_converter.export_tables import read_network_tables, write_network_tables
from pipenet_converter.graph_builder import (
    assign_diameters_to_edges,
    attach_blocks_to_graph,
    build_2d_graph_from_segments,
    graph_to_pipenetwork_2d,
)
from pipenet_converter.isometric import render_isometric_png
from pipenet_converter.sdf_parser import parse_sdf
from pipenet_converter.sdf_writer import write_sdf
from pipenet_converter.pipeline import run_pipeline_from_config
from pipenet_converter.system_graph import merge_networks, parse_system_edges_csv
from pipenet_converter.validator import (
    apply_detected_fittings,
    validate_network,
    write_validation_report,
)


def main(argv: Sequence[str] | None = None) -> None:
    """Run the command-line interface."""
    parser = _build_parser()
    args = parser.parse_args(argv)
    _configure_logging(args.verbose)
    logging.info("Running command: %s", args.command)
    args.func(args)


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="pipenet-converter")
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--verbose", action="store_true", help="Enable verbose logging.")
    common.add_argument(
        "--strict",
        action="store_true",
        help="Exit with status 1 when validation ERROR issues are produced.",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    parse_sdf_parser = subparsers.add_parser("parse-sdf", parents=[common], help="Parse SDF and export CSV tables.")
    parse_sdf_parser.add_argument("--sdf", required=True)
    parse_sdf_parser.add_argument("--out-dir", required=True)
    parse_sdf_parser.set_defaults(func=_cmd_parse_sdf)

    extract_dxf_parser = subparsers.add_parser("extract-dxf", parents=[common], help="Extract raw DXF entities.")
    extract_dxf_parser.add_argument("--dxf", required=True)
    extract_dxf_parser.add_argument("--layer-map", required=True)
    extract_dxf_parser.add_argument("--block-map", required=True)
    extract_dxf_parser.add_argument("--out-dir", required=True)
    extract_dxf_parser.set_defaults(func=_cmd_extract_dxf)

    build_network_parser = subparsers.add_parser(
        "build-network", parents=[common], help="Build a plan network from DXF."
    )
    build_network_parser.add_argument("--dxf", required=True)
    build_network_parser.add_argument("--layer-map", required=True)
    build_network_parser.add_argument("--block-map", required=True)
    build_network_parser.add_argument("--elevation-rules", required=True)
    build_network_parser.add_argument("--out-dir", required=True)
    build_network_parser.set_defaults(func=_cmd_build_network)

    merge_system_parser = subparsers.add_parser(
        "merge-system", parents=[common], help="Merge system/riser CSV with network tables."
    )
    merge_system_parser.add_argument("--network-dir", required=True)
    merge_system_parser.add_argument("--system-edges", required=True)
    merge_system_parser.add_argument("--connection-map", required=True)
    merge_system_parser.add_argument("--out-dir", required=True)
    merge_system_parser.set_defaults(func=_cmd_merge_system)

    write_sdf_parser = subparsers.add_parser("write-sdf", parents=[common], help="Write PipeNet SDF from network tables.")
    write_sdf_parser.add_argument("--network-dir", required=True)
    write_sdf_parser.add_argument("--template", required=False, default=None)
    write_sdf_parser.add_argument("--output", required=True)
    write_sdf_parser.set_defaults(func=_cmd_write_sdf)

    render_iso_parser = subparsers.add_parser("render-iso", parents=[common], help="Render isometric check PNG.")
    render_iso_parser.add_argument("--network-dir", required=True)
    render_iso_parser.add_argument("--output", required=True)
    render_iso_parser.set_defaults(func=_cmd_render_iso)

    validate_parser = subparsers.add_parser("validate", parents=[common], help="Validate network tables.")
    validate_parser.add_argument("--network-dir", required=True)
    validate_parser.add_argument("--input-node", required=False, default=None)
    validate_parser.add_argument("--out", required=True)
    validate_parser.set_defaults(func=_cmd_validate)

    compare_sdf_parser = subparsers.add_parser(
        "compare-sdf", parents=[common], help="Compare reference and candidate SDF files."
    )
    compare_sdf_parser.add_argument("--reference", required=True)
    compare_sdf_parser.add_argument("--candidate", required=True)
    compare_sdf_parser.add_argument("--out", required=True)
    compare_sdf_parser.set_defaults(func=_cmd_compare_sdf)

    run_pipeline_parser = subparsers.add_parser(
        "run-pipeline", parents=[common], help="Run the full DXF to SDF pipeline."
    )
    run_pipeline_parser.add_argument("--plan-dxf", required=True)
    run_pipeline_parser.add_argument("--layer-map", required=True)
    run_pipeline_parser.add_argument("--block-map", required=True)
    run_pipeline_parser.add_argument("--elevation-rules", required=True)
    run_pipeline_parser.add_argument("--system-edges", required=True)
    run_pipeline_parser.add_argument("--connection-map", required=True)
    run_pipeline_parser.add_argument("--template-sdf", required=False, default=None)
    run_pipeline_parser.add_argument("--reference-sdf", required=False, default=None)
    run_pipeline_parser.add_argument("--out-dir", required=True)
    run_pipeline_parser.set_defaults(func=_cmd_run_pipeline)

    run_config_parser = subparsers.add_parser("run-config", parents=[common], help="Run pipeline from YAML config.")
    run_config_parser.add_argument("--config", required=True)
    run_config_parser.set_defaults(func=_cmd_run_config)

    return parser


def _cmd_parse_sdf(args: argparse.Namespace) -> None:
    network = parse_sdf(args.sdf)
    output_dir = Path(args.out_dir)
    write_network_tables(network, output_dir)
    issues = validate_network(network)
    write_validation_report(issues, output_dir / "validation_report.csv")
    _exit_if_strict_errors(args, issues)


def _cmd_extract_dxf(args: argparse.Namespace) -> None:
    result = extract_dxf(args.dxf, args.layer_map, args.block_map)
    output_dir = Path(args.out_dir)
    write_dxf_extraction_tables(result, output_dir)


def _cmd_build_network(args: argparse.Namespace) -> None:
    output_dir = Path(args.out_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    result = extract_dxf(args.dxf, args.layer_map, args.block_map)
    write_dxf_extraction_tables(result, output_dir)

    graph = build_2d_graph_from_segments(result.segments, snap_tolerance=50.0)
    attach_blocks_to_graph(graph, result.blocks, max_distance=600.0)
    assign_diameters_to_edges(graph, result.texts, default_diameter_m=0.032)
    network = graph_to_pipenetwork_2d(graph, title=Path(args.dxf).stem)
    rules = load_elevation_rules(args.elevation_rules)
    apply_elevation_rules(network, rules)
    recompute_pipe_rise_and_length(network)
    apply_detected_fittings(network)

    issues = validate_network(network)
    write_network_tables(network, output_dir)
    write_validation_report(issues, output_dir / "validation_report.csv")
    render_isometric_png(network, output_dir / "isometric_check.png", issues=issues)
    _exit_if_strict_errors(args, issues)


def _cmd_merge_system(args: argparse.Namespace) -> None:
    plan_network = read_network_tables(args.network_dir)
    system_network = parse_system_edges_csv(args.system_edges)
    import json

    connection_map = json.loads(Path(args.connection_map).read_text(encoding="utf-8"))
    merged = merge_networks(system_network, plan_network, connection_map)
    issues = validate_network(merged, input_node_id="INPUT")
    output_dir = Path(args.out_dir)
    write_network_tables(merged, output_dir)
    write_validation_report(issues, output_dir / "validation_report.csv")
    _exit_if_strict_errors(args, issues)


def _cmd_write_sdf(args: argparse.Namespace) -> None:
    network = read_network_tables(args.network_dir)
    output = Path(args.output)
    write_sdf(network, output, template_path=args.template)
    round_trip = parse_sdf(output)
    issues = validate_network(round_trip)
    write_validation_report(issues, output.with_suffix(".validation_report.csv"))
    _exit_if_strict_errors(args, issues)


def _cmd_render_iso(args: argparse.Namespace) -> None:
    network = read_network_tables(args.network_dir)
    render_isometric_png(network, args.output)


def _cmd_validate(args: argparse.Namespace) -> None:
    network = read_network_tables(args.network_dir)
    issues = validate_network(network, input_node_id=args.input_node)
    write_validation_report(issues, args.out)
    _exit_if_strict_errors(args, issues)


def _cmd_compare_sdf(args: argparse.Namespace) -> None:
    reference = parse_sdf(args.reference)
    candidate = parse_sdf(args.candidate)
    issues = compare_networks(reference, candidate, CompareTolerance())
    write_comparison_report(issues, args.out)


def _cmd_run_pipeline(args: argparse.Namespace) -> None:
    config = PipelineConfig(
        project_name=Path(args.plan_dxf).stem,
        plan_dxf=args.plan_dxf,
        layer_map=args.layer_map,
        block_map=args.block_map,
        elevation_rules=args.elevation_rules,
        system_edges=args.system_edges,
        connection_map=args.connection_map,
        template_sdf=args.template_sdf,
        reference_sdf=args.reference_sdf,
        output_dir=args.out_dir,
        snap_tolerance=50.0,
        head_attach_tolerance=600.0,
        cad_unit_scale_to_m=0.001,
        default_diameter_m=0.032,
        input_node_id="INPUT",
    )
    result = run_pipeline_from_config(config)
    _exit_if_strict_error_count(args, result.validation_error_count)


def _cmd_run_config(args: argparse.Namespace) -> None:
    config = load_pipeline_config(args.config)
    result = run_pipeline_from_config(config)
    _exit_if_strict_error_count(args, result.validation_error_count)


def _exit_if_strict_errors(args: argparse.Namespace, issues) -> None:
    if not getattr(args, "strict", False):
        return
    error_count = sum(1 for issue in issues if issue.severity == "ERROR")
    if error_count:
        logging.error("Strict mode failed: %s validation ERROR issue(s).", error_count)
        raise SystemExit(1)


def _exit_if_strict_error_count(args: argparse.Namespace, validation_error_count: int) -> None:
    if getattr(args, "strict", False) and validation_error_count:
        logging.error("Strict mode failed: %s validation ERROR issue(s).", validation_error_count)
        raise SystemExit(1)


def _configure_logging(verbose: bool) -> None:
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.WARNING,
        format="%(levelname)s:%(name)s:%(message)s",
        force=True,
    )


if __name__ == "__main__":
    main()
