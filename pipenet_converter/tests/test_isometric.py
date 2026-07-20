"""Tests for isometric PNG rendering."""

from pathlib import Path

from pipenet_converter.isometric import project_iso, render_isometric_png
from pipenet_converter.models import Node, Nozzle, Pipe, PipeNetwork, ValidationIssue


def test_project_iso_projects_coordinates() -> None:
    projected = project_iso(10.0, 5.0, 2.0, scale=1.0, z_scale=1.0)

    assert round(projected[0], 6) == 4.330127
    assert round(projected[1], 6) == 5.5


def test_render_isometric_png_creates_non_empty_file(tmp_path: Path) -> None:
    output_path = tmp_path / "network.png"

    render_isometric_png(_sample_network(), output_path)

    assert output_path.exists()
    assert output_path.stat().st_size > 0


def test_render_isometric_png_accepts_validation_issues(tmp_path: Path) -> None:
    output_path = tmp_path / "network_with_issues.png"
    issues = [
        ValidationIssue("ERROR", "PIPE_LENGTH_INVALID", "bad length", "Pipe", "P001"),
        ValidationIssue("WARNING", "ORPHAN_NODE", "orphan", "Node", "N003"),
    ]

    render_isometric_png(_sample_network(), output_path, issues=issues)

    assert output_path.exists()
    assert output_path.stat().st_size > 0


def _sample_network() -> PipeNetwork:
    network = PipeNetwork(title="Isometric sample")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input"))
    network.add_node(Node("N001", 10.0, 0.0, 0.0, "main_pipe"))
    network.add_node(Node("N002", 10.0, 5.0, 1.0, "head"))
    network.add_node(Node("N003", 20.0, 5.0, 1.0, "pipe"))
    network.add_pipe(Pipe("P001", "INPUT", "N001", 0.15, 10.0, 0.0))
    network.add_pipe(Pipe("P002", "N001", "N002", 0.032, 6.0, 1.0))
    network.add_nozzle(Nozzle("NZ001", "N002", "@/1", 0.00266666667, status=1))
    return network
