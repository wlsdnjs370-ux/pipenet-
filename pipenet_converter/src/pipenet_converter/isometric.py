"""Isometric check rendering for 3D pipe centerline graphs."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from math import cos, radians, sin
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

from pipenet_converter.models import Node, PipeNetwork, ValidationIssue
from pipenet_converter.redline import RedlineItem


@dataclass(slots=True)
class IsometricRenderOptions:
    """Options for isometric check PNG rendering."""

    show_node_labels: bool = False
    show_pipe_labels: bool = False
    show_nozzle_labels: bool = True
    scale: float = 1.0
    z_scale: float = 1.0
    figsize: tuple[float, float] = (12, 8)


def project_iso(
    x: float,
    y: float,
    z: float,
    scale: float = 1.0,
    z_scale: float = 1.0,
) -> tuple[float, float]:
    """Project 3D coordinates to a simple 2D isometric plane."""
    angle = radians(30)
    iso_x = scale * (x - y) * cos(angle)
    iso_y = scale * (x + y) * sin(angle) - z * z_scale
    return iso_x, iso_y


def render_isometric_png(
    network: PipeNetwork,
    output_path: str | Path,
    issues: list[ValidationIssue] | None = None,
    redline_items: list[RedlineItem] | None = None,
) -> None:
    """Render a simple isometric PNG for manual network checking."""
    options = IsometricRenderOptions()
    issue_list = issues or []
    issue_nodes = {issue.object_id for issue in issue_list if issue.object_type == "Node"}
    issue_pipes = {issue.object_id for issue in issue_list if issue.object_type == "Pipe"}

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    fig, ax = plt.subplots(figsize=options.figsize)
    projected_nodes = {
        node_id: project_iso(node.x, node.y, node.z, options.scale, options.z_scale)
        for node_id, node in network.nodes.items()
    }

    for pipe in network.pipes.values():
        if pipe.from_node not in projected_nodes or pipe.to_node not in projected_nodes:
            continue
        start = projected_nodes[pipe.from_node]
        end = projected_nodes[pipe.to_node]
        is_issue = pipe.pipe_id in issue_pipes
        ax.plot(
            [start[0], end[0]],
            [start[1], end[1]],
            linewidth=2.5 if is_issue else 1.5,
            color="C3" if is_issue else "C0",
        )
        if options.show_pipe_labels:
            midpoint = ((start[0] + end[0]) / 2, (start[1] + end[1]) / 2)
            ax.text(midpoint[0], midpoint[1], pipe.pipe_id, fontsize=8)

    _draw_nodes(ax, network, projected_nodes, issue_nodes, options)
    _draw_nozzles(ax, network, projected_nodes, options)
    _draw_input_labels(ax, network, projected_nodes)
    _draw_redline_items(ax, redline_items or [], options)

    if issue_list:
        ax.text(
            0.01,
            0.99,
            _issue_summary(issue_list),
            transform=ax.transAxes,
            va="top",
            fontsize=9,
            bbox={"facecolor": "white", "alpha": 0.75, "edgecolor": "0.8"},
        )

    ax.set_title(network.title)
    ax.set_aspect("equal", adjustable="datalim")
    ax.axis("off")
    fig.tight_layout()
    fig.savefig(output, format="png", dpi=150)
    plt.close(fig)


def _draw_nodes(
    ax: plt.Axes,
    network: PipeNetwork,
    projected_nodes: dict[str, tuple[float, float]],
    issue_nodes: set[str | None],
    options: IsometricRenderOptions,
) -> None:
    for node_id, node in network.nodes.items():
        projected = projected_nodes[node_id]
        is_issue = node_id in issue_nodes
        ax.scatter(
            [projected[0]],
            [projected[1]],
            s=36 if is_issue else 14,
            color="C3" if is_issue else "C0",
            zorder=3,
        )
        if options.show_node_labels:
            ax.text(projected[0], projected[1], node_id, fontsize=8)


def _draw_nozzles(
    ax: plt.Axes,
    network: PipeNetwork,
    projected_nodes: dict[str, tuple[float, float]],
    options: IsometricRenderOptions,
) -> None:
    for nozzle in network.nozzles.values():
        if nozzle.status != 1 or nozzle.input_node not in projected_nodes:
            continue
        projected = projected_nodes[nozzle.input_node]
        ax.scatter([projected[0]], [projected[1]], s=55, marker="x", color="C2", zorder=4)
        if options.show_nozzle_labels:
            ax.text(projected[0], projected[1], nozzle.nozzle_id, fontsize=8)


def _draw_input_labels(
    ax: plt.Axes,
    network: PipeNetwork,
    projected_nodes: dict[str, tuple[float, float]],
) -> None:
    for node in network.nodes.values():
        if _is_input_node(node) and node.node_id in projected_nodes:
            projected = projected_nodes[node.node_id]
            ax.text(projected[0], projected[1], f"Input {node.node_id}", fontsize=9, fontweight="bold")


def _draw_redline_items(
    ax: plt.Axes,
    redline_items: list[RedlineItem],
    options: IsometricRenderOptions,
) -> None:
    for item in redline_items:
        projected = project_iso(item.x, item.y, item.z, options.scale, options.z_scale)
        ax.scatter([projected[0]], [projected[1]], s=90, marker="o", facecolors="none", edgecolors="C3", zorder=5)
        ax.text(projected[0], projected[1], item.redline_id, fontsize=8, color="C3")


def _is_input_node(node: Node) -> bool:
    return node.node_id.upper() == "INPUT" or node.node_type.lower() == "input"


def _issue_summary(issues: list[ValidationIssue]) -> str:
    counts = Counter(issue.severity for issue in issues)
    parts = [f"{severity}: {count}" for severity, count in sorted(counts.items())]
    return "Validation issues\n" + "\n".join(parts)
