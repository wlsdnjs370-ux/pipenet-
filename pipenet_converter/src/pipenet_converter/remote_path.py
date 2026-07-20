"""Remote-N head extraction for hydraulically most-remote area analysis.

Given a calculation-ready ``PipeNetwork`` and the node ID of the alarm valve,
this module:

1. Finds the head that is hydraulically farthest from the alarm valve, measured
   by pipe centerline length along the connected pipe graph.
2. Selects the *neighbouring* ``N`` heads around that farthest head, ranked by
   pipe-network distance to the farthest head. This keeps the selection inside
   the same branch / residence and avoids picking isolated outliers that just
   happen to be far from the valve.
3. Extracts the subtree that connects the alarm valve to each selected head
   (the union of the shortest pipe paths) while preserving the original CAD
   coordinates. No re-layout, no isometric projection — the shape on the PNG
   matches the plan drawing with the architectural background stripped away.
4. Renumbers the selected heads ``H01..HN`` (farthest head = ``H01``), the
   intermediate junctions ``J01..``, and the kept pipes ``P001..``.
5. Writes the remote-area isometric PNG and a Pipe Schedule + Remote heads
   workbook (xlsx).

The default ``N`` is 30, matching the workflow in the reference ChatGPT
session.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from math import hypot
from pathlib import Path
from typing import Iterable

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import networkx as nx
import pandas as pd

from pipenet_converter.models import (
    Node,
    Nozzle,
    Pipe,
    PipeNetwork,
    diameter_m_to_label,
)


DEFAULT_REMOTE_HEAD_COUNT = 30


@dataclass(slots=True)
class RemotePathOptions:
    """Options for selecting and exporting the hydraulically most-remote area."""

    head_count: int = DEFAULT_REMOTE_HEAD_COUNT
    head_node_types: tuple[str, ...] = ("head",)
    figsize: tuple[float, float] = (12, 9)
    show_diameter_labels: bool = True
    show_node_labels: bool = True
    head_prefix: str = "H"
    junction_prefix: str = "J"
    pipe_prefix: str = "P"


@dataclass(slots=True)
class RemotePathResult:
    """Outputs of remote-path selection.

    ``network`` is the renumbered sub-network. The maps preserve the link back
    to the source ``PipeNetwork`` so callers can reconcile the selection with
    the original CAD entities.
    """

    network: PipeNetwork
    alarm_valve_node_id: str
    alarm_valve_renamed_id: str
    farthest_head_original_id: str
    farthest_head_renamed_id: str
    head_order: list[str] = field(default_factory=list)
    head_renumber_map: dict[str, str] = field(default_factory=dict)
    junction_renumber_map: dict[str, str] = field(default_factory=dict)
    pipe_renumber_map: dict[str, str] = field(default_factory=dict)
    total_pipe_length_m: float = 0.0
    distance_from_valve_m: dict[str, float] = field(default_factory=dict)


@dataclass(slots=True)
class RemotePathArtifacts:
    """Filesystem paths of artifacts written for a remote selection."""

    isometric_png: Path
    pipe_schedule_xlsx: Path


def extract_remote_path(
    network: PipeNetwork,
    alarm_valve_node_id: str,
    options: RemotePathOptions | None = None,
) -> RemotePathResult:
    """Select the farthest head and the connected head cluster around it.

    The returned sub-network keeps the original CAD coordinates so the PNG
    rendering reproduces the plan-drawing geometry — only the architectural
    background is dropped.
    """
    opts = options or RemotePathOptions()
    if opts.head_count <= 0:
        raise ValueError(f"head_count must be positive, got {opts.head_count}.")
    if alarm_valve_node_id not in network.nodes:
        raise KeyError(
            f"alarm_valve_node_id {alarm_valve_node_id!r} is not a node in the network."
        )

    pipe_graph = _build_pipe_graph(network)
    if alarm_valve_node_id not in pipe_graph:
        raise ValueError(
            f"alarm_valve_node_id {alarm_valve_node_id!r} is not connected to any pipe."
        )

    head_nodes = _collect_head_nodes(network, opts.head_node_types)
    if not head_nodes:
        raise ValueError("PipeNetwork has no head nodes to analyse.")

    valve_distances = nx.single_source_dijkstra_path_length(
        pipe_graph, alarm_valve_node_id, weight="length_m"
    )
    reachable_heads = [hid for hid in head_nodes if hid in valve_distances]
    if not reachable_heads:
        raise ValueError(
            f"No head node is reachable from alarm valve {alarm_valve_node_id!r} via pipes."
        )

    farthest_head = max(reachable_heads, key=lambda hid: valve_distances[hid])

    farthest_distances = nx.single_source_dijkstra_path_length(
        pipe_graph, farthest_head, weight="length_m"
    )
    candidates = [hid for hid in reachable_heads if hid in farthest_distances]
    candidates.sort(key=lambda hid: (farthest_distances[hid], -valve_distances[hid]))
    selected_heads = candidates[: opts.head_count]

    keep_nodes, keep_edges = _shortest_path_union(
        pipe_graph, alarm_valve_node_id, selected_heads
    )

    sub_network, head_renumber_map, junction_renumber_map, pipe_renumber_map = (
        _build_sub_network(
            source=network,
            keep_nodes=keep_nodes,
            keep_edges=keep_edges,
            selected_heads=selected_heads,
            alarm_valve_node_id=alarm_valve_node_id,
            farthest_head=farthest_head,
            valve_distances=valve_distances,
            options=opts,
        )
    )

    head_order = sorted(selected_heads, key=lambda hid: -valve_distances[hid])
    total_pipe_length_m = sum(pipe.length_m for pipe in sub_network.pipes.values())

    return RemotePathResult(
        network=sub_network,
        alarm_valve_node_id=alarm_valve_node_id,
        alarm_valve_renamed_id=junction_renumber_map.get(
            alarm_valve_node_id, alarm_valve_node_id
        ),
        farthest_head_original_id=farthest_head,
        farthest_head_renamed_id=head_renumber_map[farthest_head],
        head_order=head_order,
        head_renumber_map=head_renumber_map,
        junction_renumber_map=junction_renumber_map,
        pipe_renumber_map=pipe_renumber_map,
        total_pipe_length_m=total_pipe_length_m,
        distance_from_valve_m={hid: valve_distances[hid] for hid in selected_heads},
    )


def write_remote_path_artifacts(
    result: RemotePathResult,
    output_dir: str | Path,
    options: RemotePathOptions | None = None,
    png_name: str = "remote_isometric.png",
    xlsx_name: str = "remote_pipe_schedule.xlsx",
) -> RemotePathArtifacts:
    """Write the remote-area isometric PNG and Pipe Schedule xlsx to ``output_dir``."""
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    png_path = out / png_name
    xlsx_path = out / xlsx_name
    render_remote_path_png(result, png_path, options)
    write_remote_path_xlsx(result, xlsx_path)
    return RemotePathArtifacts(isometric_png=png_path, pipe_schedule_xlsx=xlsx_path)


def render_remote_path_png(
    result: RemotePathResult,
    output_path: str | Path,
    options: RemotePathOptions | None = None,
) -> None:
    """Render the remote-area PNG using the original CAD coordinates."""
    opts = options or RemotePathOptions()
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    network = result.network
    fig, ax = plt.subplots(figsize=opts.figsize)

    for pipe in network.pipes.values():
        a = network.nodes[pipe.from_node]
        b = network.nodes[pipe.to_node]
        ax.plot([a.x, b.x], [a.y, b.y], color="#1f4e8f", linewidth=1.8, zorder=2)
        if opts.show_diameter_labels:
            mid_x, mid_y = (a.x + b.x) / 2.0, (a.y + b.y) / 2.0
            ax.text(
                mid_x,
                mid_y,
                _safe_diameter_label(pipe.diameter_m),
                fontsize=7,
                color="#555555",
                ha="center",
                va="center",
            )

    renamed_heads = set(result.head_renumber_map.values())
    renamed_valve = result.alarm_valve_renamed_id

    for node_id, node in network.nodes.items():
        if node_id == renamed_valve:
            continue
        if node_id in renamed_heads:
            ax.scatter([node.x], [node.y], s=55, color="#d62728", marker="o", zorder=3)
            if opts.show_node_labels:
                ax.annotate(
                    node_id,
                    (node.x, node.y),
                    textcoords="offset points",
                    xytext=(6, 6),
                    fontsize=8,
                    color="#d62728",
                )
        else:
            ax.scatter([node.x], [node.y], s=20, color="#444444", marker="s", zorder=3)
            if opts.show_node_labels:
                ax.annotate(
                    node_id,
                    (node.x, node.y),
                    textcoords="offset points",
                    xytext=(4, -10),
                    fontsize=7,
                    color="#444444",
                )

    valve_node = network.nodes[renamed_valve]
    ax.scatter([valve_node.x], [valve_node.y], s=170, marker="*", color="#2ca02c", zorder=4)
    ax.annotate(
        "ALARM VALVE",
        (valve_node.x, valve_node.y),
        textcoords="offset points",
        xytext=(10, 10),
        fontsize=9,
        color="#2ca02c",
        fontweight="bold",
    )

    ax.set_title(
        f"Remote {len(result.head_order)} Heads — {network.title} "
        f"(farthest: {result.farthest_head_renamed_id})"
    )
    ax.set_aspect("equal", adjustable="datalim")
    ax.axis("off")
    fig.tight_layout()
    fig.savefig(out, format="png", dpi=180)
    plt.close(fig)


def write_remote_path_xlsx(
    result: RemotePathResult,
    output_path: str | Path,
) -> None:
    """Write the Pipe Schedule and Remote heads tables as an xlsx workbook."""
    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    network = result.network

    pipe_rows = [
        {
            "pipe_id": pipe.pipe_id,
            "from_node": pipe.from_node,
            "to_node": pipe.to_node,
            "diameter_label": _safe_diameter_label(pipe.diameter_m),
            "diameter_m": pipe.diameter_m,
            "length_m": round(pipe.length_m, 3),
        }
        for pipe in network.pipes.values()
    ]
    pipe_df = pd.DataFrame(
        pipe_rows,
        columns=[
            "pipe_id",
            "from_node",
            "to_node",
            "diameter_label",
            "diameter_m",
            "length_m",
        ],
    )

    head_rows = []
    for original_id in result.head_order:
        renamed = result.head_renumber_map[original_id]
        node = network.nodes[renamed]
        head_rows.append(
            {
                "head_no": renamed,
                "original_node_id": original_id,
                "x": node.x,
                "y": node.y,
                "distance_from_valve_m": round(
                    result.distance_from_valve_m.get(original_id, 0.0), 3
                ),
            }
        )
    head_df = pd.DataFrame(
        head_rows,
        columns=["head_no", "original_node_id", "x", "y", "distance_from_valve_m"],
    )

    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pipe_df.to_excel(writer, sheet_name="Pipe Schedule", index=False)
        head_df.to_excel(writer, sheet_name="Remote Heads", index=False)


def find_nearest_node(
    network: PipeNetwork,
    point: tuple[float, float],
    node_types: Iterable[str] | None = None,
) -> str:
    """Return the node ID closest to ``point`` in the network's XY plane.

    Convenience helper for callers that know the alarm valve location in CAD
    coordinates rather than as a node ID.
    """
    allowed = set(node_types) if node_types else None
    best_id: str | None = None
    best_distance = float("inf")
    for node_id, node in network.nodes.items():
        if allowed is not None and node.node_type not in allowed:
            continue
        distance = hypot(node.x - point[0], node.y - point[1])
        if distance < best_distance:
            best_id = node_id
            best_distance = distance
    if best_id is None:
        raise ValueError("PipeNetwork has no nodes (or no nodes matching node_types).")
    return best_id


def _build_pipe_graph(network: PipeNetwork) -> nx.Graph:
    graph: nx.Graph = nx.Graph()
    for node_id, node in network.nodes.items():
        graph.add_node(node_id, x=node.x, y=node.y, node_type=node.node_type)
    for pipe in network.pipes.values():
        if pipe.from_node not in graph or pipe.to_node not in graph:
            continue
        if graph.has_edge(pipe.from_node, pipe.to_node):
            existing = graph[pipe.from_node][pipe.to_node]
            if pipe.length_m < existing["length_m"]:
                existing.update(pipe_id=pipe.pipe_id, length_m=pipe.length_m)
            continue
        graph.add_edge(
            pipe.from_node,
            pipe.to_node,
            pipe_id=pipe.pipe_id,
            length_m=pipe.length_m,
        )
    return graph


def _collect_head_nodes(
    network: PipeNetwork,
    head_node_types: tuple[str, ...],
) -> list[str]:
    by_nozzle = sorted(
        {
            nozzle.input_node
            for nozzle in network.nozzles.values()
            if nozzle.input_node in network.nodes
        }
    )
    if by_nozzle:
        return by_nozzle
    return sorted(
        node_id
        for node_id, node in network.nodes.items()
        if node.node_type in head_node_types
    )


def _shortest_path_union(
    pipe_graph: nx.Graph,
    alarm_valve_node_id: str,
    selected_heads: list[str],
) -> tuple[set[str], set[frozenset[str]]]:
    keep_nodes: set[str] = {alarm_valve_node_id}
    keep_edges: set[frozenset[str]] = set()
    for head_id in selected_heads:
        path = nx.shortest_path(
            pipe_graph, alarm_valve_node_id, head_id, weight="length_m"
        )
        keep_nodes.update(path)
        for u, v in zip(path, path[1:]):
            keep_edges.add(frozenset((u, v)))
    return keep_nodes, keep_edges


def _build_sub_network(
    source: PipeNetwork,
    keep_nodes: set[str],
    keep_edges: set[frozenset[str]],
    selected_heads: list[str],
    alarm_valve_node_id: str,
    farthest_head: str,
    valve_distances: dict[str, float],
    options: RemotePathOptions,
) -> tuple[PipeNetwork, dict[str, str], dict[str, str], dict[str, str]]:
    head_set = set(selected_heads)
    head_order = sorted(selected_heads, key=lambda hid: -valve_distances[hid])
    head_renumber_map = {
        original: _format_id(options.head_prefix, index + 1, width=2)
        for index, original in enumerate(head_order)
    }

    junction_ids = [
        node_id
        for node_id in sorted(keep_nodes, key=lambda nid: valve_distances.get(nid, 0.0))
        if node_id not in head_set and node_id != alarm_valve_node_id
    ]
    junction_renumber_map = {
        alarm_valve_node_id: "ALARM_VALVE",
        **{
            original: _format_id(options.junction_prefix, index + 1, width=2)
            for index, original in enumerate(junction_ids)
        },
    }

    node_renumber_map: dict[str, str] = {}
    node_renumber_map.update(head_renumber_map)
    node_renumber_map.update(junction_renumber_map)

    sub = PipeNetwork(title=f"{source.title} — Remote{len(selected_heads)}")

    for original_id in keep_nodes:
        if original_id not in source.nodes:
            continue
        node = source.nodes[original_id]
        new_id = node_renumber_map.get(original_id, original_id)
        sub.add_node(
            Node(
                node_id=new_id,
                x=node.x,
                y=node.y,
                z=node.z,
                node_type=_classify_node_type(
                    original_id, head_set, alarm_valve_node_id, farthest_head, node.node_type
                ),
                source=node.source,
                metadata={
                    **node.metadata,
                    "original_node_id": original_id,
                    "is_alarm_valve": original_id == alarm_valve_node_id,
                    "is_remote_head": original_id in head_set,
                    "is_farthest_head": original_id == farthest_head,
                    "distance_from_valve_m": float(
                        valve_distances.get(original_id, 0.0)
                    ),
                },
            )
        )

    pipe_renumber_map: dict[str, str] = {}
    pipe_counter = 1
    for pipe in source.pipes.values():
        if frozenset((pipe.from_node, pipe.to_node)) not in keep_edges:
            continue
        if pipe.from_node not in node_renumber_map or pipe.to_node not in node_renumber_map:
            continue
        new_id = _format_id(options.pipe_prefix, pipe_counter, width=3)
        pipe_renumber_map[pipe.pipe_id] = new_id
        sub.add_pipe(
            Pipe(
                pipe_id=new_id,
                from_node=node_renumber_map[pipe.from_node],
                to_node=node_renumber_map[pipe.to_node],
                diameter_m=pipe.diameter_m,
                length_m=pipe.length_m,
                rise_m=pipe.rise_m,
                c_factor=pipe.c_factor,
                material=pipe.material,
                status=pipe.status,
                fittings=list(pipe.fittings),
                equipment=list(pipe.equipment),
                waypoints=list(pipe.waypoints),
                metadata={**pipe.metadata, "original_pipe_id": pipe.pipe_id},
            )
        )
        pipe_counter += 1

    for nozzle in source.nozzles.values():
        if nozzle.input_node not in head_set:
            continue
        renamed_input = node_renumber_map[nozzle.input_node]
        renamed_output = f"@/{nozzle.nozzle_id}"
        if nozzle.output_node in source.nodes and renamed_output not in sub.nodes:
            output_node = source.nodes[nozzle.output_node]
            sub.add_node(
                Node(
                    node_id=renamed_output,
                    x=output_node.x,
                    y=output_node.y,
                    z=output_node.z,
                    node_type=output_node.node_type or "head_output",
                    source=output_node.source,
                    metadata={**output_node.metadata, "original_node_id": nozzle.output_node},
                )
            )
        sub.add_nozzle(
            Nozzle(
                nozzle_id=nozzle.nozzle_id,
                input_node=renamed_input,
                output_node=renamed_output,
                flow_m3s=nozzle.flow_m3s,
                status=nozzle.status,
                library_item=nozzle.library_item,
                metadata={**nozzle.metadata, "original_input_node": nozzle.input_node},
            )
        )

    return sub, head_renumber_map, junction_renumber_map, pipe_renumber_map


def _classify_node_type(
    original_id: str,
    head_set: set[str],
    alarm_valve_node_id: str,
    farthest_head: str,
    fallback: str,
) -> str:
    if original_id == alarm_valve_node_id:
        return "alarm_valve"
    if original_id == farthest_head:
        return "farthest_head"
    if original_id in head_set:
        return "head"
    return fallback or "junction"


def _format_id(prefix: str, index: int, width: int) -> str:
    return f"{prefix}{index:0{width}d}"


def _safe_diameter_label(diameter_m: float) -> str:
    try:
        return diameter_m_to_label(diameter_m)
    except ValueError:
        return f"{diameter_m * 1000:.0f}mm"
