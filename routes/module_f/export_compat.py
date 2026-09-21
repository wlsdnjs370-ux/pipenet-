"""Module F boundary between calculation data and application display formats."""
from __future__ import annotations

from pathlib import Path
import sys
from typing import Collection


def _package_path() -> None:
    source = str(Path(__file__).resolve().parents[2] / "pipenet_converter" / "src")
    if source not in sys.path:
        sys.path.insert(0, source)


def prepare_sdf_export(path: str | Path, *,
                       nozzle_reference_labels: Collection[str] | None = None) -> list[str]:
    """Normalize PIPENET endpoints and return explicit export diagnostics."""
    _package_path()
    from pipenet_converter.sdf_compat import prepare_sdf
    report = prepare_sdf(path, nozzle_reference_labels=nozzle_reference_labels)
    messages = []
    if report.capped_nodes:
        messages.append("수요가 없는 배관 말단을 0유량 출구로 지정: " + ", ".join(report.capped_nodes))
    if report.unresolved_nozzles:
        messages.append("상·하향을 표고로 판정할 수 없는 노즐(기존 방향 유지): " + ", ".join(report.unresolved_nozzles))
    return messages


def emit_physical_kfp(sdf: str | Path, kfp: str | Path) -> Path:
    """Use unprojected SDF bearings, real elevations and declared pipe lengths.

    Never call this with a baked isometric SDF. Both view variants share this same
    physical KFP; the solver itself projects its XYZ coordinates for display.
    """
    _package_path()
    from kfp_sdf_converter import parse_sdf, emit_kfp
    from pipenet_converter.physical_layout import LayoutEdge, physical_layout
    net = parse_sdf(sdf)
    if any(p.waypoints for p in net.pipes.values()):
        raise ValueError("KFP 출력 전 꺾인 배관의 중간 절점을 분리하세요.")
    # A pump is a link in SDF but an inline node in KFP. Equal-elevation pump
    # endpoints form a zero-length equipment constraint, not a made-up pipe.
    pump_edges = []
    for key, node in net.nodes.items():
        outlet = node.raw.get("pump_output")
        if not outlet:
            continue
        end = net.nodes.get(outlet)
        if (end is None or end.kind != "base" or
                end.raw.get("io_node", "No") != "No" or
                abs(end.elevation_m - node.elevation_m) > 1e-6):
            raise ValueError(f"펌프 {key}: 토출 절점을 펌프 노드에 합칠 수 없습니다. 표고와 경계조건을 확인하세요.")
        pump_edges.append(LayoutEdge(f"pump:{key}", key, outlet, 0.0))
    positions = physical_layout(
        {key: (n.x, n.y, n.elevation_m) for key, n in net.nodes.items()},
        [LayoutEdge(p.id, p.start, p.end, p.length_m) for p in net.pipes.values()] + pump_edges,
        roots=[key for key, n in net.nodes.items() if n.raw.get("io_node") == "Input"])
    for key, (x, y, _z) in positions.items():
        net.nodes[key].x, net.nodes[key].y = x, y
    for edge in pump_edges:
        for pipe in net.pipes.values():
            if pipe.start == edge.end:
                pipe.start = edge.start
            if pipe.end == edge.end:
                pipe.end = edge.start
        del net.nodes[edge.end]
        net.nodes[edge.start].raw.pop("pump_output", None)
    # Keep every real pipe/corner; only zero-length pump endpoints fold into the
    # inline pump above. No straight-node simplification erases short fittings.
    emit_kfp(net, kfp, physical_coordinates=True)
    return Path(kfp)
