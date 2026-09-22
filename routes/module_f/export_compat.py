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
                       nozzle_reference_labels: Collection[str] | None = None,
                       display_spread: float = 1.0) -> list[str]:
    """Normalize PIPENET endpoints and return explicit export diagnostics.

    `display_spread` 는 좌표를 몇 배로 넓혀 쓴 .sdf 인가다(기본 1 = 종전 그대로).
    넓힌 만큼 노즐 꼬리 비율을 줄여 **꼬리 길이(표시 단위)는 넓히기 전과 같게**
    둔다 — 배관만 길어지고 기호·꼬리는 그대로라 노드가 작아 보인다.
    """
    _package_path()
    from pipenet_converter.sdf_compat import prepare_sdf, NOZZLE_DISPLAY_GAP_RATIO
    if not (display_spread > 0 and display_spread != float("inf")):
        raise ValueError("표시 좌표 배수는 0보다 큰 유한한 값이어야 합니다.")
    report = prepare_sdf(path, nozzle_reference_labels=nozzle_reference_labels,
                         nozzle_gap_ratio=NOZZLE_DISPLAY_GAP_RATIO / display_spread)
    messages = []
    if report.capped_nodes:
        messages.append("수요가 없는 배관 말단을 0유량 출구로 지정: " + ", ".join(report.capped_nodes))
    if report.unresolved_nozzles:
        messages.append("상·하향을 표고로 판정할 수 없는 노즐(기존 방향 유지): " + ", ".join(report.unresolved_nozzles))
    return messages


def use_remote_nozzle_export(path: str | Path) -> list[str]:
    """펌프 요소 없는 펌프 가압망 → PIPENET «가장 먼 헤드» 계산 방식.

    수원(맨 아래)에 대기압(0 g)을 고정해 두면 헤드까지 물을 올릴 힘이 없어
    PIPENET 이 헤드→수원 쪽으로 거꾸로 흐른다고 계산한다. 회사 수작업
    펌프가압구간 모델(`*_REMOTE.sdf`)처럼 급수원 압력을 비워 두고 «가장 먼
    헤드가 제 유량을 내게» 하면 PIPENET 이 급수원(펌프 토출)에 필요한 압력을
    구하고 물은 펌프 → 헤드로 흐른다. 계산 방식과 급수원 지정만 바뀐다.
    못 바꾸면 산출을 막지 않고 그 사유를 돌려준다(조용히 넘기지 않는다).
    """
    _package_path()
    from pipenet_converter.sdf_compat import use_most_remote_nozzle
    try:
        use_most_remote_nozzle(path)
    except ValueError as exc:
        return [f"{Path(path).name}: «가장 먼 헤드» 계산 방식으로 못 바꿨습니다 — {exc}"]
    return [f"{Path(path).name}: 펌프 제원이 없어 PIPENET «가장 먼 헤드»(Most Remote Nozzle)"
            " 방식으로 저장 — 급수원(펌프 토출) 압력은 PIPENET 이 계산합니다."]


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
