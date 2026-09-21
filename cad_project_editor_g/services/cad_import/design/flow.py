"""Edit-board adapter for the shared, source-preserving calculation tree."""
from __future__ import annotations

from typing import TYPE_CHECKING
from src.pipenet_converter.graph.flow import FlowTree, build_flow_tree

if TYPE_CHECKING:
    from services.cad_import.edit.board import EditBoard


def source_index(board: EditBoard, selected_source: str | None = None) -> int:
    """Resolve one explicit valve; never silently pick among multiple sources."""
    sources = list(board.sources)
    if selected_source is None or str(selected_source).strip() == "":
        if len(sources) == 1:
            return 0
        raise ValueError("알람밸브 기준을 하나 선택한 뒤 물흐름을 실행하세요.")
    wanted = str(selected_source).strip().upper()
    for i in range(len(sources)):
        if wanted in (f"Z{i + 1}", str(i + 1)):
            return i
    raise ValueError(f"급수원 '{selected_source}'를 찾지 못했습니다.")


def flow_for_board(board: EditBoard, *, selected_source: str | None = None,
                   index: int | None = None) -> FlowTree:
    """Reuse exact-input tree, independent of K, candidate area, and view zoom."""
    idx = source_index(board, selected_source) if index is None else int(index)
    if not 0 <= idx < len(board.sources):
        raise ValueError("급수원 번호가 범위를 벗어났습니다.")
    centers = getattr(board, "head_centers", ()) or ()
    # Completed head centers are authoritative. For unsupported/uncompleted
    # symbols retain the existing attachment candidates so the later probe can
    # report a conversion failure, rather than silently hiding that head.
    hnodes = [((centers[i],) if i < len(centers) and centers[i] is not None else tuple(ns))
              for i, ns in enumerate(board.hnodes)]
    lengths = getattr(board, "edge_len_mm", None) or {}
    normalized = {}
    for key, value in lengths.items():
        pair = key.split("|") if isinstance(key, str) else key
        a, b = (int(v) for v in pair)
        normalized[(min(a, b), max(a, b))] = float(value)
    inferred = tuple(sorted(getattr(board, "inferred_edges", ()) or ()))
    stamp = (idx, tuple(board.sources), tuple(tuple(p) for p in board.pts),
             tuple(sorted(tuple(sorted(e)) for e in board.edges)),
             tuple(tuple(sorted(ns)) for ns in hnodes),
             tuple(tuple(d) for d in board.disks), tuple(sorted(normalized.items())), inferred)
    cached = getattr(board, "_calculation_flow", None)
    if cached is not None and cached[0] == stamp:
        return cached[1]
    flow = build_flow_tree(board.pts, board.edges, hnodes, [board.sources[idx]],
                           head_xy=board.disks, lengths_mm=normalized, inferred_edges=inferred)
    board._calculation_flow = (stamp, flow)
    return flow


def water_state(board: EditBoard, *, index: int | None = None) -> dict:
    """Compatible animation state, restricted to the head-carrying tree."""
    tree = flow_for_board(board, index=index)
    active = {n for e in tree.edges for n in e} | set(tree.roots)
    return {"reach": active, "hop": {n: tree.hop[n] for n in active},
            "wet_edges": tree.edges, "wet_heads": set(tree.head_node),
            "total_heads": len(board.disks), "flow_report": tree.report()}


def stamp_planar_flow(built: dict, flow: FlowTree) -> None:
    """Map merged planar pipes to exact source paths; preserve mm→m lengths."""
    net = built["kfp"]
    refs = built.get("edge_ref") or {}
    for pid, pipe in net["pipe_data"].items():
        ref = refs.get(pid)
        path = flow.between(int(ref[0]), int(ref[1])) if ref else []
        if not path or not set(path) <= flow.edges:
            raise ValueError(f"배관 {pid}의 원본 물흐름 경로를 확인할 수 없습니다.")
        pipe["head_count_all"] = max(flow.loads[e] for e in path)
        pipe["length_m"] = sum(flow.lengths_mm[e] for e in path) / 1000.0
    net["flow_revision"] = flow.revision
    net["flow_report"] = flow.report()


def physical_pipe_loads(kfp: dict) -> dict[str, int]:
    """Propagate full source loads through generated risers/head connections.

    Original pipes have head_count_all from their exact source-tree path.
    Synthetic pipes carry their downstream physical loads, not selected-K loads.
    Fail on loops/disconnection instead of silently selecting another tree.
    """
    from services.cad_import.design.anchor import require_anchor
    nodes, pipes = kfp["nodes_meta_runtime"], kfp["pipe_data"]
    root = require_anchor(nodes, what="물흐름 담당 헤드 수")
    adj = {n: [] for n in nodes}
    for pid, p in pipes.items():
        a, b = p["start"], p["end"]
        adj[a].append((b, pid))
        adj[b].append((a, pid))
    order, parents = [root], {root: None}
    for n in order:
        for other, pid in adj[n]:
            if other not in parents:
                parents[other] = (n, pid)
                order.append(other)
    if len(order) != len(nodes) or len(pipes) != len(nodes) - 1:
        raise ValueError("전개 후 배관망이 단일 경로 트리가 아닙니다. 계산을 중단합니다.")
    counts = {n: int(nodes[n].get("type_id") == "head") for n in order}
    loads = {}
    for n in reversed(order[1:]):
        parent, pid = parents[n]
        load = max(counts[n], int(pipes[pid].get("head_count_all") or 0))
        loads[pid] = load
        counts[parent] += load
    return loads
