"""Opt-in H all-area selection over F's existing graph/expansion engine."""
from __future__ import annotations

from typing import TYPE_CHECKING

from src.pipenet_converter.graph.area_selection import area_corridor, select_area_heads
from src.pipenet_converter.graph.area_network import AREA_SCOPE, select_area_network
from src.pipenet_converter.graph.regions import zones_signature

if TYPE_CHECKING:
    from services.cad_import.edit.board import EditBoard
    from services.cad_import.design.tables import PipeTablesG


def validate_area_selection(sess: dict, requested_zones: list | None = None) -> None:
    """Refuse missing/stale area selections instead of falling back to K ranking."""
    from services.cad_import.design.flow import flow_for_board, network_for_board, source_index
    w = sess.get("worst") or {}
    if w.get("selection_mode") != "area_all" or not w.get("heads"):
        raise ValueError("영역을 지정하고 「영역 배관망 추출」을 먼저 실행하세요.")
    b = sess["edit"].board
    if requested_zones is not None and zones_signature(w["zones"]) != zones_signature(requested_zones):
        raise ValueError("화면의 영역이 추출 결과와 다릅니다. 영역 배관망을 다시 추출하세요.")
    if zones_signature(w["zones"]) != zones_signature(getattr(b, "selection_zones", [])):
        raise ValueError("영역이 변경되었습니다. 영역 배관망을 다시 추출하세요.")
    index = source_index(b, w.get("source_tag"))
    network = network_for_board(b, index=index) if getattr(b, "network_mode", "tree") != "tree" else None
    flow = network.reference if network else flow_for_board(b, index=index)
    selection = select_area_heads(b.disks, w["zones"], flow)
    if (selection.disconnected or list(selection.heads) != w["heads"]
            or list(selection.candidates) != w.get("area_candidates")
            or w.get("flow_revision") != (network.revision if network else flow.revision)):
        raise ValueError("영역 안 헤드 또는 연결이 변경되었습니다. 영역 배관망을 다시 추출하세요.")
    if network:
        scoped = select_area_network(network, b.pts, selection.heads, w["zones"])
        if ((w.get("area_scope") or {}).get("policy") != AREA_SCOPE
                or set(w.get("edges", ())) != scoped.edges):
            raise ValueError("영역 추출 범위가 변경되었습니다. 영역 배관망을 다시 추출하세요.")


def validate_area_nozzles(board: EditBoard, worst: dict, tables: PipeTablesG) -> None:
    """Reject loss during vertical expansion/overrides; combo symbols have two heads."""
    expected = sum(2 if board.disk_kinds[h] == "상하향식" else 1 for h in worst["heads"])
    if len(tables.nozzles) != expected:
        raise ValueError(f"영역 헤드와 입력표 노즐 수가 다릅니다 ({expected} → {len(tables.nozzles)}). "
                         "누락·추가된 헤드와 편집 기록을 확인한 뒤 다시 추출하세요.")


def compute_area_selection(sess: dict, body: dict) -> tuple[dict | None, dict | None]:
    """Validate every region head and publish one explicit selection; ignore K/sheet."""
    from routes.module_f.api_edit import _wfail
    from routes.module_f.attach import picked_heads_wet
    from services.cad_import.design.flow import flow_for_board, network_for_board, source_index
    es, b = sess["edit"], sess["edit"].board
    # No stale successful selection remains usable after a rejected new request.
    sess["selection_mode"] = "area_all"
    sess["worst"] = sess["worst_edit"] = None
    try:
        index = source_index(b, body.get("source"))
        network = network_for_board(b, index=index) if getattr(b, "network_mode", "tree") != "tree" else None
        flow = network.reference if network else flow_for_board(b, index=index)
        selected = select_area_heads(b.disks, body.get("zones"), flow)
        blocked = [(h, "급수원 미연결") for h in selected.disconnected]
        blocked.extend((h, "헤드 종류 미지정") for h in selected.heads
                       if h >= len(b.disk_kinds) or b.disk_kinds[h] == "미지정")
        blocked.extend((h, "겹친 헤드 기호의 종류 불일치") for h, representative in selected.aliases
                       if b.disk_kinds[h] != b.disk_kinds[representative])
        if not blocked and not network:
            probe = picked_heads_wet(es, list(selected.heads), selected_source=f"Z{index+1}")
            if not probe.get("ok"):
                raise ValueError(probe.get("error") or "영역 헤드의 배관 접속을 검증하지 못했습니다.")
            wet = set(probe["wet"]) - set(probe.get("shared") or ())
            blocked.extend((h, str((probe.get("reason") or {}).get(h) or "접속 확인 필요"))
                           for h in selected.heads if h not in wet)
        if blocked:
            return None, _wfail(f"영역 안 헤드 {len(blocked)}개를 연결·정의하지 못했습니다. 누락 없이 추출하려면 해당 위치를 수정하세요.",
                extra={"not_attached": {"n": len(blocked), "items": [
                    {"disk": h, "xy": list(b.disks[h][:2]), "why": reason,
                     "todo": "배관 연결과 헤드 종류를 확인하세요."} for h, reason in blocked]}})
        w = area_corridor(selected, b.disks, flow)
        if network:
            scoped = select_area_network(network, b.pts, selected.heads, selected.zones)
            w.update(edges=set(scoped.edges), area_scope=scoped.report(network), loads={}, physical_loads={},
                     max_load=None, network_mode=network.mode, flow_revision=network.revision,
                     flow_report=network.report(), attachment_validation="source_graph_only")
            w["nodes"] = {n for edge in w["edges"] for n in edge} | set(flow.roots)
            w["total_m"] = round(sum(flow.lengths_mm[e] for e in w["edges"])/1000, 2)
            # The full-network reference path may use an excluded exterior
            # detour. Rebuild diagnostic distances on the actual scoped graph.
            from src.pipenet_converter.graph.flow import build_flow_tree
            scoped_flow = build_flow_tree(b.pts, scoped.edges,
                [{flow.head_node[h]} if h in flow.head_node else set() for h in range(len(b.disks))],
                flow.roots, head_xy=b.disks, lengths_mm=flow.lengths_mm,
                inferred_edges=network.inferred)
            diagnostics = area_corridor(selected, b.disks, scoped_flow)
            for key in ("worst_head", "worst_path", "dists", "worst_path_m", "far_m", "near_m"):
                w[key] = diagnostics[key]
        state = es.flow(source_index=index)
        while es.flow_tick():
            pass
        sess["flow_report"] = state["flow_report"]
        sess["water_path"] = [[*b.pts[a][:2], *b.pts[z][:2]] for a, z in sorted(state["wet_edges"])]
        w.update(sheet=None, source_tag=f"Z{index+1}", source_index=index)
        b.selection_zones = w["zones"]
        sess.update(worst=w, worst_cand=list(selected.candidates), worst_zones=w["zones"],
                    worst_k=len(w["heads"]), worst_edits=0,
                    worst_args={"selection_mode": "area_all", "source": w["source_tag"], "zones": w["zones"]})
        # k is now a derived count for legacy metadata consumers, never an input.
        if sess.get("design_settings"):
            sess["design_settings"] = {**sess["design_settings"], "k": len(w["heads"]),
                                       "selection_mode": "area_all", "fill_short": False}
        summary = {key: w[key] for key in ("selection_mode", "candidates", "reachable", "far_m", "near_m",
                   "span_m", "area_w_m", "area_h_m", "area_m2", "total_m", "max_load", "merged", "merged_xy", "worst_path_m")}
        summary.update(k=len(w["heads"]), zones=len(w["zones"]), source=w["source_tag"], sheet=None,
                       network_mode=getattr(b, "network_mode", "tree"), hydraulically_verified=False,
                       path_edges=len(w["edges"]), rank_invariant=None)
        if "area_scope" in w:
            summary["area_scope"] = w["area_scope"]
        return summary, None
    except ValueError as exc:
        return None, _wfail(str(exc))
