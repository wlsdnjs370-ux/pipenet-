# -*- coding: utf-8 -*-
"""손질 입출력 — 자동망 오픈·유저손질 저장. 화면 없음.

캐시 스탬프는 disp_cache 가 잰다. 찍기 JSON·disp_cache 는 덮지 않는다.
"""
import json
import os

from services.cad_import.edit.board import EditBoard
from services.cad_import.kinds import require_head_kinds
from services.cad_import.pipeline.disp_cache import (
    _disp_cache_load, _disp_cache_save,
)
from services.cad_import.pipeline.expand import stage1_body
from services.cad_import.pipeline.flow import ho_from_spots, pipeline
from services.cad_import.pipeline.handoff import default_edits_dir
from services.cad_import.pipeline.user_net import (
    SNAP, apply_deletes, apply_head_bridge, apply_joins, apply_kind_overrides,
    apply_tee_bridge, insert_node_on_pipe, user_edits_path,
)


def _board_from_data(key, data):
    return EditBoard(
        key, data["pts"], {tuple(e) for e in data["edges"]},
        data["hcov"],
        original_edges={tuple(e) for e in data["edges1"]},
        hnodes=data.get("hnodes"),
        ups=data.get("ups") or (),
        head_kinds=require_head_kinds(
            data["hcov"], data.get("head_kinds") or ()),
        ho=data.get("ho") or ())


def open_board(key, use_cache=True):
    """disp_cache HIT 이면 그 망, MISS 이면 제품 pipeline 후 캐시 저장.

    use_cache=False 는 찍기 다음 — 방금 쓴 스펙으로 1~6단계를 다시 돌린다.
    """
    if use_cache:
        cached = _disp_cache_load(key)
        if cached is not None:
            return _board_from_data(key, cached)
    st = stage1_body(key)
    result = pipeline(st)
    _disp_cache_save(key, result)
    data = {
        "pts": result["pts"],
        "edges": result["edges"],
        "edges1": result["edges1"],
        "hcov": result["hcov"],
        "hnodes": result["hnodes"],
        "ups": result.get("ups") or (),
        "head_kinds": result.get("head_kinds") or (),
        "ho": ho_from_spots(result.get("spots")),
    }
    return _board_from_data(key, data)


def write_edits(board, out_dir=None):
    """payload 를 *_유저손질.json 에 쓴다. 찍기·disp_cache 경로는 안 쓴다."""
    path = user_edits_path(board.key, out_dir or default_edits_dir())
    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(board.payload(), f, ensure_ascii=False, indent=2)
    return path


def load_edits(board, out_dir=None):
    """기존 *_유저손질.json 을 보드에 적용. 없으면 False."""
    path = user_edits_path(board.key, out_dir or default_edits_dir())
    if not os.path.exists(path):
        return False
    raw = json.load(open(path, encoding="utf-8"))
    board._invalidate_source_cache()
    for rec in raw.get("joins") or []:
        bridges = [(tuple(x[0]), tuple(x[1])) for x in rec.get("bridges") or []]
        if (rec.get("kind") == "헤드" or rec.get("head")) and bridges:
            for br in bridges:
                board.pts, board.edges = apply_head_bridge(
                    board.pts, board.edges, br)
        elif rec.get("kind") == "T자" and bridges:
            trunks = rec.get("trunks") or []
            for i, br in enumerate(bridges):
                if i < len(trunks):
                    tr = trunks[i]
                    board.pts, board.edges = apply_tee_bridge(
                        board.pts, board.edges, br[0], br[1],
                        (tuple(tr[0]), tuple(tr[1])))
                else:
                    board.pts, board.edges = apply_joins(
                        board.pts, board.edges, [br])
        else:
            board.pts, board.edges = apply_joins(board.pts, board.edges, bridges)
        board.joins.append(rec)
    board.deletes = list(raw.get("deletes") or [])
    board.edges = apply_deletes(board.pts, board.edges, board.deletes)
    board._complete_heads()
    for rec in raw.get("valve_picks") or []:
        xy = rec.get("xy") if isinstance(rec, dict) else rec
        if not xy or len(xy) < 2:
            continue
        board.pts, board.edges, node = insert_node_on_pipe(
            board.pts, board.edges, float(xy[0]), float(xy[1]), max_d=SNAP)
        if node is not None and node not in board.valves:
            board.valves.append(node)
    source_recs = raw.get("sources") or []
    hnodes = board._head_nodes() if source_recs else None
    source_nodes = board._source_nodes(hnodes) if source_recs else None
    for rec in source_recs:
        xy = rec.get("xy") or ()
        if len(xy) == 2:
            node = board.snap_source(
                *xy, hnodes=hnodes, source_nodes=source_nodes)
            if node is not None and node not in board.sources:
                board.sources.append(node)
    ovs = list(raw.get("kind_overrides") or [])
    if ovs:
        board.kind_overrides = [dict(r) for r in ovs if isinstance(r, dict)]
        board.head_kinds = apply_kind_overrides(
            board.head_kinds, board.kind_overrides)
    board.head_kinds = require_head_kinds(board.disks, board.head_kinds)
    board._refresh_kind_views()
    return True
