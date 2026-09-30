"""Non-mutating loop/grid source projection; the legacy tree stays untouched."""
from __future__ import annotations

import copy
import math
from typing import TYPE_CHECKING

from src.pipenet_converter.graph.cycle_closures import (
    ClosureCandidate, ClosureResult, recover_cycle_closures,
)
from src.pipenet_converter.graph.head_junctions import HeadJunction, preserve_head_junctions
from src.pipenet_converter.graph.junction_cleanup import duplicate_symbol_joins

if TYPE_CHECKING:
    from services.cad_import.edit.board import EditBoard


def cycle_source(board: EditBoard) -> ClosureResult:
    """Revalidate cached source evidence against current edits and deletions."""
    rows = getattr(board, "cycle_candidates", ()) or ()
    candidates = tuple(ClosureCandidate.from_record(row) for row in rows)
    junctions = tuple(HeadJunction.from_record(row) for row in
                      (getattr(board, "head_junctions", ()) or ()))
    centers = tuple(n for n in getattr(board, "head_centers", ()) if n is not None)
    forbidden = []
    for raw in getattr(board, "deletes", ()):
        seg = (raw["a"], raw["b"]) if isinstance(raw, dict) else raw
        if all(isinstance(n, int) for n in seg):
            if all(0 <= n < len(board.pts) for n in seg):
                seg = [board.pts[n] for n in seg]
            else:
                continue
        forbidden.append(tuple(tuple(xy) for xy in seg))
    manual = set()
    xy_node = {tuple(p[:2]):i for i,p in enumerate(board.pts)}
    for rec in getattr(board,'joins',()) or ():
        for a,b in rec.get('bridges',()):
            if tuple(a[:2]) in xy_node and tuple(b[:2]) in xy_node:
                manual.add(tuple(sorted((xy_node[tuple(a[:2])],xy_node[tuple(b[:2])]))))
    for key in (getattr(board,'edge_len_mm',None) or {}):
        pair = key.split('|') if isinstance(key,str) else key
        manual.add(tuple(sorted(int(n) for n in pair)))
    arcs = getattr(board,'ho',()) or ()
    stamp = (tuple(tuple(p) for p in board.pts), board.edges,
             candidates, junctions, centers, tuple(forbidden),str(arcs),frozenset(manual))
    cached = getattr(board, "_cycle_source", None)
    if cached is None or cached[0] != stamp:
        duplicate = duplicate_symbol_joins(board.pts,board.edges,board.original_edges,arcs,
                                          protected_edges=manual)
        result = recover_cycle_closures(board.pts, set(board.edges)-set(duplicate), candidates,
                                       head_centers=centers, forbidden_segments=forbidden)
        if junctions:
            result = preserve_head_junctions(board.pts, result, junctions,
                                             head_centers=centers, forbidden_segments=forbidden)
        if duplicate:
            result = ClosureResult(result.edges,result.added,
                result.records+tuple(dict(row,status='duplicate_removed') for row in duplicate.values()),
                result.removed|frozenset(duplicate))
        board._cycle_source = stamp, result
    return board._cycle_source[1]


def projected_board(board: EditBoard) -> EditBoard:
    """Read-only shallow view: only derived edges and caches differ from source."""
    result = cycle_source(board)
    projected = copy.copy(board)
    projected.edges = result.edges
    projected.inferred_edges = (set(getattr(board, "inferred_edges", ())) - set(result.removed)) | set(result.added)
    # A drawn run replaced by two center-connected spans keeps its declared
    # total length. Never fall back to display distance when a length exists.
    lengths = getattr(board, "edge_len_mm", None) or {}
    if lengths and result.removed:
        normalized = {}
        for key, value in lengths.items():
            pair = key.split('|') if isinstance(key, str) else key
            normalized[tuple(sorted(int(n) for n in pair))] = float(value)
        for row in result.records:
            center = row.get('center')
            if center is None:
                continue
            for a, b in row.get('split_edges', ()):
                length = normalized.pop(tuple(sorted((a, b))), None)
                if length is None:
                    continue
                la = math.dist(board.pts[a][:2], board.pts[center][:2])
                lb = math.dist(board.pts[b][:2], board.pts[center][:2])
                for n, portion in ((a, la), (b, lb)):
                    edge = tuple(sorted((n, center)))
                    if edge in result.edges:
                        normalized[edge] = length * portion / (la + lb)
        projected.edge_len_mm = normalized
    projected._calculation_flow = None
    return projected
