"""Boundary between preserved extraction and the legacy calculation pipeline."""
from __future__ import annotations

from src.pipenet_converter.graph.network import network_mode


def calculation_block(sess: dict, *, all_slots: bool = False) -> str | None:
    """Prevent stale tree outputs from being presented as loop/grid outputs."""
    snapshots = [sess]
    if all_slots:
        snapshots.extend((sess.get("slots") or {}).values())
    for snapshot in snapshots:
        board = getattr(snapshot.get("edit"), "board", None)
        mode = network_mode(getattr(board, "network_mode", "tree"))
        if mode != "tree" and ((snapshot.get('design') or {}).get('got') or {}).get('network_mode') != mode:
            return "루프·그리드 검토용 기본 관경을 지정하고 수리계산의 「표 확정」을 먼저 실행하세요. 이전 트리 파일은 출력하지 않습니다."
    return None
