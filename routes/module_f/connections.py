"""Authoritative shared nodes at Module F drawing boundaries."""
from __future__ import annotations

import math
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from remote30_full_network import CombinedTables


def machine_connection(machine_room: dict) -> dict:
    """Resolve the extracted connection node, never a pre-snap mouse click.

    Old extraction results without a label use the last path node, matching the
    shared engine contract. An explicit but missing/ambiguous label is an error.
    """
    nodes = machine_room.get("nodes") or []
    label = str(machine_room.get("conn_node_label") or
                (nodes[-1]["label"] if nodes else ""))
    matches = [n for n in nodes if str(n.get("label")) == label]
    if len(matches) != 1:
        raise ValueError(f"기계실 접속 절점 «{label}»을 하나로 확인할 수 없습니다.")
    node = matches[0]
    try:
        valid = all(math.isfinite(float(node[k])) for k in ("x", "y"))
    except (KeyError, TypeError, ValueError):
        valid = False
    if not valid:
        raise ValueError(f"기계실 접속 절점 «{label}»의 좌표가 올바르지 않습니다.")
    return node


def align_plan_connection(combined: CombinedTables, head_nodes: list[dict],
                          anchor: str) -> None:
    """Translate upstream drawings to the exact plan AV without adding a pipe.

    The legacy layout anchors the final row and rounds to integer mm. Neither
    row order nor rounding may move the actual shared node away from the plan.
    Elevations have already been rebased by stitch_riser_and_heads.
    """
    plan = {str(n["label"]): n for n in head_nodes}
    at = {str(n["label"]): n for n in combined.nodes}
    if anchor not in plan or anchor not in at:
        raise ValueError(f"평면도–계통도 공통 절점 «{anchor}»이 없습니다.")
    target, current = plan[anchor], at[anchor]
    dx = float(target["x"]) - float(current["x"])
    dy = float(target["y"]) - float(current["y"])
    rows = []
    for old in combined.nodes:
        node = dict(old)
        label = str(node["label"])
        if label == anchor:
            node["x"], node["y"] = float(target["x"]), float(target["y"])
        elif label not in plan:
            node["x"], node["y"] = float(node["x"]) + dx, float(node["y"]) + dy
        rows.append(node)
    combined.nodes = rows
    combined.machine_room_plan_edges = [
        [e[0] + dx, e[1] + dy, e[2] + dx, e[3] + dy]
        for e in combined.machine_room_plan_edges or []]
    # The old metadata measured before exact alignment and elevation rebasing.
    combined.meta = [(k, v) for k, v in combined.meta if k != "S740 두 기준점 거리"]
    aligned = next(n for n in rows if str(n["label"]) == anchor)
    gap = math.hypot(float(aligned["x"]) - float(target["x"]),
                     float(aligned["y"]) - float(target["y"]))
    combined.meta.append(("S740 두 기준점 거리", f"{gap:.1f} mm · 표고차 0.000 m"))


def align_new_pump_ports(combined: CombinedTables, existing_nodes: set[str]) -> None:
    """Keep zero-length pump ports at one physical XYZ in an extracted network.

    The legacy engine offsets the outlet by 400 display mm. That offset must
    not change the bearing or length of the first real downstream pipe.
    """
    at = {str(n["label"]): n for n in combined.nodes}
    for pump in combined.pumps:
        if str(pump.get("out")) in existing_nodes:
            continue  # User/extracted equipment endpoints are not display padding.
        source, outlet = at.get(str(pump.get("in"))), at.get(str(pump.get("out")))
        if source is not None and outlet is not None:
            for key in ("x", "y", "elevation"):
                outlet[key] = source[key]
