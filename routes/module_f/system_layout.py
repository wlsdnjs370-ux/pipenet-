"""Preserve selected system-path bearings independently of drawing underlays."""
from __future__ import annotations

from copy import deepcopy
from pathlib import Path
import math
import sys
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from remote30_full_network import CombinedTables


def set_elevation_mode(riser: dict, mode: str, measured: dict | None = None) -> dict:
    """Apply an explicit same-level choice; never infer height from drawing Y.

    'drawing' retains the existing inferred elevations. 'same_level' starts all
    nodes at the AV elevation; local offsets can then be edited in the network.
    Measured edge lengths come from the extraction graph, before rise clamping.
    """
    if mode not in {"drawing", "same_level"}:
        raise ValueError("계통도 표고 기준이 올바르지 않습니다.")
    result = deepcopy(riser)
    result["elevation_mode"] = mode
    if mode == "drawing":
        return result
    nodes = {str(n["label"]): n for n in result["nodes"]}
    for node in nodes.values():
        node["inferred_elevation"] = node.get("elevation", 0)
        node["elevation"] = 0.0
        node["elev_source"] = "user_same_level"
    for pipe in result["pipes"]:
        a, b = nodes[str(pipe["in"])], nodes[str(pipe["out"])]
        pa, pb = (float(a["x"]), float(a["y"])), (float(b["x"]), float(b["y"]))
        key = (min(pa, pb), max(pa, pb))
        # Clean-network extraction already records the measured length.
        length_mm = (measured or {}).get(key, pipe.get("length_measured_mm"))
        if length_mm is None:
            raise ValueError(f"{pipe['label']}: 표고 보정 전 실측 길이를 찾을 수 없습니다.")
        length_m = float(length_mm) / 1000.0
        if not math.isfinite(length_m) or length_m <= 0:
            raise ValueError(f"{pipe['label']}: 실측 배관 길이가 올바르지 않습니다.")
        pipe["inferred_elev"] = pipe.get("elev", 0)
        pipe["inferred_length"] = pipe["length"]
        pipe["elev"] = 0.0
        pipe["elev_source"] = "user_same_level"
        pipe["length"] = round(length_m, 3)
    result["total_pipe_length_m"] = sum(float(p["length"]) for p in result["pipes"])
    return result


def layout_selected_system(combined: CombinedTables, riser: dict, machine_labels: list[str],
                           pump_junction: str | None, *, machine_room: dict | None = None,
                           source_drop_m: float = 0.0) -> bool:
    """Place extracted paths using metre lengths and real elevations, at the AV.

    Template risers retain their legacy display convention. DXF-extracted paths
    retain every edge; loop and length validation is shared with the KFP export.
    Machine-room XYZ uses the same metre scale, rebased at its shared node.
    A nonzero source-drop setting must agree with the supplied elevations.
    The drawing coordinates in the source object remain unchanged.
    """
    if riser.get("extracted_from") not in {"dxf", "dxf_clean_network"}:
        return False
    source = str(Path(__file__).resolve().parents[2] / "pipenet_converter" / "src")
    if source not in sys.path:
        sys.path.insert(0, source)
    from pipenet_converter.physical_layout import LayoutEdge, physical_layout
    raw = {str(n["label"]): n for n in riser["nodes"]}
    anchor = str(riser["av_node_label"])
    positions = physical_layout(
        {k: (float(n["x"]),
             0.0 if riser.get("system_coordinate_mode") == "section_xz" else float(n["y"]),
             float(n.get("elevation", 0)))
         for k, n in raw.items()},
        [LayoutEdge(str(p["label"]), str(p["in"]), str(p["out"]), float(p["length"]))
         for p in riser["pipes"]], roots=[anchor])
    at = {str(n["label"]): n for n in combined.nodes}
    ax, ay = float(at[anchor]["x"]), float(at[anchor]["y"])
    dest = {k: (ax+x*1000, ay+y*1000, z) for k, (x, y, z) in positions.items()}
    dx = dy = 0.0
    if pump_junction and pump_junction in dest:
        dx = dest[pump_junction][0] - float(at[pump_junction]["x"])
        dy = dest[pump_junction][1] - float(at[pump_junction]["y"])
    machine_dest = {}
    machine_edges = None
    if machine_room and pump_junction in dest:
        from routes.module_f.connections import machine_connection
        conn = machine_connection(machine_room)
        conn_label = str(conn["label"])
        px, py, pz = dest[pump_junction]
        # Both copies of the joint denote ONE physical point. Offset the whole
        # machine room by its connection elevation, not its source elevation.
        dz = pz - float(conn.get("elevation", 0))
        room_nodes = {
            str(n["label"]): (float(n["x"]), float(n["y"]),
                              float(n.get("elevation", 0)) + dz)
            for n in machine_room["nodes"]}
        if source_drop_m:
            source_node = next((n for n in machine_room["nodes"]
                                if str(n.get("io_node", "")).lower() == "input"), None)
            expected = min(float(n.get("elevation", 0)) for n in raw.values()) - abs(source_drop_m)
            if source_node is None or abs(room_nodes[str(source_node["label"])][2] - expected) > .001:
                raise ValueError("기계실 공통 절점의 표고와 입력한 수원 낙차가 맞지 않습니다. "
                                 "기계실·계통도 표고와 수원 낙차를 확인하세요.")
        room_positions = physical_layout(
            room_nodes,
            [LayoutEdge(str(p["label"]), str(p["in"]), str(p["out"]), float(p["length"]))
             for p in machine_room["pipes"]], roots=[conn_label])
        machine_dest = {k: (px + x*1000, py + y*1000, z)
                        for k, (x, y, z) in room_positions.items() if k != conn_label}
        # Original reference geometry uses its own translation, never a bbox
        # scale that changes when the plan extent or background lines change.
        tx, ty = px - float(conn["x"]), py - float(conn["y"])
        machine_edges = [[e[0]+tx, e[1]+ty, e[2]+tx, e[3]+ty]
                         for e in machine_room.get("plan_edges") or []]
    mr = set(machine_labels)
    rows = []
    for old in combined.nodes:
        node = dict(old)
        label = str(node["label"])
        if label in dest:
            node["x"], node["y"], node["elevation"] = dest[label]
        elif label in machine_dest:
            node["x"], node["y"], node["elevation"] = machine_dest[label]
        elif label in mr:
            node["x"], node["y"] = float(node["x"])+dx, float(node["y"])+dy
        rows.append(node)
    combined.nodes = rows
    if machine_edges is not None:
        combined.machine_room_plan_edges = machine_edges
    else:
        combined.machine_room_plan_edges = [
            [e[0]+dx, e[1]+dy, e[2]+dx, e[3]+dy]
            for e in combined.machine_room_plan_edges or []]
    if machine_dest:
        combined.layout_status["machineroom"] = "ok:공통 절점·실제 길이/표고"
    combined.layout_status["riser"] = (
        "ok:선택 경로·" + riser['elevation_notice'] if riser.get('elevation_notice')
        else "ok:선택 경로·입력 표고")
    return True
