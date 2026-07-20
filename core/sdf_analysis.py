# -*- coding: utf-8 -*-
"""SDF 스프링클러 배관망 분석 (Phase2b 코어 추출 — domain-slim).

.sdf(PIPENET) 파싱·인접·최단경로·구경/피팅 통계. 기하는 dxf_geometry 에 의존.
"""
from __future__ import annotations

import math
import xml.etree.ElementTree as ET
from pathlib import Path
from dxf_geometry import _to_float


def _sdf_counts_only(sdf_path: Path | None) -> dict:
    if sdf_path is None or not sdf_path.exists():
        return {}
    root = ET.parse(sdf_path).getroot()
    return {
        "pipes": len(root.findall(".//Pipe")),
        "nozzles": len(root.findall(".//Nozzle")),
        "equipment": len(root.findall(".//Equipment")),
    }

def _sdf_parse_nodes(root) -> dict[str, dict]:
    """SDF <Node> → {label: {id, x, y, z(elevation)}}."""
    nodes: dict[str, dict] = {}
    for node in root.findall(".//Node"):
        label = node.attrib.get("label", "")
        pos = node.find("Position")
        if not label or pos is None:
            continue
        nodes[label] = {
            "id": label,
            "x": _to_float(pos.attrib.get("x")),
            "y": _to_float(pos.attrib.get("y")),
            "z": _to_float(node.attrib.get("elevation")),
        }
    return nodes

def _sdf_parse_pipes_equipment(root) -> tuple[list[dict], list[dict]]:
    """SDF <Pipe-set>/<Pipe> → (pipes, equipment). material 은 직전 Pipe-type Name 을 따른다."""
    pipes: list[dict] = []
    equipment: list[dict] = []
    material = "UNKNOWN"
    for pipe_set in root.findall(".//Pipe-set"):
        pipe_type = pipe_set.find("Pipe-type")
        name = pipe_type.find("Name") if pipe_type is not None else None
        if name is not None and name.text:
            material = name.text.strip()
        for pipe in pipe_set.findall("Pipe"):
            label = pipe.attrib.get("label", "")
            input_node = pipe.attrib.get("input", "")
            output_node = pipe.attrib.get("output", "")
            bore_mm = _to_float(pipe.attrib.get("bore")) * 1000.0
            length_m = _to_float(pipe.attrib.get("length"))
            rise_m = _to_float(pipe.attrib.get("rise"))
            c_factor = _to_float(pipe.attrib.get("roughness-or-c"))
            fittings: list[dict] = []
            for fitting in pipe.findall(".//Fitting"):
                fittings.append(
                    {
                        "type": fitting.attrib.get("type", ""),
                        "count": int(_to_float(fitting.attrib.get("count"), 0)),
                    }
                )
            waypoint_positions: list[dict] = []
            waypoints = pipe.find("Waypoints")
            if waypoints is not None:
                for wp in waypoints.findall("Position"):
                    waypoint_positions.append(
                        {
                            "x": _to_float(wp.attrib.get("x")),
                            "y": _to_float(wp.attrib.get("y")),
                        }
                    )
            pipes.append(
                {
                    "label": label,
                    "input_node": input_node,
                    "output_node": output_node,
                    "bore_mm": bore_mm,
                    "length_m": length_m,
                    "rise_m": rise_m,
                    "c_factor": c_factor,
                    "material": material,
                    "fittings": fittings,
                    "fitting_summary": ", ".join(f"{f['type']}({f['count']})" for f in fittings) or "-",
                    "waypoints": waypoint_positions,
                }
            )
            for eq in pipe.findall(".//Equipment"):
                equipment.append(
                    {
                        "label": eq.attrib.get("label", ""),
                        "pipe_label": label,
                        "description": eq.attrib.get("description", ""),
                        "equivalent_length_m": _to_float(eq.attrib.get("equivalent-length")),
                        "rel_position": _to_float(eq.attrib.get("rel-position"), 0.5),
                    }
                )
    return pipes, equipment

def _sdf_parse_nozzles(root, nodes: dict) -> list[dict]:
    """SDF <Nozzle> → 입력노드 좌표를 붙인 노즐(헤드) 리스트."""
    nozzles: list[dict] = []
    for nozzle in root.findall(".//Nozzle"):
        label = nozzle.attrib.get("label", "")
        input_node = nozzle.attrib.get("input", "")
        node = nodes.get(input_node, {})
        nozzles.append(
            {
                "label": label,
                "input_node": input_node,
                "x": node.get("x"),
                "y": node.get("y"),
                "z": node.get("z"),
            }
        )
    return nozzles

def _sdf_build_adjacency(pipes: list[dict]) -> tuple[dict, dict]:
    """pipes → (outgoing[input_node]→pipes, adjacency[node]→(이웃,길이,라벨) 무방향)."""
    outgoing: dict[str, list[dict]] = {}
    adjacency: dict[str, list[tuple[str, float, str]]] = {}
    for pipe in pipes:
        outgoing.setdefault(pipe["input_node"], []).append(pipe)
        adjacency.setdefault(pipe["input_node"], []).append((pipe["output_node"], pipe["length_m"], pipe["label"]))
        adjacency.setdefault(pipe["output_node"], []).append((pipe["input_node"], pipe["length_m"], pipe["label"]))
    return outgoing, adjacency

def _sdf_av_node(pipes: list[dict], equipment: list[dict]) -> tuple[str, str]:
    """알람밸브(A/V) 앵커 노드 추정 → (av_node, av_pipe_label). 못 찾으면 첫 배관 입력노드."""
    pipe_by_label = {p["label"]: p for p in pipes}
    av_equipment = next((e for e in equipment if (e.get("description") or "").upper().replace(" ", "") in {"A/V", "AV"}), None)
    av_node = ""
    av_pipe_label = ""
    if av_equipment:
        av_pipe_label = str(av_equipment.get("pipe_label") or "")
        av_pipe = pipe_by_label.get(av_pipe_label)
        if av_pipe:
            av_node = av_pipe.get("output_node") or av_pipe.get("input_node") or ""
    if not av_node and pipes:
        av_node = pipes[0]["input_node"]
    return av_node, av_pipe_label

def _sdf_dijkstra(av_node: str, adjacency: dict) -> dict[str, float]:
    """A/V 앵커에서 각 노드까지 최단(누적 length) 거리."""
    dist = {av_node: 0.0} if av_node else {}
    visited: set[str] = set()
    while dist:
        current = min((n for n in dist if n not in visited), key=lambda n: dist[n], default=None)
        if current is None:
            break
        visited.add(current)
        for nxt, length, _pipe_label in adjacency.get(current, []):
            nd = dist[current] + max(length, 0.0)
            if nxt not in dist or nd < dist[nxt]:
                dist[nxt] = nd
    return dist

def _sdf_farthest_heads(nozzles: list[dict], dist: dict) -> list[dict]:
    """A/V 에서 먼 순으로 정렬한 헤드 상위 30개 (가장 먼 구간 검토용)."""
    return sorted(
        [
            {**n, "distance_from_av_m": dist.get(str(n.get("input_node")), 0.0)}
            for n in nozzles
        ],
        key=lambda r: r.get("distance_from_av_m", 0.0),
        reverse=True,
    )[:30]

def _sdf_length_checks(pipes: list[dict], nodes: dict) -> list[dict]:
    """SDF length 와 XY(+rise) 기하 길이가 허용오차(5% 또는 0.5m) 초과인 배관."""
    length_checks: list[dict] = []
    for pipe in pipes:
        n1 = nodes.get(pipe["input_node"])
        n2 = nodes.get(pipe["output_node"])
        if not n1 or not n2:
            continue
        pts = [(n1["x"], n1["y"])]
        pts.extend((wp["x"], wp["y"]) for wp in pipe.get("waypoints") or [])
        pts.append((n2["x"], n2["y"]))
        xy_m = 0.0
        for i in range(len(pts) - 1):
            xy_m += math.hypot(pts[i + 1][0] - pts[i][0], pts[i + 1][1] - pts[i][1]) / 1000.0
        geom_m = math.hypot(xy_m, pipe.get("rise_m") or 0.0)
        diff_m = abs(geom_m - pipe["length_m"])
        tol_m = max(0.5, pipe["length_m"] * 0.05)
        if diff_m > tol_m:
            length_checks.append(
                {
                    "pipe_label": pipe["label"],
                    "sdf_length_m": round(pipe["length_m"], 3),
                    "xy_length_m": round(geom_m, 3),
                    "diff_m": round(diff_m, 3),
                    "reason": "SDF length와 XY 좌표거리 차이가 허용오차(5% 또는 0.5m)를 초과합니다.",
                }
            )
    return length_checks

def _sdf_bore_reductions(pipes: list[dict], outgoing: dict) -> list[dict]:
    """노드에서 하류 배관 구경이 작아지는(축소) 지점."""
    bore_reductions: list[dict] = []
    for pipe in pipes:
        for child in outgoing.get(pipe["output_node"], []):
            if child["bore_mm"] and pipe["bore_mm"] and child["bore_mm"] < pipe["bore_mm"]:
                bore_reductions.append(
                    {
                        "from_pipe": pipe["label"],
                        "to_pipe": child["label"],
                        "node": pipe["output_node"],
                        "from_bore_mm": round(pipe["bore_mm"], 1),
                        "to_bore_mm": round(child["bore_mm"], 1),
                    }
                )
    return bore_reductions

def _sdf_branch_nodes(pipes: list[dict], nodes: dict) -> list[dict]:
    """차수(degree) 3 이상 분기 노드 (degree 내림차순)."""
    node_degree: dict[str, int] = {}
    for pipe in pipes:
        node_degree[pipe["input_node"]] = node_degree.get(pipe["input_node"], 0) + 1
        node_degree[pipe["output_node"]] = node_degree.get(pipe["output_node"], 0) + 1
    return [
        {"node": node, "degree": degree, **nodes.get(node, {})}
        for node, degree in sorted(node_degree.items(), key=lambda x: (-x[1], x[0]))
        if degree >= 3
    ]

def _sdf_fitting_stats(pipes: list[dict]) -> tuple[dict, list]:
    """부속(엘보/티 등) 총계 + 부속 집중(>=2) 핫스팟."""
    fitting_summary: dict[str, int] = {}
    fitting_hotspots: list[dict] = []
    for pipe in pipes:
        total = 0
        for fitting in pipe["fittings"]:
            fitting_summary[fitting["type"]] = fitting_summary.get(fitting["type"], 0) + fitting["count"]
            total += fitting["count"]
        if total >= 2:
            fitting_hotspots.append(
                {
                    "pipe_label": pipe["label"],
                    "fitting_count": total,
                    "fittings": pipe["fitting_summary"],
                    "reason": "엘보/티 등 부속 집중 구간입니다. CAD 도면의 굴곡/분기 위치와 대조가 필요합니다.",
                }
            )
    return fitting_summary, fitting_hotspots

def _sdf_vertical_pipes(pipes: list[dict]) -> list[dict]:
    """|rise| >= 3m 수직 배관 (층고/단면 대조용)."""
    return [
        {
            "pipe_label": p["label"],
            "input_node": p["input_node"],
            "output_node": p["output_node"],
            "length_m": round(p["length_m"], 3),
            "rise_m": round(p["rise_m"], 3),
            "bore_mm": round(p["bore_mm"], 1),
        }
        for p in pipes
        if abs(p.get("rise_m") or 0.0) >= 3.0
    ]

def _sdf_graph_pipes(pipes: list[dict], nodes: dict,
                     length_checks: list[dict], bore_reductions: list[dict]) -> list[dict]:
    """프론트 시각화용 배관 폴리라인 + 상태색(길이이상=red, 구경축소=orange)."""
    graph_pipes: list[dict] = []
    for p in pipes:
        n1 = nodes.get(p["input_node"])
        n2 = nodes.get(p["output_node"])
        if not n1 or not n2:
            continue
        path = [[n1["x"], n1["y"]]]
        path.extend([[wp["x"], wp["y"]] for wp in p.get("waypoints") or []])
        path.append([n2["x"], n2["y"]])
        status = "red" if any(x["pipe_label"] == p["label"] for x in length_checks) else "normal"
        if any(x["to_pipe"] == p["label"] or x["from_pipe"] == p["label"] for x in bore_reductions):
            status = "orange" if status == "normal" else status
        graph_pipes.append(
            {
                "label": p["label"],
                "input_node": p["input_node"],
                "output_node": p["output_node"],
                "bore_mm": round(p["bore_mm"], 1),
                "length_m": round(p["length_m"], 3),
                "material": p["material"],
                "status": status,
                "path": path,
            }
        )
    return graph_pipes

def _analyze_sdf_sprinkler_network(sdf_path: Path) -> dict:
    root = ET.parse(sdf_path).getroot()

    titles = [t.text.strip() for t in root.findall(".//Title") if t.text and t.text.strip()]
    nodes = _sdf_parse_nodes(root)
    pipes, equipment = _sdf_parse_pipes_equipment(root)
    nozzles = _sdf_parse_nozzles(root, nodes)

    outgoing, adjacency = _sdf_build_adjacency(pipes)
    av_node, av_pipe_label = _sdf_av_node(pipes, equipment)
    dist = _sdf_dijkstra(av_node, adjacency)

    farthest_heads = _sdf_farthest_heads(nozzles, dist)
    length_checks = _sdf_length_checks(pipes, nodes)
    bore_reductions = _sdf_bore_reductions(pipes, outgoing)
    branch_nodes = _sdf_branch_nodes(pipes, nodes)
    fitting_summary, fitting_hotspots = _sdf_fitting_stats(pipes)
    vertical_pipes = _sdf_vertical_pipes(pipes)
    graph_pipes = _sdf_graph_pipes(pipes, nodes, length_checks, bore_reductions)

    return {
        "title": " / ".join(titles) or sdf_path.name,
        "filename": sdf_path.name,
        "summary": {
            "node_count": len(nodes),
            "pipe_count": len(pipes),
            "nozzle_count": len(nozzles),
            "equipment_count": len(equipment),
            "av_node": av_node,
            "av_pipe_label": av_pipe_label,
            "length_issue_count": len(length_checks),
            "bore_reduction_count": len(bore_reductions),
            "branch_node_count": len(branch_nodes),
            "vertical_pipe_count": len(vertical_pipes),
        },
        "nodes": list(nodes.values()),
        "pipes": graph_pipes,
        "nozzles": nozzles,
        "equipment": equipment,
        "farthest_heads": farthest_heads,
        "length_checks": length_checks[:80],
        "bore_reductions": bore_reductions[:80],
        "branch_nodes": branch_nodes[:80],
        "fitting_summary": [{"type": k, "count": v} for k, v in sorted(fitting_summary.items())],
        "fitting_hotspots": fitting_hotspots[:80],
        "vertical_pipes": vertical_pipes[:80],
        "checklist": [
            "CAD 도면의 알람밸브 위치가 SDF A/V 추정 노드와 일치하는지 확인",
            "SDF 최원단 헤드 30개가 CAD 평면도상 검토 영역의 헤드 30개와 1:1 매칭되는지 확인",
            "배관 길이 불일치 후보는 CAD 실측 길이와 SDF length 값을 대조",
            "구경 축소 지점은 CAD 라벨의 관경 표기와 SDF bore 값을 대조",
            "엘보/티 집중 구간은 도면상 굴곡/분기 개수와 SDF Fittings count를 대조",
            "수직 배관은 건축 단면/층고와 SDF rise 및 length를 대조",
        ],
    }
