# -*- coding: utf-8 -*-
"""C3 — 유수검지장치 위치와 방호구역 (지시서 §7).

**밸브는 사람이 찍는다.** 여기서 하는 일은 후보를 추리고, 사람이 찍은 자리를
근거로 구역을 나누고, 나눈 구역이 실제로 배관이 닿는 범위인지 보는 것뿐이다.

구역 배정의 거리는 **직선거리가 아니라 문을 지나는 경로거리**다(§7.2). 직선으로
재면 벽 하나를 사이에 둔 실이 가까워 보여, 실제로는 복도를 크게 돌아야 닿는 실이
엉뚱한 밸브에 붙는다. 문은 C160 이 낸 가상 간선(`virtual_edges`)이다 — 벽
중심선은 지나갈 수 없고 가상 간선만 지나갈 수 있다.

닿지 않는 실은 **가까운 밸브에 억지로 붙이지 않는다.** 붙여 두면 그 실의 헤드가
어느 유수검지장치에도 물려 있지 않은 채로 C4·C5 를 통과한다.

실측(양주옥정 공동주택, 실 58·가상 간선 718): 문으로 이어진 쌍은 11 개뿐이고
49 개 실이 `ROOM_UNREACHABLE` 로 남는다. 공차 문제가 아니다 — 간선 중점에서 가장
가까운 실 경계까지의 거리가 mm 가 아니라 m 단위로 흩어져 있다(≤2mm 130 / ≤1m 112
/ ≤5m 307 / >5m 90). C1 이 채택한 실이 도면의 일부만 덮고 있어서, 나머지 문들은
버려진 face 의 경계다. 공차를 늘려 메우면 없는 통행로를 지어내는 것이므로 늘리지
않는다. 이 수치를 올리는 일은 C1 인식의 몫이다.
"""
from __future__ import annotations

import heapq
import math
from dataclasses import dataclass, field
from typing import Any

from ..recognize.spatial import NodeIndex, centroid, point_in_polygon

# §7.3 — 부압식을 빠뜨리지 마라. 파이프라인 v4~v6 문서가 5종에서 이걸 누락했었다.
SYSTEM_TYPES = ("습식", "건식", "준비작동식", "부압식", "일제살수식")

# §7.1 1항 — 후보는 확정된 SHAFT·STAIR 코어뿐이다. 승강기 샤프트는 유수검지장치실이
# 될 수 없어 후보에서 뺀다.
VALVE_CORE_KINDS = ("SHAFT", "STAIR")

# §7.1 체크리스트 중 도면으로 판정할 수 없는 네 항목. 사람이 참/거짓으로 답한다.
MANUAL_REQUIREMENTS = ("mount_height_ok", "door_size_ok", "room_temp_ok", "signage_ok")

REQUIREMENT_LABELS = {
    "mount_height_ok": "설치 높이 0.8~1.5m",
    "door_size_ok": "전용실 출입문 0.5m × 1.0m 이상",
    "room_temp_ok": "실온 4℃ 이상 유지",
    "signage_ok": "「유수검지장치실」 표지",
}

# 실 폴리곤 꼭짓점과 가상 간선 끝점을 같은 점으로 묶는 공차. C170 은 폴리곤을
# 0.1mm 로, C160 은 간선을 0.01mm 로 반올림해 내보내므로 그 차이만 흡수하면 된다.
# 격자 반올림이 아니라 epsilon 군집이다 — 격자는 눈금을 사이에 둔 두 점을 갈라놓는다.
_JOIN_TOL_MM = 2.0

_MM_PER_M = 1000.0


class ZoningError(ValueError):
    """사람에게 그대로 보여줄 실패. `code` 는 화면이 분기에 쓴다."""

    def __init__(self, code: str, message: str, **detail: Any):
        self.code = code
        self.detail = detail
        super().__init__(message)


# ────────────────────────────────────────────────────────────────────────────
# C3.1 밸브 후보 (§7.1)
# ────────────────────────────────────────────────────────────────────────────

def valve_candidates(draft) -> list[dict[str, Any]]:
    """확정된 SHAFT·STAIR 코어. 화면이 이걸 하이라이트하고 사람이 그 안을 찍는다.

    `confirmed` 가 `None` 인 코어는 후보가 아니다 — "아직 모른다" 를 후보로 내면
    사람이 확정하지 않은 코어에 입상관이 서고, GATE 가 왜 있는지가 사라진다.
    """
    out = []
    for core in draft.cores:
        if core.confirmed is not True or core.kind.upper() not in VALVE_CORE_KINDS:
            continue
        point = representative_point(core.polygon)
        out.append({
            "core_id": core.id, "kind": core.kind,
            "polygon": core.polygon, "area_m2": core.area_m2,
            "center": [round(v, 1) for v in point] if point else None,
        })
    return out


def representative_point(polygon) -> tuple[float, float] | None:
    """폴리곤 **안**의 점 하나. 없으면 `None`.

    무게중심을 그냥 쓰면 안 된다 — 오목한 실에서는 밖으로 나간다(양주옥정 코어
    11개 중 1개). 화면은 이 점으로 밸브 자리를 제안하는데, 밖으로 나간 점을
    제안하면 사람이 제안대로 찍었을 때 "실 밖" 이라며 거절당한다.

    밖으로 나갔으면 무게중심의 y 로 수평선을 그어, 폴리곤 내부 구간 중 가장 넓은
    곳의 중점을 쓴다. 어느 쪽도 안 되면 지어내지 않고 `None` 이다.
    """
    if len(polygon) < 3:
        return None
    cx, cy = centroid(polygon)
    if point_in_polygon((cx, cy), polygon):
        return (cx, cy)

    crossings = []
    for i, p in enumerate(polygon):
        x1, y1 = float(p[0]), float(p[1])
        q = polygon[(i + 1) % len(polygon)]
        x2, y2 = float(q[0]), float(q[1])
        if (y1 > cy) != (y2 > cy):
            crossings.append(x1 + (cy - y1) * (x2 - x1) / (y2 - y1))
    crossings.sort()

    widest = max(zip(crossings[0::2], crossings[1::2]),
                 key=lambda span: span[1] - span[0], default=None)
    if widest is None:
        return None
    point = ((widest[0] + widest[1]) * 0.5, cy)
    return point if point_in_polygon(point, polygon) else None


def check_requirements(answers: dict) -> dict[str, bool]:
    """네 항목을 참/거짓으로 받는다. 빠진 답을 거짓으로 접지 않는다.

    빠진 답을 거짓으로 읽으면 "묻지 않았다" 가 "요건 미충족" 으로 확정되고,
    참으로 읽으면 확인하지 않은 요건이 충족으로 남는다. 어느 쪽도 사람의 답이
    아니므로 답이 없으면 여기서 멈춘다.
    """
    missing = [key for key in MANUAL_REQUIREMENTS if not isinstance(answers.get(key), bool)]
    if missing:
        raise ZoningError(
            "VALVE_REQUIREMENTS_REQUIRED",
            "설치 요건 " + ", ".join(REQUIREMENT_LABELS[k] for k in missing)
            + " 을(를) 참/거짓으로 확인해야 합니다.", fields=missing)
    return {key: bool(answers[key]) for key in MANUAL_REQUIREMENTS}


# ────────────────────────────────────────────────────────────────────────────
# 문을 지나는 실 인접 그래프
# ────────────────────────────────────────────────────────────────────────────

@dataclass
class RoomGraph:
    """실 사이의 통행 그래프. 간선은 문(가상 간선) 하나에 하나씩."""

    centers: dict = field(default_factory=dict)
    doors: dict = field(default_factory=dict)
    orphans: list = field(default_factory=list)

    def neighbors(self, room_id: str):
        return self.doors.get(room_id, ())


def build_room_graph(draft) -> RoomGraph:
    """`virtual_edges` 로 실을 잇는다. 벽 중심선은 잇지 않는다.

    가상 간선 하나는 두 실의 공유 변이다 — C170 이 같은 평면 그래프에서 face 를
    뽑으므로 좌표가 정확히 맞는다. GATE 에서 실을 합치면 그 사이의 문은 어느
    폴리곤의 변도 아니게 되어 자연히 사라지고(안쪽이 됐으므로 맞다), 실을 자르면
    새로 생긴 변은 가상 간선이 아니라 통행이 안 된다(사람이 그은 칸막이다).
    """
    index = NodeIndex(_JOIN_TOL_MM)
    edge_rooms: dict[tuple[int, int], list[str]] = {}
    centers: dict[str, tuple[float, float]] = {}

    for room in draft.rooms:
        if len(room.polygon) < 3:
            continue
        centers[room.id] = centroid(room.polygon)
        nodes = [index.add(float(p[0]), float(p[1])) for p in room.polygon]
        for i, a in enumerate(nodes):
            b = nodes[(i + 1) % len(nodes)]
            if a != b:
                edge_rooms.setdefault((a, b) if a < b else (b, a), []).append(room.id)

    doors: dict[str, list[tuple[str, float, tuple[float, float]]]] = {}
    for edge in draft.virtual_edges:
        p1, p2 = edge.get("p1"), edge.get("p2")
        if not p1 or not p2:
            continue
        n1 = index.find(float(p1[0]), float(p1[1]))
        n2 = index.find(float(p2[0]), float(p2[1]))
        if n1 is None or n2 is None or n1 == n2:
            continue
        rooms = sorted(set(edge_rooms.get((n1, n2) if n1 < n2 else (n2, n1)) or ()))
        if len(rooms) != 2:
            # 0·1개는 한쪽이 건물 바깥이거나(현관) 편집으로 사라진 문이다. 통행로로
            # 세면 바깥을 거쳐 도달했다고 말하게 된다. 3개 이상은 실 폴리곤이 같은
            # 변을 두 번 지나는 슬리버라 어느 쪽이 통행로인지 말할 수 없다.
            continue
        mid = ((float(p1[0]) + float(p2[0])) * 0.5, (float(p1[1]) + float(p2[1])) * 0.5)
        a, b = rooms
        for src, dst in ((a, b), (b, a)):
            doors.setdefault(src, []).append((dst, _hop_mm(centers, src, dst, mid), mid))

    orphans = [rid for rid in centers if rid not in doors]
    return RoomGraph(centers=centers, doors=doors, orphans=orphans)


def _hop_mm(centers, src: str, dst: str, door: tuple[float, float]) -> float:
    """실 중심 → 문 → 실 중심. 배관이 실제로 도는 길이의 하한이다."""
    sx, sy = centers[src]
    dx, dy = centers[dst]
    return math.hypot(sx - door[0], sy - door[1]) + math.hypot(dx - door[0], dy - door[1])


# ────────────────────────────────────────────────────────────────────────────
# C3.2 방호구역 분할 (§7.2)
# ────────────────────────────────────────────────────────────────────────────

def seed_room(draft, point) -> str | None:
    """밸브를 찍은 자리를 품은 실. 없으면 `None` — 가까운 실로 끌어붙이지 않는다."""
    px, py = float(point[0]), float(point[1])
    for room in draft.rooms:
        if len(room.polygon) >= 3 and point_in_polygon((px, py), room.polygon):
            return room.id
    return None


def assign_rooms(draft, valves: list[dict], graph: RoomGraph) -> tuple[dict, dict]:
    """실마다 경로거리가 가장 짧은 밸브. 반환 `(실→밸브, 실→거리m)`.

    여러 밸브를 동시에 출발점으로 놓고 한 번만 훑는다(다중 출발 다익스트라).
    밸브별로 따로 훑으면 같은 답을 밸브 수만큼 다시 계산한다.
    """
    dist: dict[str, float] = {}
    owner: dict[str, str] = {}
    # 거리가 같으면 밸브 id 순으로 갈린다. 순서를 정해 두지 않으면 같은 도면이
    # 실행할 때마다 다른 구역으로 갈라져 감사에서 재현되지 않는다.
    heap = [(0.0, valve["id"], valve["room_id"])
            for valve in valves if valve.get("room_id")]
    heapq.heapify(heap)

    while heap:
        d, valve_id, room_id = heapq.heappop(heap)
        if room_id in dist:
            continue
        dist[room_id] = d
        owner[room_id] = valve_id
        for nxt, hop, _mid in graph.neighbors(room_id):
            if nxt not in dist:
                heapq.heappush(heap, (d + hop, valve_id, nxt))

    return owner, {rid: round(d / _MM_PER_M, 2) for rid, d in dist.items()}


def split_zones(draft, valves: list[dict], constraints: dict) -> dict[str, Any]:
    """밸브가 구역을 정의한다. 반환은 design.json 의 `zones` 와 `flags`.

    §7.2 3항은 면적 초과 시 자동 분할을 시킨다. 그런데 방호구역에는 유수검지장치가
    하나씩 있어야 하므로(NFTC 103 2.3.1.2) 자동으로 쪼개면 밸브 없는 구역이 생긴다.
    밸브는 사람이 찍는 것이라 코드가 지어낼 수 없다 — 그래서 쪼개는 대신 초과
    사실을 flag 로 내보내 사람이 밸브를 하나 더 찍게 한다. [문서정합 §7.2]
    """
    graph = build_room_graph(draft)
    owner, dist_m = assign_rooms(draft, valves, graph)
    area_max = constraints.get("zone_area_max_m2")
    floors_max = constraints.get("zone_floors_max")

    by_valve: dict[str, list] = {v["id"]: [] for v in valves}
    for room in draft.rooms:
        valve_id = owner.get(room.id)
        if valve_id:
            by_valve[valve_id].append(room)

    zones: list[dict[str, Any]] = []
    flags: list[dict[str, Any]] = []
    for i, valve in enumerate(valves, start=1):
        rooms = by_valve[valve["id"]]
        area = round(sum(r.area_m2 for r in rooms), 2)
        floors = sorted({r.floor for r in rooms if r.floor})
        zone = {
            "id": f"Z-{valve['floor']}-{i:02d}",
            "valve_id": valve["id"], "floor": valve["floor"],
            "rooms": [r.id for r in rooms], "area_m2": area,
            "reachability": "ok" if rooms else "no_rooms",
            "farthest_room_m": max((dist_m[r.id] for r in rooms), default=0.0),
            "scenario_head_count": constraints.get("scenario_head_count"),
            "scenario_head_count_source": "conservative_max",
        }
        zones.append(zone)
        if not rooms:
            flags.append(_flag("ZONE_EMPTY", zone["id"],
                               f"{valve['id']} 에서 문으로 닿는 실이 없습니다. "
                               "밸브를 찍은 실이 다른 실과 문으로 이어져 있는지 보세요."))
        if area_max and area > area_max:
            zone["reachability"] = "area_exceeded"
            flags.append(_flag(
                "ZONE_AREA_EXCEEDED", zone["id"],
                f"담당 면적 {area}㎡ 가 상한 {area_max}㎡ 를 넘습니다. 방호구역에는 "
                "유수검지장치가 하나씩 있어야 하므로 밸브를 더 찍어 나누세요."))
        if floors_max and len(floors) > floors_max:
            flags.append(_flag(
                "ZONE_FLOORS_EXCEEDED", zone["id"],
                f"담당 층이 {len(floors)}개입니다 (상한 {floors_max}개): "
                + ", ".join(floors)))

    unreached = [r.id for r in draft.rooms if r.id not in owner]
    if unreached:
        flags.append({
            "code": "ROOM_UNREACHABLE", "rooms": unreached,
            "message": f"어느 밸브에서도 문으로 닿지 않는 실 {len(unreached)}개입니다. "
                       "문이 안 잡혔거나 밸브가 모자랍니다 — 가까운 밸브에 임의로 "
                       "붙이지 않았습니다.",
        })
    return {"zones": zones, "flags": flags,
            "room_distance_m": dist_m, "unreached": unreached,
            "isolated_rooms": graph.orphans}


def _flag(code: str, zone_id: str, message: str) -> dict[str, Any]:
    return {"code": code, "zone_id": zone_id, "message": message}


# ────────────────────────────────────────────────────────────────────────────
# C3.4 헤드 사양 (§7.4)
# ────────────────────────────────────────────────────────────────────────────

# 한백 표준 F사 유형. §7.4 가 지정한 값 그대로다.
FLEX_SPEC = {"equivalent_length_m": 22.4, "dn": 25, "inner_dia_mm": 28.0,
             "c_factor": 120.0, "physical_length_m": 0.7}


def head_spec(room, constraints: dict) -> dict[str, Any]:
    """`constraints` 를 **읽기만** 한다. R·표시온도를 여기서 다시 정하지 않는다.

    다시 정하면 같은 건물에 두 벌의 기준이 생기고, 어느 쪽으로 검사를 통과했는지
    말할 수 없게 된다.
    """
    has_finish = room.ceiling.has_finish
    if has_finish is None:
        raise ZoningError("MISSING_BUILDING_FACT",
                          f"{room.id} 의 반자 유무가 확정되지 않았습니다.",
                          field="rooms[].ceiling.has_finish", room_id=room.id)
    return {
        # 반자가 있으면 하향식이고 신축배관은 반자 속에만 쓴다.
        "orientation": "pendent" if has_finish else "upright",
        "temp_rating_c": constraints.get("temp_rating_c"),
        "quick_response": constraints.get("quick_response_required"),
        "k_factor": constraints.get("k_factor"),
        "flex": dict(FLEX_SPEC) if has_finish else None,
    }
