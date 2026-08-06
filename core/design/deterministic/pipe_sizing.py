"""지시서 §9.2 C550 + §9.5 C570 — 헤드를 어떻게 매달고, 관을 몇 A 로 놓는가.

[문서정합 §9.6] 명세는 PR-9 를 `pipe_routing.py` 한 파일로 적었다. 그 파일은 이미
C510~C540 으로 960 줄이고, 여기서 하는 일은 좌표를 하나도 보지 않는다 — 부하를
세고 표를 찾고 유속을 재는 것뿐이다. 골격(어디에 놓는가)과 굵기(몇 A 인가)를 한
파일에 두면 축·격자를 고칠 때마다 관경 시험이 함께 흔들린다.

관경은 두 근거로만 올라간다. 별표1(2.5.3.3)과 그 단서의 유속 상한이다. 현장은
한 단계 더 올려 잡는 관행이 있지만 규정이 아니므로 여기서 하지 않는다(금지사항 G5).
"""

from __future__ import annotations

import math
from collections import deque
from dataclasses import dataclass
from typing import Any, Callable, Iterable

from . import nftc_tables as T
from .zoning import ZoningError, head_spec

__all__ = [
    "head_connections", "edge_load", "velocity_mps", "PipeSize",
    "size_pipes", "enforce_monotonic",
]


# ────────────────────────────────────────────────────────────────────────────
# C550 — 신축배관 부착 (§9.2)
# ────────────────────────────────────────────────────────────────────────────

def head_connections(heads: Iterable, rooms: Iterable,
                     constraints: dict) -> tuple[list[dict], list[dict]]:
    """가지배관과 헤드 사이를 무엇으로 잇는지 정한다.

    반자 유무를 여기서 다시 읽지 않고 `zoning.head_spec` 을 부른다. 두 곳에서 읽으면
    한쪽만 고쳐졌을 때 하향식 헤드에 신축배관이 빠지거나 그 반대가 되고, 신축배관은
    등가길이 22.4 m 라 빠지면 그만큼의 손실이 통째로 사라진다.

    반자 유무가 확정되지 않은 실은 채워서 넘기지 않고 플래그로 돌려준다(금지사항 G9).
    """
    by_id = {room.id: room for room in rooms}
    out: list[dict] = []
    flags: list[dict] = []
    failed: set[str] = set()

    for head in heads:
        head_id = head.id if hasattr(head, "id") else head["id"]
        room_id = head.room_id if hasattr(head, "room_id") else head["room_id"]
        room = by_id.get(room_id)
        if room is None:
            flags.append({"code": "HEAD_ROOM_MISSING", "head_id": head_id,
                          "room_id": room_id,
                          "message": f"{head_id} 가 속한 실 {room_id} 가 없습니다."})
            continue
        try:
            spec = head_spec(room, constraints)
        except ZoningError as err:
            if room_id not in failed:
                failed.add(room_id)
                flags.append({"code": err.code, "room_id": room_id,
                              "message": str(err), **err.detail})
            continue
        flex = spec["flex"]
        out.append({
            "head_id": head_id, "room_id": room_id,
            "orientation": spec["orientation"],
            # 반자가 없으면 상향식 헤드를 가지배관에 직결한다. 신축배관은 반자 속
            # 위치 조정을 위한 것이라 노출 배관에서는 쓸 이유가 없다.
            "kind": "flex" if flex else "direct",
            "flex": dict(flex) if flex else None,
        })
    return out, flags


# ────────────────────────────────────────────────────────────────────────────
# C570 — 담당 헤드 수 (§9.5)
# ────────────────────────────────────────────────────────────────────────────

def edge_load(adjacency: dict, source, heads_at) -> tuple[dict, list]:
    """간선별 담당 헤드 수. 급수원에서 뻗은 트리를 후위로 훑어 올린다.

    [문서정합 §9.5] 명세의 `compute_edge_load` 는 모듈 A 에 이미 있으나 그것은 좌표
    튜플로 짜인 추출 파이프라인용이고 ezdxf 를 끌고 들어온다. C5 의 노드는 id 이므로
    같은 함수를 쓸 수 없어 방법만 같게 다시 적었다(모듈 A 수정 금지, G8).

    닿지 않는 헤드는 0 으로 세지 않고 따로 돌려준다. 0 으로 세면 그 헤드 몫의 유량이
    관경에서 통째로 빠지고, 어디서 빠졌는지는 아무 데도 남지 않는다.
    """
    counts = (dict(heads_at) if isinstance(heads_at, dict)
              else {node: 1 for node in heads_at})

    parent: dict = {source: None}
    order = [source]
    queue = deque([source])
    while queue:
        node = queue.popleft()
        for nxt in adjacency.get(node, ()):
            if nxt in parent:
                continue
            parent[nxt] = node
            order.append(nxt)
            queue.append(nxt)

    unreached = [node for node in counts if node not in parent]

    load = {node: int(counts.get(node, 0)) for node in order}
    edges: dict = {}
    for node in reversed(order[1:]):
        up = parent[node]
        edges[(up, node)] = load[node]
        load[up] += load[node]
    return edges, unreached


# ────────────────────────────────────────────────────────────────────────────
# C570 — 관경 (§9.5)
# ────────────────────────────────────────────────────────────────────────────

def velocity_mps(flow_lpm: float, inner_dia_mm: float) -> float:
    """유속. `core/hydraulic_solver.velocity_mps` 와 같은 식이다.

    그 모듈을 부르지 않는 이유는 평면 import(`from hb_rules import ...`)로 짜여 있어
    `core` 자체가 sys.path 에 올라야만 import 되기 때문이다. 식이 갈라지지 않도록
    `tests/design/test_pipe_sizing.py` 가 두 구현의 수치를 맞대어 본다.
    """
    if inner_dia_mm <= 0:
        raise ValueError("내경이 0 이하입니다. 유속을 구할 수 없습니다.")
    area_m2 = math.pi * (inner_dia_mm / 1000.0) ** 2 / 4.0
    return (flow_lpm / 60000.0) / area_m2


@dataclass
class PipeSize:
    """간선 하나의 관경과 그 근거. `basis` 가 없으면 굵기를 설명할 수 없다."""

    dn: int
    load: int
    flow_lpm: float
    velocity_mps: float
    limit_mps: float
    role: str
    basis: str = "table"

    def to_dict(self) -> dict[str, Any]:
        return {"dn": self.dn, "load": self.load,
                "flow_lpm": round(self.flow_lpm, 1),
                "velocity_mps": round(self.velocity_mps, 3),
                "limit_mps": self.limit_mps, "role": self.role,
                "basis": self.basis}


def _limit_of(role: str, limits: dict) -> float:
    # 2.5.3.3 단서는 호칭경이 아니라 배관의 역할로 가른다. 모듈 A 의
    # `classify_pipe_roles` 와 같은 두 갈래("branch" / 그 밖)를 쓴다.
    return float(limits["branch" if role == "branch" else "other"])


def size_pipes(loads: dict, roles: dict, constraints: dict,
               inner_dia_mm: Callable[[int], float],
               *, column: str = "가") -> tuple[dict, list[dict]]:
    """§9.5 — 별표1 최소를 잡고, 유속이 넘칠 때만 한 단계씩 올린다.

    `inner_dia_mm` 을 주입받는 이유는 내경이 관 재질에 따라 달라지고 그 재질이 DXF 에
    없어 설계 협의에서 정해지기 때문이다. 여기서 재질을 하나 골라 두면 협의 결과가
    바뀌어도 관경은 옛 내경으로 굳는다.

    150A 에서도 유속이 넘으면 더 올리지 않고 플래그로 돌려준다. 별표1 사다리 밖의
    호칭경을 지어내면 그 관경으로는 별표1 조회 자체가 되지 않는다.
    """
    limits = constraints["velocity_limit_mps"]
    flow_per_head = float(constraints["flow_lpm_min"])
    cross_min = int(constraints["cross_main_min_dn"])

    sizes: dict = {}
    flags: list[dict] = []
    for edge, load in loads.items():
        role = roles.get(edge, "other")
        limit = _limit_of(role, limits)
        flow = load * flow_per_head

        dn = T.min_dn(load, column=column)
        basis = "table"
        if role == "cross" and dn < cross_min:
            dn, basis = cross_min, "cross_main_min"

        v = velocity_mps(flow, inner_dia_mm(dn))
        while v > limit:
            nxt = T.next_dn(dn, column)
            if nxt is None:
                flags.append({
                    "code": "VELOCITY_EXCEEDED", "edge": [edge[0], edge[1]],
                    "dn": dn, "velocity_mps": round(v, 3), "limit_mps": limit,
                    "message": f"{edge[0]}→{edge[1]} 는 {dn}A 에서도 유속 "
                               f"{v:.2f} m/s 로 상한 {limit} m/s 를 넘습니다.",
                })
                break
            dn, basis = nxt, "velocity"
            v = velocity_mps(flow, inner_dia_mm(dn))

        sizes[edge] = PipeSize(dn=dn, load=load, flow_lpm=flow, velocity_mps=v,
                               limit_mps=limit, role=role, basis=basis)
    return sizes, flags


def enforce_monotonic(sizes: dict, source,
                      inner_dia_mm: Callable[[int], float]) -> list[dict]:
    """상류보다 굵은 하류를 상류에 맞춰 내린다(§13.3 R5). `sizes` 를 제자리에서 고친다.

    내린 결과가 2.5.3.3 단서의 유속 상한을 다시 넘길 수 있다. 가지배관 상한이 6 m/s
    로 그 밖의 배관 10 m/s 보다 좁아서, 유량이 더 작은 하류가 더 굵어질 수 있기
    때문이다. 명세는 그 경우를 적지 않았으므로 명세대로 내리되 조용히 넘기지 않는다.
    """
    children: dict = {}
    for up, down in sizes:
        children.setdefault(up, []).append(down)

    adjusted: list[dict] = []
    stack = [(source, None)]
    seen = {source}
    while stack:
        node, upstream_dn = stack.pop()
        for down in children.get(node, ()):
            size = sizes[(node, down)]
            if upstream_dn is not None and size.dn > upstream_dn:
                before = size.dn
                size.dn = upstream_dn
                size.basis = "monotonic"
                size.velocity_mps = velocity_mps(size.flow_lpm,
                                                 inner_dia_mm(size.dn))
                over = size.velocity_mps > size.limit_mps
                adjusted.append({
                    "code": "MONOTONIC_VELOCITY_CONFLICT" if over else "MONOTONIC_CLAMP",
                    "edge": [node, down], "from_dn": before, "to_dn": size.dn,
                    "velocity_mps": round(size.velocity_mps, 3),
                    "limit_mps": size.limit_mps,
                    "message": f"{node}→{down} 를 상류 {upstream_dn}A 에 맞춰 "
                               f"{before}A 에서 내렸습니다."
                               + (f" 유속이 {size.velocity_mps:.2f} m/s 로 상한 "
                                  f"{size.limit_mps} m/s 를 넘습니다." if over else ""),
                })
            if down in seen:
                continue
            seen.add(down)
            stack.append((down, size.dn))
    return adjusted
