"""지시서 §9.6 C560 — 라우팅 결과를 위상 그래프로 확정하고 모듈 A 스키마로 방출한다.

C510~C550 이 남긴 것은 좌표 폴리라인이다. 폴리라인은 어디를 지나는지만 말하고 무엇이
무엇에 붙어 있는지는 말하지 않는다. 여기서 그것을 노드와 간선으로 확정한 뒤,
`core/remote30_full_network.CombinedTables` 와 **필드 단위로 같은** dict 를 낸다
(`docs/module_a_graph_schema.md` §6-1).

`CombinedTables` 를 직접 만들지 않고 dict 만 내는 이유는 그 모듈이 평면
import(`from remote30_constants import ...`)로 짜여 있어 `core/` 가 sys.path 에 올라야만
읽히기 때문이다. 설계 패키지가 그것을 끌고 들어오면 C1~C4 단위 시험까지 서버 경로에
묶인다. 조립은 라우트 계층이 한다.

지켜야 할 단위 (같은 dict 안에서 갈린다 — 스키마 문서 §6-2):
    nodes.x, nodes.y      mm
    nodes.elevation       m
    pipes.length, .elev   m
    pipes.dia             호칭경 mm (내경 아님)
"""

from __future__ import annotations

import math
from collections import deque
from dataclasses import dataclass, field
from typing import Any, Iterable

__all__ = [
    "Node", "Edge", "DesignTopology", "GraphTables",
    "build_topology", "reconstruct_fittings", "emit_tables",
]

_MM_PER_M = 1000.0

# 같은 점으로 볼 거리. 라우팅은 격자 위에서 돌아 좌표가 이미 정수 mm 라 반올림
# 오차만 흡수하면 된다. 격자 스냅으로 뭉개면 안 된다 — 붙지 않은 관이 붙는다.
NODE_EPS_MM = 1.0

# 꺾임 각도 구간. 모듈 A 의 `remote30_prototype.py:6141-6147` 과 같은 경계다.
COLLINEAR_TOL_DEG = 2.0
ELBOW_45_MIN_DEG, ELBOW_45_MAX_DEG = 43.5, 46.5
ELBOW_MIN_DEG = 70.0
# 이보다 크게 꺾이면 흡수하지 않고 노드로 남긴다(모듈 A `ELBOW_MERGE_TOL_DEG`).
ELBOW_MERGE_MAX_DEG = 95.0

DEFAULT_PIPE_TYPE = "KSD 3507"
DEFAULT_C_FACTOR = "120"

# 신축배관 규격 이름. 값 자체는 `zoning.FLEX_SPEC` 이 쥐고 있고 여기서는 하류
# 편집기가 아는 프로파일 키만 적는다. 모르는 키를 적으면 `override_flag` 없이는
# FX 검토 게이트가 막는다(`remote30_prototype.py:7458`).
FLEX_SPEC_REF = "한백표준"

_ROLE_RISER, _ROLE_MAIN, _ROLE_CROSS, _ROLE_BRANCH = "riser", "main", "cross", "branch"


@dataclass
class Node:
    """`elevation_m` 이 수리 실표고고 `display_z_m` 은 표시 전용이다.

    스키마 문서 §2.4 — 하향식 헤드의 수직 드롭은 표시 z 로만 존재한다. 그것을
    `elevation` 에 넣으면 낙차가 두 번 계상된다.
    """

    id: str
    x: float
    y: float
    elevation_m: float
    display_z_m: float
    kind: str
    floor: str | None = None
    head_id: str | None = None


@dataclass
class Edge:
    """`elbows` 는 흡수된 꺾임의 각도(deg)다. 흡수만 하고 버리면 등가길이가 샌다."""

    role: str
    length_m: float
    rise_m: float
    elbows: list[float] = field(default_factory=list)
    flex: dict | None = None


@dataclass
class DesignTopology:
    nodes: dict[str, Node] = field(default_factory=dict)
    edges: dict[tuple[str, str], Edge] = field(default_factory=dict)
    source: str | None = None
    head_nodes: dict[str, str] = field(default_factory=dict)
    flags: list[dict] = field(default_factory=list)

    def adjacency(self) -> dict[str, list[str]]:
        adj: dict[str, list[str]] = {n: [] for n in self.nodes}
        for a, b in self.edges:
            adj[a].append(b)
            adj[b].append(a)
        return adj


# ────────────────────────────────────────────────────────────────────────────
# 노드 색인
# ────────────────────────────────────────────────────────────────────────────

class _NodeIndex:
    """층별 격자 버킷으로 같은 점을 찾는다. 좌표는 뭉개지 않고 처음 값을 보존한다."""

    def __init__(self, topology: DesignTopology, eps_mm: float):
        self._topo = topology
        self._eps = eps_mm
        self._buckets: dict[tuple, list[str]] = {}
        self._seq = 0

    def _cells(self, floor, x, y):
        cx, cy = int(x // self._eps), int(y // self._eps)
        for dx in (-1, 0, 1):
            for dy in (-1, 0, 1):
                yield (floor, cx + dx, cy + dy)

    def find(self, floor, x, y) -> str | None:
        for cell in self._cells(floor, x, y):
            for nid in self._buckets.get(cell, ()):
                node = self._topo.nodes[nid]
                if math.hypot(node.x - x, node.y - y) <= self._eps:
                    return nid
        return None

    def add(self, floor, x, y, *, elevation_m, display_z_m, kind,
            node_id: str | None = None, head_id: str | None = None,
            distinct: bool = False) -> str:
        """`distinct` 는 색인을 건너뛴다 — 평면상 같은 자리에 두 노드를 두어야 할 때.

        하향식 헤드가 그렇다. 헤드는 제 급수구 바로 아래에 매달리므로 평면 좌표가
        급수구와 같다. 색인에 맡기면 둘이 한 노드로 합쳐지고 신축배관이 통째로
        사라진다 — 등가길이 22.4 m 가 소리 없이 빠진다.
        """
        found = None if distinct else self.find(floor, x, y)
        if found is not None:
            node = self._topo.nodes[found]
            # 헤드·밸브는 통과점보다 센 정체다. 같은 자리에 겹치면 그쪽으로 승격한다.
            if head_id and node.head_id is None:
                node.head_id, node.kind = head_id, kind
                self._topo.head_nodes[head_id] = found
            return found
        if node_id is None:
            self._seq += 1
            node_id = f"N{self._seq:04d}"
        self._topo.nodes[node_id] = Node(
            id=node_id, x=float(x), y=float(y), elevation_m=float(elevation_m),
            display_z_m=float(display_z_m), kind=kind, floor=floor, head_id=head_id)
        if not distinct:
            cell = (floor, int(x // self._eps), int(y // self._eps))
            self._buckets.setdefault(cell, []).append(node_id)
        if head_id:
            self._topo.head_nodes[head_id] = node_id
        return node_id


def _key(a: str, b: str) -> tuple[str, str]:
    """간선 키는 항상 정규화한다 (스키마 문서 §6-4)."""
    return (a, b) if a <= b else (b, a)


def _link(topo: DesignTopology, a: str, b: str, role: str, *,
          flex: dict | None = None) -> None:
    if a == b:
        return
    na, nb = topo.nodes[a], topo.nodes[b]
    # 도면상 길이는 3차원으로 잰다. 평면만 재면 입상관이 길이 0 이 된다 —
    # 같은 (x, y) 에 서서 표시 z 로만 올라가기 때문이다.
    drawn_m = math.hypot(nb.x - na.x, nb.y - na.y) / _MM_PER_M
    drawn_m = math.hypot(drawn_m, nb.display_z_m - na.display_z_m)
    rise = nb.elevation_m - na.elevation_m
    # §9.2 C540 — 배관장 = max(도면상 길이, |표고차|). 도면상 길이만 쓰면 눌러 그린
    # 층에서 낙차가 사라지고, 표고차만 쓰면 꺾여 올라간 관이 직선으로 줄어든다.
    length = max(drawn_m, abs(rise))
    if flex is not None:
        length = float(flex["physical_length_m"])
    topo.edges.setdefault(_key(a, b), Edge(role=role, length_m=length,
                                           rise_m=rise, flex=flex))


# ────────────────────────────────────────────────────────────────────────────
# C560 — 위상 확정
# ────────────────────────────────────────────────────────────────────────────

def _polyline(topo: DesignTopology, index: _NodeIndex, points: Iterable,
              *, floor, elevation_m, display_z_m, role: str) -> list[str]:
    ids = [index.add(floor, float(p[0]), float(p[1]), elevation_m=elevation_m,
                     display_z_m=display_z_m, kind=role) for p in points]
    for a, b in zip(ids, ids[1:]):
        _link(topo, a, b, role)
    return ids


def _param_on(node: Node, a: Node, b: Node, eps_mm: float) -> float | None:
    """점이 선분 **안쪽**에 닿아 있으면 그 매개변수를, 아니면 None."""
    dx, dy = b.x - a.x, b.y - a.y
    span = dx * dx + dy * dy
    if span <= 0.0:
        return None
    t = ((node.x - a.x) * dx + (node.y - a.y) * dy) / span
    if not 0.0 < t < 1.0:
        return None
    if math.hypot(a.x + t * dx - node.x, a.y + t * dy - node.y) > eps_mm:
        return None
    # 끝점 언저리는 색인이 이미 합쳤다. 여기서 또 자르면 길이 0 간선이 생긴다.
    if min(math.hypot(node.x - a.x, node.y - a.y),
           math.hypot(node.x - b.x, node.y - b.y)) <= eps_mm:
        return None
    return t


def _splice_junctions(topo: DesignTopology, eps_mm: float) -> int:
    """폴리라인 한복판에 떨어진 접속점을 그 간선에 끼워 넣는다.

    C530 은 곧게 이어지는 중간점을 지운다(`pipe_routing._simplify`). 분기점은
    교차배관 한복판에 서는 일이 대부분이라, 그대로 두면 분기점이 어디에도 붙지 않은
    섬으로 남고 그 아래 가지배관 전체가 급수원에서 끊긴다. 주배관 분기(`spurs`)가
    붙는 자리도 같다.
    """
    spliced = 0
    for key, edge in list(topo.edges.items()):
        na, nb = topo.nodes[key[0]], topo.nodes[key[1]]
        if edge.flex is not None or na.floor != nb.floor:
            continue
        inner = sorted(
            (t, nid) for nid, node in topo.nodes.items()
            if nid not in key and node.floor == na.floor
            for t in (_param_on(node, na, nb, eps_mm),) if t is not None)
        if not inner:
            continue
        del topo.edges[key]
        chain = [key[0]] + [nid for _, nid in inner] + [key[1]]
        for u, v in zip(chain, chain[1:]):
            _link(topo, u, v, edge.role)
        spliced += len(inner)
    return spliced


def build_topology(riser, zones: list[dict], heads: Iterable,
                   connections: Iterable, *, eps_mm: float = NODE_EPS_MM
                   ) -> DesignTopology:
    """라우팅 산출물을 노드·간선으로 확정한다.

    `zones` 한 항목 = `{"plan": ZonePlan, "floor": "1F"}`. 층 표고는 여기서 구하지
    않고 입상관 층 노드에서 읽는다. 표고를 두 곳에서 구하면 한쪽만 고쳐졌을 때
    주배관과 입상관이 다른 높이에 서고, 그 차이가 그대로 낙차 압력이 된다.

    입상관에 없는 층의 구역은 채우지 않고 플래그로 돌려준다(금지사항 G9).
    """
    topo = DesignTopology()
    index = _NodeIndex(topo, eps_mm)

    levels = {level.floor: level for level in riser.ordered}
    riser_ids: dict[str, str] = {}
    for level in riser.ordered:
        riser_ids[level.floor] = index.add(
            level.floor, riser.point[0], riser.point[1],
            elevation_m=level.elevation_m, display_z_m=level.display_z_m,
            kind="riser", node_id=riser.node_id(level))
    for lo, hi in zip(riser.ordered, riser.ordered[1:]):
        _link(topo, riser_ids[lo.floor], riser_ids[hi.floor], _ROLE_RISER)

    topo.source = riser.source_node
    if topo.source is None:
        topo.flags.append({
            "code": "RISER_EMPTY",
            "message": "입상관에 층이 하나도 없어 급수원을 정할 수 없습니다.",
        })
        return topo

    head_xy = {h.id if hasattr(h, "id") else h["id"]:
               (float(h.x if hasattr(h, "x") else h["x"]),
                float(h.y if hasattr(h, "y") else h["y"])) for h in heads}
    conn_by_head = {c["head_id"]: c for c in connections}

    for entry in zones:
        plan, floor = entry["plan"], entry["floor"]
        level = levels.get(floor)
        if level is None:
            plan_flags = [h for b in plan.branches for h in b.heads]
            topo.flags.append({
                "code": "RISER_LEVEL_MISSING", "zone_id": plan.zone_id,
                "floor": floor, "heads": plan_flags,
                "message": f"{floor} 의 입상관 층 노드가 없어 구역 {plan.zone_id} 를 "
                           f"급수원에 이을 수 없습니다. 헤드 {len(plan_flags)}개가 "
                           f"미방호입니다.",
            })
            continue
        elev, disp = level.elevation_m, level.display_z_m
        geom = dict(floor=floor, elevation_m=elev, display_z_m=disp)

        main = plan.main
        if main is None or not main.path:
            topo.flags.append({
                "code": "MAIN_MISSING", "zone_id": plan.zone_id,
                "message": f"구역 {plan.zone_id} 에 주배관 경로가 없습니다.",
            })
            continue
        valve_id = index.add(floor, main.valve[0], main.valve[1],
                             elevation_m=elev, display_z_m=disp, kind="valve",
                             node_id=f"AV-{plan.zone_id}")
        _link(topo, riser_ids[floor], valve_id, _ROLE_MAIN)
        _polyline(topo, index, main.path, role=_ROLE_MAIN, **geom)
        for spur in main.spurs:
            _polyline(topo, index, spur, role=_ROLE_MAIN, **geom)

        for cross in plan.crosses:
            _polyline(topo, index, cross.path, role=_ROLE_CROSS, **geom)
            for spur in cross.spurs:
                _polyline(topo, index, spur, role=_ROLE_CROSS, **geom)

        for branch in plan.branches:
            _build_branch(topo, index, plan, branch, head_xy, conn_by_head, geom)

    _splice_junctions(topo, eps_mm)
    return topo


def _build_branch(topo, index, plan, branch, head_xy, conn_by_head, geom) -> None:
    """분기점에서 양쪽으로 뻗는 가지배관. 한쪽씩 따로 이어야 빗살이 된다.

    양쪽 헤드를 한 줄로 묶어 이으면 분기점을 지나치는 간선이 생겨 담당 헤드 수가
    한쪽으로 몰린다.
    """
    axes = plan.axes
    tee_a, tee_b = axes.project(*branch.tee)
    tee_id = index.add(geom["floor"], branch.tee[0], branch.tee[1],
                       elevation_m=geom["elevation_m"],
                       display_z_m=geom["display_z_m"], kind=_ROLE_BRANCH)

    for side in (branch.left, branch.right):
        ordered = sorted((h for h in side if h in head_xy),
                         key=lambda h: abs(axes.project(*head_xy[h])[1] - tee_b))
        upstream = tee_id
        for head_id in ordered:
            hx, hy = head_xy[head_id]
            # 가지배관은 축을 따라 곧게 간다. 헤드 접속점은 그 선 위의 발이다.
            feed = axes.world(tee_a, axes.project(hx, hy)[1])
            conn = conn_by_head.get(head_id)
            if conn is None:
                topo.flags.append({
                    "code": "HEAD_CONNECTION_MISSING", "head_id": head_id,
                    "zone_id": plan.zone_id,
                    "message": f"{head_id} 의 헤드 접속 방식(신축배관/직결)이 "
                               f"정해지지 않았습니다.",
                })
                continue
            if conn["kind"] == "direct":
                # 상향식은 가지배관에 직결한다 — 접속점이 곧 헤드다.
                node = index.add(geom["floor"], feed[0], feed[1],
                                 elevation_m=geom["elevation_m"],
                                 display_z_m=geom["display_z_m"],
                                 kind="head", head_id=head_id)
                _link(topo, upstream, node, _ROLE_BRANCH)
                upstream = node
                continue
            feed_id = index.add(geom["floor"], feed[0], feed[1],
                                elevation_m=geom["elevation_m"],
                                display_z_m=geom["display_z_m"], kind=_ROLE_BRANCH)
            _link(topo, upstream, feed_id, _ROLE_BRANCH)
            # 스키마 §2.4 — 매달린 만큼은 표시 z 로만 내린다. `elevation` 에 넣으면
            # 그 낙차가 수리계산에서 한 번 더 계상된다.
            head_node = index.add(geom["floor"], hx, hy,
                                  elevation_m=geom["elevation_m"],
                                  display_z_m=geom["display_z_m"]
                                  - float(conn["flex"]["physical_length_m"]),
                                  kind="head", head_id=head_id,
                                  node_id=f"HD-{head_id}", distinct=True)
            _link(topo, feed_id, head_node, _ROLE_BRANCH, flex=conn["flex"])
            upstream = feed_id


# ────────────────────────────────────────────────────────────────────────────
# C560 — 꺾임 흡수
# ────────────────────────────────────────────────────────────────────────────

def _turn_deg(a: Node, mid: Node, b: Node) -> float:
    v1 = math.atan2(mid.y - a.y, mid.x - a.x)
    v2 = math.atan2(b.y - mid.y, b.x - mid.x)
    return math.degrees(abs(((v2 - v1 + math.pi) % (2 * math.pi)) - math.pi))


def absorb_bends(topo: DesignTopology) -> int:
    """차수 2 통과점을 간선 하나로 합치고 꺾임 각도를 그 간선에 적어 둔다.

    합치지 않으면 모든 폴리라인 꼭짓점이 노드로 남아, 같은 형상인데도 모듈 A 가
    낸 그래프보다 노드가 몇 배로 많아진다. 두 파이프라인의 산출 그래프가 같은
    자료구조여야 한다는 §0.1 이 그래서 깨진다.

    헤드·밸브·입상관, 그리고 역할이 갈리는 지점은 합치지 않는다. 역할이 섞이면
    유속 상한이 어느 쪽 기준인지 말할 수 없게 된다.
    """
    merged = 0
    while True:
        adj = topo.adjacency()
        target = None
        for nid, nbrs in adj.items():
            node = topo.nodes[nid]
            if len(nbrs) != 2 or node.kind in ("head", "valve", "riser"):
                continue
            if node.head_id or nid == topo.source:
                continue
            a, b = nbrs
            ea, eb = topo.edges[_key(a, nid)], topo.edges[_key(nid, b)]
            if ea.role != eb.role or ea.flex or eb.flex:
                continue
            if _key(a, b) in topo.edges:
                continue  # 삼각형을 닫으면 트리가 아니게 된다.
            angle = _turn_deg(topo.nodes[a], node, topo.nodes[b])
            if angle > ELBOW_MERGE_MAX_DEG:
                continue
            target = (nid, a, b, ea, eb, angle)
            break
        if target is None:
            return merged

        nid, a, b, ea, eb, angle = target
        elbows = list(ea.elbows) + list(eb.elbows)
        if angle > COLLINEAR_TOL_DEG:
            elbows.append(angle)
        del topo.edges[_key(a, nid)]
        del topo.edges[_key(nid, b)]
        del topo.nodes[nid]
        topo.edges[_key(a, b)] = Edge(
            role=ea.role, length_m=ea.length_m + eb.length_m,
            rise_m=ea.rise_m + eb.rise_m, elbows=elbows)
        merged += 1


# ────────────────────────────────────────────────────────────────────────────
# C560 — 방출
# ────────────────────────────────────────────────────────────────────────────

@dataclass
class GraphTables:
    """`CombinedTables` 와 필드 이름이 같은 순수 dict 묶음."""

    nodes: list[dict] = field(default_factory=list)
    pipes: list[dict] = field(default_factory=list)
    nozzles: list[dict] = field(default_factory=list)
    fittings: list[dict] = field(default_factory=list)
    equipment: list[dict] = field(default_factory=list)
    meta: list[tuple[str, str]] = field(default_factory=list)
    flags: list[dict] = field(default_factory=list)
    labels: dict[str, str] = field(default_factory=dict)

    def to_dict(self) -> dict[str, Any]:
        return {"nodes": self.nodes, "pipes": self.pipes,
                "nozzles": self.nozzles, "fittings": self.fittings,
                "equipment": self.equipment, "meta": [list(m) for m in self.meta],
                "flags": self.flags}


def orient(topo: DesignTopology) -> list[tuple[str, str]]:
    """급수원에서 물이 흐르는 방향으로 간선을 세운다.

    표를 위에서 아래로 읽으면 흐르는 순서가 되고, 트리에 못 들어간 간선이 곧
    "길이 잘못 트인" 후보로 드러난다(모듈 A 와 같은 취급).
    """
    adj = topo.adjacency()
    seen = {topo.source}
    ordered: list[tuple[str, str]] = []
    queue = deque([topo.source])
    while queue:
        node = queue.popleft()
        for nxt in adj.get(node, ()):
            if nxt in seen:
                continue
            seen.add(nxt)
            ordered.append((node, nxt))
            queue.append(nxt)

    in_tree = {_key(u, v) for u, v in ordered}
    off_tree = [k for k in topo.edges
                if k[0] in seen and k[1] in seen and k not in in_tree]
    if off_tree:
        topo.flags.append({
            "code": "OFF_TREE_EDGE", "edges": [list(k) for k in off_tree],
            "message": f"급수원에서 뻗은 트리에 들지 않는 간선이 "
                       f"{len(off_tree)}개 있습니다 — 배관망에 고리가 있습니다.",
        })
    unreached = [n for n in topo.nodes if n not in seen]
    if unreached:
        topo.flags.append({
            "code": "NODE_UNREACHABLE", "nodes": unreached,
            "message": f"급수원에서 닿지 않는 노드가 {len(unreached)}개 있습니다.",
        })
    return ordered


def reconstruct_fittings(pipes: list[dict], elbows_by_pipe: dict) -> list[dict]:
    """부속류를 원본 표기에서 승계하지 않고 위상에서 다시 만든다(§9.6).

    [문서정합 §"C560 직전"] 명세는 tee 를 "방향벡터 최대각이 최소인 관로를 분기로
    판정" 이라 적었고 "판정 로직은 모듈 A 와 같은 함수를 써야 한다" 고 했다. 모듈 A
    의 실제 판정은 각도를 보지 않는다 — `remote30_prototype.py:6160-6165` 는
    **파이프의 `in` 노드 차수 ≥ 3** 하나만 본다(스키마 문서 §4.2). 게다가 그 코드는
    3천 줄짜리 함수 안에 인라인이라 부를 수 있는 함수가 아니다(모듈 A 수정 금지, G8).
    같은 형상에 같은 부속류가 붙어야 왕복이 맞으므로 **모듈 A 의 실제 규칙**을 옮겼다.

    [문서정합] 모듈 A 는 짧은 두 세그먼트(각 500mm 내외)에서만 꺾임을 부속류로
    적는다(`:5000-5001`). 그것은 추출에서 잡음 꺾임을 거르는 장치다. 순방향 설계는
    꺾은 사람이 우리라 모든 꺾임이 실물이므로 길이 조건을 걸지 않는다. 걸면 긴
    교차배관의 엘보가 통째로 사라지고 그만큼의 등가길이가 조용히 빠진다.
    """
    out: list[dict] = []
    for pipe in pipes:
        for angle in elbows_by_pipe.get(pipe["label"], ()):
            if ELBOW_45_MIN_DEG <= angle <= ELBOW_45_MAX_DEG:
                ftype = "elbow-45"
            elif angle >= ELBOW_MIN_DEG:
                ftype = "elbow"
            else:
                continue
            out.append({"pipe": pipe["label"], "in": pipe["in"],
                        "out": pipe["out"], "type": ftype, "count": "1"})

    degree: dict[str, int] = {}
    for pipe in pipes:
        degree[pipe["in"]] = degree.get(pipe["in"], 0) + 1
        degree[pipe["out"]] = degree.get(pipe["out"], 0) + 1
    for pipe in pipes:
        if degree.get(pipe["in"], 0) >= 3:
            out.append({"pipe": pipe["label"], "in": pipe["in"],
                        "out": pipe["out"], "type": "tee", "count": "1"})
    return out


def _unbinned(elbows_by_pipe: dict) -> list[float]:
    return [a for angles in elbows_by_pipe.values() for a in angles
            if a > COLLINEAR_TOL_DEG
            and not (ELBOW_45_MIN_DEG <= a <= ELBOW_45_MAX_DEG)
            and a < ELBOW_MIN_DEG]


def emit_tables(topo: DesignTopology, sizes: dict, constraints: dict, *,
                pipe_type: str = DEFAULT_PIPE_TYPE,
                c_factor: str = DEFAULT_C_FACTOR,
                meta: Iterable[tuple[str, str]] = ()) -> GraphTables:
    """확정된 위상 + 관경 → 모듈 A 스키마 dict.

    라벨은 모듈 A 와 같이 급수원 10 부터 흐르는 순으로 매긴다. 설계 쪽 id 는
    `design_id` 로 함께 실어 화면에서 되짚을 수 있게 둔다.
    """
    if topo.source is None:
        return GraphTables(flags=list(topo.flags))

    # `orient` 이 고리와 미도달 노드를 topo.flags 에 적는다. 표를 먼저 만들면 그
    # 두 플래그가 표에 실리지 않아 화면까지 오지 못한다.
    directed = orient(topo)
    tables = GraphTables(flags=list(topo.flags))
    counter = 10
    for node_id in [topo.source] + [b for _, b in directed]:
        if node_id not in tables.labels:
            tables.labels[node_id] = str(counter)
            counter += 1

    for node_id, label in tables.labels.items():
        node = topo.nodes[node_id]
        tables.nodes.append({
            "label": label, "elevation": round(node.elevation_m, 4),
            "io_node": "Input" if node_id == topo.source else "No",
            "x": int(round(node.x)), "y": int(round(node.y)),
            "display_z_m": round(node.display_z_m, 4),
            "design_id": node_id, "kind": node.kind,
        })

    elbows_by_pipe: dict[str, list[float]] = {}
    flex_pipes: list[tuple[dict, dict]] = []
    missing_size: list[list[str]] = []
    for seq, (up, down) in enumerate(directed, start=1):
        edge = topo.edges[_key(up, down)]
        # 배관 라벨은 순수 숫자여야 한다. `"P1"` 같은 접두어를 붙이면 모듈 A 의 FX
        # 실배관(`_materialize_fx_pipes` 가 1 부터 붙인다)과 SDF 재파싱 단계에서
        # 같은 id `P1` 로 뭉개진다 — 배관 8 개가 조용히 사라졌다.
        label = str(seq)
        size = sizes.get((up, down)) or sizes.get(_key(up, down))
        if edge.flex is not None:
            dn = int(edge.flex["dn"])
        elif size is None:
            missing_size.append([up, down])
            continue
        else:
            dn = int(size.dn if hasattr(size, "dn") else size["dn"])
        pipe = {
            "label": label,
            "in": tables.labels[up], "out": tables.labels[down],
            "type": pipe_type, "dia": dn,
            # 길이 0 인 관은 솔버가 특이행렬로 거부한다. |낙차| 아래로도 못 내린다.
            "length": max(round(edge.length_m, 3), 0.01, abs(round(edge.rise_m, 3))),
            "elev": round(edge.rise_m, 3),
            "c": str(edge.flex["c_factor"]) if edge.flex else c_factor,
            "status": "Normal", "group": "Unset",
            "role": edge.role,
        }
        tables.pipes.append(pipe)
        if edge.elbows:
            elbows_by_pipe[label] = edge.elbows
        if edge.flex is not None:
            flex_pipes.append((pipe, edge.flex))

    if missing_size:
        tables.flags.append({
            "code": "PIPE_SIZE_MISSING", "edges": missing_size,
            "message": f"관경이 정해지지 않은 간선이 {len(missing_size)}개 있습니다 "
                       f"— 그 구간은 방출에서 빠졌습니다.",
        })

    flow_lpm = float(constraints["flow_lpm_min"])
    for i, (head_id, node_id) in enumerate(sorted(topo.head_nodes.items()), start=1):
        if node_id not in tables.labels:
            continue
        tables.nozzles.append({
            "label": str(i), "in": tables.labels[node_id], "out": f"@/{i}",
            "status": "1", "lib": "SP-HEAD", "head_id": head_id,
            # m³/s 는 L/min 에서 유도한다 — 손으로 자른 상수를 쓰면 되돌릴 때 어긋난다.
            "flow_lmin": flow_lpm, "flow_m3s": flow_lpm / 60000.0,
        })

    for i, (pipe, flex) in enumerate(flex_pipes, start=1):
        tables.equipment.append({
            "pipe": pipe["label"], "in": pipe["in"], "out": pipe["out"],
            "label": str(i), "desc": "FX",
            # 등가길이는 형식승인 기준의 고정값이다. 도면 물리길이를 여기 쓰면 안 된다.
            "eq_len": float(flex["equivalent_length_m"]), "rel_pos": 0.5,
            "spec_ref": FLEX_SPEC_REF, "source": "designed",
            "override_flag": False, "override_note": "",
            "drawing_len_mm": round(float(flex["physical_length_m"]) * _MM_PER_M, 1),
        })

    tables.fittings = reconstruct_fittings(tables.pipes, elbows_by_pipe)
    stray = _unbinned(elbows_by_pipe)
    if stray:
        tables.flags.append({
            "code": "BEND_NOT_A_FITTING", "angles": [round(a, 1) for a in stray],
            "message": f"부속류 각도 구간 어디에도 들지 않는 꺾임이 {len(stray)}개 "
                       f"있습니다 — 등가길이에 반영되지 않았습니다.",
        })

    tables.meta = list(meta)
    return tables
