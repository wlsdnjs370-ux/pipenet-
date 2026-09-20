"""
topology.py — 순수 토폴로지 분석 함수 모음 (도메인 계층)

PySide6·wntr 의존성 없는 순수 Python.

함수 목록:
  - analyze_connectivity   : BFS 활성 네트워크 분석 (수원→전체 노드 도달 가능성)
  - analyze_pump_topology  : 펌프 역할 분류 (source vs inline/booster)
  - get_neighbors          : 특정 노드의 인접 노드 목록
  - get_subtree_nodes      : BFS 서브트리 탐색 (무방향)
  - get_subtree_one_direction : 한 방향 BFS 서브트리 탐색

출처:
  - analyze_connectivity  ← infra/epanet_adapter._analyze_connectivity
  - analyze_pump_topology ← infra/epanet_adapter._build_base_network (펌프 분류)
                            + dialog_calc_setup.load_pumps (부스터 감지)
  - get_subtree_*         ← domain/node_manager (standalone 함수 버전)
"""
from __future__ import annotations

import collections
import logging
from typing import Any
from domain.constants import fitting_category_from_id, is_check_valve_category

logger = logging.getLogger(__name__)


def _obj_get(obj: Any, key: str, default: Any = None) -> Any:
    if isinstance(obj, dict):
        return obj.get(key, default)
    return getattr(obj, key, default)


def _safe_float(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return float(default)


# ==============================================================================
# 활성 네트워크 분석
# ==============================================================================

def analyze_connectivity(
    nodes: dict,
    pipe_registry: dict,
    node_types: dict,
    node_state_map: dict,
) -> set[str]:
    """활성 네트워크 BFS 분석: 수원·펌프·탱크에서 도달 가능한 노드 집합 반환.

    체크밸브 방향성과 닫힌 노드(is_closed)를 모두 고려한다.

    Args:
        nodes        : {node_id: (x, y, z)} — 노드 좌표 dict
        pipe_registry: {pipe_id: Pipe} — .start/.end 속성을 가진 Pipe 객체 dict
        node_types   : {node_id: str} — 노드 타입 문자열 dict (예: "PUMP", "WT")
        node_state_map: {node_id: dict} — 노드 상태 메타데이터 dict

    Returns:
        활성 노드 id 집합
    """
    # canonical category가 SSOT다. category 판정 불가일 때만 마이그레이션되지 않은
    # 레거시 플래그를 보존한다.
    check_valve_info: dict[str, str] = {}
    for n, m in node_state_map.items():
        if not isinstance(m, dict):
            continue
        category = fitting_category_from_id(m.get("category_id", ""))
        is_check_valve = (
            is_check_valve_category(category)
            if category is not None
            else bool(m.get("is_check_valve"))
        )
        if not is_check_valve:
            continue
        upstream = m.get("check_valve_direction", "")
        if upstream:
            check_valve_info[n] = upstream

    # 무방향 인접 리스트 빌드
    adj: dict[str, list[str]] = {n: [] for n in nodes}
    for p in pipe_registry.values():
        s, e = p.start, p.end
        if s in adj and e in adj:
            adj[s].append(e)
            adj[e].append(s)

    # 시작 노드: 펌프·수조·상수원 타입
    queue: collections.deque = collections.deque()
    active_nodes: set[str] = set()
    for n in nodes:
        t_str = str(node_types.get(n, "")).lower()
        if "pump" in t_str or "wt" in t_str or "res" in t_str or "source" in t_str:
            queue.append(n)
            active_nodes.add(n)

    while queue:
        curr = queue.popleft()

        if node_state_map.get(curr, {}).get("is_closed", False):
            continue

        for neighbor in adj[curr]:
            if neighbor in active_nodes:
                continue

            blocked = False

            # 현재 노드에 체크밸브 → 지정 업스트림 방향 역류 차단
            if curr in check_valve_info:
                if neighbor == check_valve_info[curr]:
                    blocked = True

            # 이웃 노드에 체크밸브 → 업스트림이 curr가 아니면 차단
            if not blocked and neighbor in check_valve_info:
                if curr != check_valve_info[neighbor]:
                    blocked = True

            if blocked:
                continue

            active_nodes.add(neighbor)
            queue.append(neighbor)

    return active_nodes


# ==============================================================================
# 펌프 토폴로지 분류
# ==============================================================================

def analyze_pump_topology(
    pump_nodes: list[str],
    pipe_registry: dict,
) -> dict[str, Any]:
    """펌프 노드를 source(수원) 펌프와 inline(부스터) 펌프로 분류한다.

    Args:
        pump_nodes   : 펌프 노드 id 목록 (editor.get_nodes_by_type_id("pump") 결과)
        pipe_registry: {pipe_id: Pipe} — .start/.end 속성을 가진 Pipe 객체 dict

    Returns:
        dict:
          - 'inline_pumps'    : set[str] — 양쪽 배관이 연결된 부스터/승압 펌프
          - 'source_pumps'    : set[str] — 수원 역할(토출만) 펌프
          - 'pump_degree'     : dict[str, int] — 펌프별 총 연결 배관 수
          - 'pump_has_incoming': dict[str, bool] — 유입 배관 존재 여부
    """
    pump_in_degree: dict[str, int] = {p: 0 for p in pump_nodes}
    pump_out_degree: dict[str, int] = {p: 0 for p in pump_nodes}
    pump_degree: dict[str, int] = {p: 0 for p in pump_nodes}
    pump_has_incoming: dict[str, bool] = {p: False for p in pump_nodes}

    for p_obj in pipe_registry.values():
        s, e = p_obj.start, p_obj.end
        if s in pump_out_degree:
            pump_out_degree[s] += 1
            pump_degree[s] += 1
        if e in pump_in_degree:
            pump_in_degree[e] += 1
            pump_degree[e] += 1
            pump_has_incoming[e] = True

    inline_pumps: set[str] = set()
    source_pumps: set[str] = set()

    for p_node in pump_nodes:
        if pump_in_degree.get(p_node, 0) >= 1 and pump_out_degree.get(p_node, 0) >= 1:
            inline_pumps.add(p_node)
        else:
            source_pumps.add(p_node)

    return {
        "inline_pumps": inline_pumps,
        "source_pumps": source_pumps,
        "pump_degree": pump_degree,
        "pump_has_incoming": pump_has_incoming,
    }


def _pump_churn_bar(node_obj: Any) -> float:
    """펌프 체절압[bar] — shutoff_p 우선, 없으면 정격압×1.2, 둘 다 없으면 0."""
    s = _safe_float(_obj_get(node_obj, "shutoff_p", 0.0))
    if s > 0.0:
        return s
    rp = _safe_float(_obj_get(node_obj, "rated_p", 0.0))
    return rp * 1.2 if rp > 0.0 else 0.0


def max_supply_pressure_at_node_bar(
    nodes: dict,
    pipe_registry: dict,
    target_id: str,
    *,
    bar_to_head_m: float = 10.197,  # 1 bar ≈ 10.197 m 수두 (엔진 BAR_TO_M과 일치)
) -> float | None:
    """target 노드에서 펌프가 낼 수 있는 최대 정압[bar, 게이지] 추정(무유량=체절 기준).

    = (병렬 수원펌프 최대 체절압 + 직렬 부스터 체절압 합) − (target표고 − 수원표고)/bar_to_head_m

    병렬 수원펌프는 압력이 더해지지 않으므로 가장 센 1개만, 직렬 부스터는 양정이 더해지므로
    합산한다(중간 표고는 상쇄돼 수원표고만 남는다). 체절 기준이라 마찰손실은 0.
    펌프가 없거나 체절압을 못 구하면 None(검사 불가). 부스터 합산은 후한 상한이라 경고용.
    """
    target = nodes.get(target_id) if hasattr(nodes, "get") else None
    if target is None:
        return None
    pump_ids = [nid for nid, n in nodes.items()
                if str(_obj_get(n, "type_id", "") or "").lower() == "pump"]
    if not pump_ids:
        return None
    topo = analyze_pump_topology(pump_ids, pipe_registry)
    sources = topo["source_pumps"]
    boosters = topo["inline_pumps"]
    if sources:
        base = max(sources, key=lambda p: _pump_churn_bar(nodes[p]))
        base_churn = _pump_churn_bar(nodes[base])
        ref_z = _safe_float(_obj_get(nodes[base], "elevation_m", 0.0))
    else:  # 수원펌프 없음(전부 인라인) — 합산하고 가장 낮은 펌프를 기준 표고로.
        base_churn = 0.0
        ref_z = min(_safe_float(_obj_get(nodes[p], "elevation_m", 0.0)) for p in pump_ids)
    total = base_churn + sum(_pump_churn_bar(nodes[p]) for p in boosters)
    if total <= 0.0:
        return None
    target_z = _safe_float(_obj_get(target, "elevation_m", 0.0))
    return total - (target_z - ref_z) / bar_to_head_m


def _is_hydraulic_source_node(node_obj: Any) -> bool:
    """수조·상수원 등 외부 수원 노드인지 판별한다."""
    if node_obj is None:
        return False
    type_id = str(
        _obj_get(node_obj, "type_id", "")
        or _obj_get(node_obj, "type", "")
        or ""
    ).strip().lower()
    markers = ("wt", "tank", "reservoir", "res", "source")
    if any(m in type_id for m in markers):
        return True
    if _safe_float(_obj_get(node_obj, "water_level", 0.0)) > 0.0:
        return True
    return False


def _pump_suction_upstream_reaches_source(
    pump_id: str,
    nodes_by_id: dict[str, Any],
    pipe_registry: dict,
) -> bool:
    """펌프 흡입측에서 배관 방향(start→end)을 거슬러 올라가며 수원 노드에 닿는지 본다.

    시나리오 JSON의 pipe는 유량 방향을 start→end로 두는 경우가 많아,
    펌프 직전 노드가 `base`여도 그 앞에 수조(wt)가 있으면 수원 흡입으로 본다.
    """
    pid = str(pump_id)
    seeds: list[str] = []
    for pipe_obj in pipe_registry.values():
        start = str(_obj_get(pipe_obj, "start", "") or "")
        end = str(_obj_get(pipe_obj, "end", "") or "")
        if end == pid and start:
            seeds.append(start)

    if not seeds:
        return False

    visited: set[str] = set()
    queue: collections.deque[str] = collections.deque(seeds)
    while queue:
        nid = queue.popleft()
        if nid in visited:
            continue
        visited.add(nid)
        node_obj = nodes_by_id.get(str(nid))
        if _is_hydraulic_source_node(node_obj):
            return True
        for pipe_obj in pipe_registry.values():
            start = str(_obj_get(pipe_obj, "start", "") or "")
            end = str(_obj_get(pipe_obj, "end", "") or "")
            if end == nid and start and start not in visited:
                queue.append(start)

    return False


def get_pump_display_context(
    pump_id: str,
    nodes_by_id: dict[str, Any],
    pipe_registry: dict,
    pump_roles: dict[str, str] | None = None,
    summary: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """펌프 계산 모드와 설치 역할을 분리해 화면/보고서용 표시 문구를 만든다.

    - 계산 모드(Demand/Supply)는 pump_roles를 우선 사용한다.
    - 부스터 여부는 압력이 아니라 토폴로지와 상류 공급원 타입으로 판단한다.
    - 수조/상수원/자연수두에서 흡입하는 펌프는 incoming 배관이 있어도 booster로 보지 않는다.
    """
    pid = str(pump_id)
    roles = pump_roles or {}
    summ = summary or {}

    calc_mode = str(roles.get(pid, "")).upper()
    if calc_mode not in {"DEMAND", "SUPPLY"}:
        source_node = str(summ.get("source_node", "")).upper()
        calc_mode = "DEMAND" if "HYBRID" in source_node else "SUPPLY"

    incoming_nodes: list[str] = []
    outgoing_nodes: list[str] = []
    for pipe_obj in pipe_registry.values():
        start = str(_obj_get(pipe_obj, "start", "") or "")
        end = str(_obj_get(pipe_obj, "end", "") or "")
        if start == pid:
            outgoing_nodes.append(end)
        if end == pid:
            incoming_nodes.append(start)

    has_upstream_source = _pump_suction_upstream_reaches_source(pid, nodes_by_id, pipe_registry)
    if not has_upstream_source:
        for upstream_id in incoming_nodes:
            upstream_obj = nodes_by_id.get(str(upstream_id))
            if _is_hydraulic_source_node(upstream_obj):
                has_upstream_source = True
                break

    is_booster = bool(incoming_nodes and outgoing_nodes and not has_upstream_source)

    if calc_mode == "DEMAND":
        desc_key = "pump.desc.demand"
    elif is_booster:
        desc_key = "pump.desc.booster"
    else:
        desc_key = "pump.desc.supply"

    return {
        "pump_id": pid,
        "calc_mode": calc_mode,
        "is_booster": is_booster,
        "desc_key": desc_key,
        "has_incoming": bool(incoming_nodes),
        "has_outgoing": bool(outgoing_nodes),
    }


def get_system_curve_static_head(
    static_head_bar: float,
    suction_pressure_bar: float,
    is_booster: bool,
) -> float:
    """부스터 펌프는 추가 양정 기준으로 시스템 곡선 시작점을 잡는다."""
    static_head = max(_safe_float(static_head_bar), 0.0)
    suction_head = max(_safe_float(suction_pressure_bar), 0.0)
    if is_booster:
        return max(static_head - suction_head, 0.0)
    return static_head


def should_show_discharge_overlay(
    pump_head_bar: float,
    discharge_pressure_bar: float,
    tolerance_bar: float = 0.05,
) -> bool:
    """토출점이 양정 운전점과 사실상 같은 경우 중복 표기를 숨긴다."""
    return abs(_safe_float(discharge_pressure_bar) - _safe_float(pump_head_bar)) > max(tolerance_bar, 0.0)


# ==============================================================================
# BFS 유틸리티 — 서브트리 탐색
# ==============================================================================

def get_neighbors(name: str, nodes: dict, pipe_registry: dict) -> list[str]:
    """주어진 노드의 인접 노드 목록을 반환한다 (무방향)."""
    if name not in nodes:
        return []
    neighbors: set[str] = set()
    for p in pipe_registry.values():
        if p.start == name:
            neighbors.add(p.end)
        elif p.end == name:
            neighbors.add(p.start)
    return sorted(neighbors)


def get_subtree_nodes(
    root: str,
    nodes: dict,
    pipe_registry: dict,
    blocked_neighbor: str | None = None,
) -> set[str]:
    """루트 노드에서 BFS로 연결된 모든 노드 집합을 반환한다.

    blocked_neighbor가 지정되면 root ↔ blocked_neighbor 간 경로를 차단한다.
    """
    if root not in nodes:
        return set()

    adj: dict[str, set[str]] = {}
    for p in pipe_registry.values():
        s, e = p.start, p.end
        adj.setdefault(s, set()).add(e)
        adj.setdefault(e, set()).add(s)

    visited: set[str] = {root}
    queue: collections.deque = collections.deque([root])

    while queue:
        n = queue.popleft()
        for nb in adj.get(n, set()):
            if blocked_neighbor is not None:
                if (n == root and nb == blocked_neighbor) or (
                    n == blocked_neighbor and nb == root
                ):
                    continue
            if nb not in visited:
                visited.add(nb)
                queue.append(nb)

    return visited


def get_subtree_one_direction(
    root: str,
    toward_neighbor: str | None,
    nodes: dict,
    pipe_registry: dict,
) -> set[str]:
    """루트 노드에서 특정 방향으로만 BFS 탐색한 노드 집합을 반환한다.

    toward_neighbor가 None이면 전방향 탐색(get_subtree_nodes와 동일).
    """
    if root not in nodes:
        return set()
    if toward_neighbor is None:
        return get_subtree_nodes(root, nodes, pipe_registry)

    adj: dict[str, set[str]] = {}
    for p in pipe_registry.values():
        s, e = p.start, p.end
        adj.setdefault(s, set()).add(e)
        adj.setdefault(e, set()).add(s)

    if toward_neighbor not in adj.get(root, set()):
        return get_subtree_nodes(root, nodes, pipe_registry)

    visited: set[str] = {root, toward_neighbor}
    queue: collections.deque = collections.deque([toward_neighbor])

    while queue:
        n = queue.popleft()
        for nb in adj.get(n, set()):
            if nb == root:
                continue
            if nb not in visited:
                visited.add(nb)
                queue.append(nb)

    return visited
