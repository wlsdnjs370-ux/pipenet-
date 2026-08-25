"""
validation.py — 순수 검증 함수 모음 (도메인 계층)

PySide6·wntr 의존성 없는 순수 Python.
UI 메시지(QMessageBox) 표시는 호출자(UI 계층)가 담당.

함수 목록:
  - validate_topology_directions : 펌프·PRV·오리피스 배관 방향성 검사
  - check_isolation              : BFS 물리적 고립 구간 검사

출처:
  - validate_topology_directions ← infra/epanet_adapter.validate_topology
  - check_isolation              ← main.py._perform_pre_flight_check (BFS 부분)
"""
from __future__ import annotations

import collections
import logging
from typing import Callable

from .geometry import get_vector, angle_between_vectors
from .node_predicates import is_pump_type_text, is_water_tank_type_text

logger = logging.getLogger(__name__)


# ==============================================================================
# 배관 방향성 검증
# ==============================================================================

def validate_topology_directions(
    nodes: dict,
    pipe_registry: dict,
    get_nodes_by_type_id: Callable[[str], list[str]],
) -> dict:
    """펌프·PRV·오리피스 노드의 배관 방향성 검사.

    각 특수 노드에 대해 유입·유출 배관이 올바르게 연결됐는지 확인한다.

    Args:
        nodes               : {node_id: ...} — 노드 dict (존재 여부 확인용)
        pipe_registry       : {pipe_id: Pipe} — .start/.end 속성의 Pipe 객체 dict
        get_nodes_by_type_id: 타입 id 문자열을 받아 노드 id 목록 반환하는 callable

    Returns:
        {'valid': True}
        또는
        {'valid': False, 'target': str, 'title': str, 'msg': str}
    """
    degree_map: dict[str, dict[str, int]] = {
        n: {"in": 0, "out": 0, "total": 0} for n in nodes
    }

    for p in pipe_registry.values():
        s, e = p.start, p.end
        if s in degree_map:
            degree_map[s]["out"] += 1
            degree_map[s]["total"] += 1
        if e in degree_map:
            degree_map[e]["in"] += 1
            degree_map[e]["total"] += 1

    # 펌프 검사 — 승압(inline) 펌프는 유입·유출이 모두 있어야 함
    pumps = get_nodes_by_type_id("pump")
    for pid in pumps:
        deg = degree_map.get(pid, {"in": 0, "out": 0, "total": 0})
        if deg["total"] >= 2 and (deg["in"] == 0 or deg["out"] == 0):
            return {
                "valid": False,
                "target": pid,
                "title": "설계 오류 (승압펌프 방향)",
                "msg": (
                    f"⛔ <b>승압펌프({pid})의 유동 방향이 끊겨 있습니다.</b><br><br>"
                    f"배관이 양쪽에 연결되었으나, 물이 흐를 수 없는 구조입니다.<br>"
                    f"(유입: {deg['in']}개, 유출: {deg['out']}개)<br>"
                    f"배관 화살표 방향을 확인해주세요."
                ),
            }

    # PRV(감압변) 검사 — 반드시 중간 설치 (유입·유출 모두 존재)
    prvs = get_nodes_by_type_id("prv")
    for pid in prvs:
        deg = degree_map.get(pid, {"in": 0, "out": 0})
        if deg["in"] == 0 or deg["out"] == 0:
            return {
                "valid": False,
                "target": pid,
                "title": "설계 오류 (감압변 연결)",
                "msg": (
                    f"⛔ <b>감압변({pid})의 연결 상태가 올바르지 않습니다.</b><br><br>"
                    f"감압변은 반드시 <b>[유입 배관 → 감압변 → 유출 배관]</b> 순서로<br>"
                    f"중간에 설치되어야 합니다.<br><br>"
                    f"현재 상태: 유입 {deg['in']}개 / 유출 {deg['out']}개<br>"
                    f"(말단에 설치되었거나, 배관 방향이 거꾸로 되어 있습니다.)"
                ),
            }

    # 오리피스 검사 — 방향 무관, 최소 2개 이상의 배관 연결 필요
    orifices = get_nodes_by_type_id("orifice")
    for pid in orifices:
        deg = degree_map.get(pid, {"in": 0, "out": 0})
        total_deg = deg["in"] + deg["out"]
        if total_deg < 2:
            return {
                "valid": False,
                "target": pid,
                "title": "설계 오류 (오리피스 연결)",
                "msg": (
                    f"⛔ <b>오리피스({pid})의 연결 상태가 올바르지 않습니다.</b><br><br>"
                    f"오리피스는 배관 중간에 설치되어야 하므로 최소 2개의 배관이 연결되어야 합니다.<br><br>"
                    f"현재 상태: 총 연결 배관 {total_deg}개<br>"
                    f"(말단에 설치되어 있습니다.)"
                ),
            }

    return {"valid": True}


# ==============================================================================
# 물리적 고립 구간 검사
# ==============================================================================

def check_isolation(
    nodes: dict,
    pipe_registry: dict,
    node_types: dict,
    pump_roles: dict | None = None,
) -> dict:
    """수원 도달 불가 노드 BFS 검사 (물리적 고립 감지).

    Args:
        nodes        : {node_id: ...} — 전체 노드 dict
        pipe_registry: {pipe_id: Pipe} — .start/.end 속성의 Pipe 객체 dict
        node_types   : {node_id: str} — 노드 타입 문자열 dict
        pump_roles   : {pump_id: 'DEMAND'|'SUPPLY'} — 펌프 역할 지정 dict (선택)

    Returns:
        {'isolated': False}
        또는
        {'isolated': True, 'unreachable': list[str]}
    """
    if pump_roles is None:
        pump_roles = {}

    start_nodes: list[str] = []

    # (A) 물탱크(WT) 수집
    for n, t in node_types.items():
        if is_water_tank_type_text(t):
            start_nodes.append(n)

    # (B) DEMAND 펌프 수집 (지정된 경우), 없으면 전체 펌프
    has_demand = False
    if pump_roles:
        for pid, role in pump_roles.items():
            if role == "DEMAND":
                start_nodes.append(pid)
                has_demand = True

    if not has_demand:
        for n, t in node_types.items():
            if is_pump_type_text(t):
                start_nodes.append(n)

    if not start_nodes:
        return {"isolated": False}

    # 무방향 인접 리스트 빌드
    adj: dict[str, set[str]] = {n: set() for n in nodes}
    for p in pipe_registry.values():
        if p.start in adj and p.end in adj:
            adj[p.start].add(p.end)
            adj[p.end].add(p.start)

    # BFS 탐색
    visited: set[str] = set()
    queue: collections.deque = collections.deque()
    for n in start_nodes:
        if n in nodes:
            visited.add(n)
            queue.append(n)

    while queue:
        curr = queue.popleft()
        for neighbor in adj.get(curr, set()):
            if neighbor not in visited:
                visited.add(neighbor)
                queue.append(neighbor)

    unreachable = [n for n in nodes if n not in visited]
    if unreachable:
        return {"isolated": True, "unreachable": unreachable}

    return {"isolated": False}


# ==============================================================================
# 직렬 기기(밸브 / PRV / 오리피스 / 펌프 / 수조) DN·직선 설치 가드
# ==============================================================================

# 티·꺾임 자리에 설치를 막는 type_id 집합 (설치 적용·프리플라이트 공통)
SERIAL_COMPONENT_GUARD_TYPE_IDS: frozenset[str] = frozenset(
    {"valve", "prv", "orifice", "pump", "wt"}
)

_STRAIGHT_TOL_DEG: float = 1e-6  # 직선 판정 허용 오차 (degree) — 부동소수점 오차만 허용


def check_valve_pipe_guard(
    node_name: str,
    connected_pipes: list,
    get_coords: Callable[[str], "tuple[float, float, float] | None"],
    *,
    require_dn_match: bool = True,
    require_collinear: bool = True,
) -> dict:
    """직렬 기기(밸브 / PRV / 오리피스 / 펌프 / 수조) 설치 가능 여부 검사.

    Args:
        node_name          : 검사 대상 노드 ID
        connected_pipes    : 해당 노드에 연결된 Pipe 객체 목록 (.start/.end/.nominal_mm 필드 필요)
        get_coords         : node_id → (x, y, z) 또는 None 반환하는 callable
        require_dn_match   : True면 전후 배관 DN 일치 필수.
                             펌프는 흡입측을 토출보다 키우는 경우가 많아 False로 둔다.
        require_collinear  : True면 전후 배관이 3D 직선이어야 함.
                             펌프는 흡입·토출이 꺾이는(예: 90°) 경우가 많아 False로 둔다.

    Returns:
        dict with keys:
          ok         (bool)         — True 면 설치 허용
          dn_text    (str)          — UI 표시용 문자열
          reason     (str | None)   — 차단 이유 코드 (ok=False 일 때)
          nominal_mm (float | None) — 단일 DN 값 (orifice pipe_dn 자동갱신용)

    설치 허용 조건:
      - 0개(미연결)  : 허용 (편집 중 일시 상태)
      - 1개(말단)    : 허용
      - 2개          : (require_collinear 시 3D 직선) + (require_dn_match 시 DN 일치)
      - 3개 이상     : 항상 차단
    """
    count = len(connected_pipes)

    if count >= 3:
        return {"ok": False, "dn_text": "분기 연결 오류", "reason": "branch", "nominal_mm": None}

    if count == 0:
        return {"ok": True, "dn_text": "(미연결)", "reason": None, "nominal_mm": None}

    if count == 1:
        dn = connected_pipes[0].nominal_mm
        dn_text = f"DN{int(dn)}" if dn > 0 else "DN?(미설정)"
        return {"ok": True, "dn_text": dn_text, "reason": None, "nominal_mm": float(dn) if dn > 0 else None}

    # 2개 배관: (선택) DN 일치 + (선택) 직선 검사
    p1, p2 = connected_pipes[0], connected_pipes[1]
    dn1, dn2 = p1.nominal_mm, p2.nominal_mm

    if require_dn_match and (dn1 <= 0 or dn2 <= 0 or dn1 != dn2):
        t1 = f"DN{int(dn1)}" if dn1 > 0 else "DN?"
        t2 = f"DN{int(dn2)}" if dn2 > 0 else "DN?"
        return {"ok": False, "dn_text": f"{t1} / {t2} 불일치", "reason": "dn_mismatch", "nominal_mm": None}

    if require_collinear:
        nc = get_coords(node_name)
        other1 = p1.end if p1.start == node_name else p1.start
        other2 = p2.end if p2.start == node_name else p2.start
        c1 = get_coords(other1)
        c2 = get_coords(other2)

        if nc is None or c1 is None or c2 is None:
            return {"ok": False, "dn_text": "직선 연결 오류", "reason": "coord_missing", "nominal_mm": None}

        v1 = get_vector(nc, c1)
        v2 = get_vector(nc, c2)

        # 영벡터 방어 — normalize_vector 가 (0,0,0) 반환 시 acos(0)=90° 오판 방지
        if all(abs(x) < 1e-12 for x in v1) or all(abs(x) < 1e-12 for x in v2):
            return {"ok": False, "dn_text": "직선 연결 오류", "reason": "zero_length_pipe", "nominal_mm": None}

        try:
            angle = angle_between_vectors(v1, v2)
        except Exception:
            return {"ok": False, "dn_text": "직선 연결 오류", "reason": "angle_calc_failed", "nominal_mm": None}

        if abs(angle - 180.0) > _STRAIGHT_TOL_DEG:
            return {"ok": False, "dn_text": "직선 연결 오류", "reason": "not_collinear", "nominal_mm": None}

    if dn1 > 0 and dn2 > 0 and dn1 == dn2:
        dn = int(dn1)
        return {"ok": True, "dn_text": f"DN{dn}", "reason": None, "nominal_mm": float(dn)}

    t1 = f"DN{int(dn1)}" if dn1 > 0 else "DN?"
    t2 = f"DN{int(dn2)}" if dn2 > 0 else "DN?"
    return {"ok": True, "dn_text": f"{t1} / {t2}", "reason": None, "nominal_mm": None}
