"""
domain/node_manager.py

노드 관련 비즈니스 로직을 담당하는 매니저 클래스.
PySide6 의존성 없음 — 순수 도메인 계층.

[DR-2 전환 완료]
과거 메타 딕셔너리 참조를 graph.get_node() 필드 접근으로 전환하였다.
"""
from __future__ import annotations

import copy
import logging
from typing import TYPE_CHECKING, Callable

logger = logging.getLogger(__name__)

from .component_specs import component_spec_owned_field_names
from .constants import fitting_category_from_id, is_check_valve_category
from .models import Node, NodeTypeId, normalize_node_type_id, _NODE_TYPE_LOOKUP
from .geometry import snap_point, is_close, dist_3d, get_vector, angle_between_vectors

if TYPE_CHECKING:
    from .network_graph import NetworkGraph


# ------------------------------------------------------------------
# 모듈 상수 — sanitize_node_data 전용
# ------------------------------------------------------------------

# 타입별로 '정리(purge)' 대상에서 제외할 Node 필드 집합.
# ComponentSpec 소유 필드 + category_id만 보존한다. is_active/is_closed 같은
# 공통 상태는 _PURGEABLE_FIELDS 대상이 아니므로 이 맵에서 소유하지 않는다.
def _keep_fields_for_type(type_id: str) -> set[str]:
    fields = set(component_spec_owned_field_names(type_id))
    if type_id and type_id != "base":
        fields.add("category_id")
    return fields


_KEEP_MAP: dict[str, set[str]] = {
    "head": _keep_fields_for_type("head"),
    "nozzle": _keep_fields_for_type("nozzle"),
    "pump": _keep_fields_for_type("pump"),
    "wt": _keep_fields_for_type("wt"),
    "tank": _keep_fields_for_type("wt"),
    "prv": _keep_fields_for_type("prv"),
    "orifice": _keep_fields_for_type("orifice"),
    "valve": _keep_fields_for_type("valve"),
    "hose_stream": _keep_fields_for_type("hose_stream"),
}

# 정리 대상 필드 전체 목록과 기본값
_PURGEABLE_FIELDS: dict[str, object] = {
    "category_id":           "",
    "k_factor_si":           None,
    "head_spec_name":        None,
    "nozzle_name":           None,
    "required_pressure_bar": 0.0,
    "fitting_id":            None,
    "pressure_setting_bar":  None,
    "loss_coefficient":      0.0,
    "hole_diameter":         0.0,
    "orifice_discharge_coeff": 0.0,
    "pipe_dn":               0.0,
    "rated_q":               0.0,
    "rated_p":               0.0,
    "shutoff_p":             0.0,
    "peak_q":                0.0,
    "peak_p":                0.0,
    "pump_library_id":       None,
    "pump_curve_data":       None,   # list는 각 호출마다 새 객체 생성
    "water_level":              0.0,
    "has_check_valve":          True,
    "hose_stream_flow_lps":     400.0 / 60.0,
}

# 소문자 내부 ID → 사람이 읽기 쉬운 표시명 매핑
# editor_core.py 직접 배치 경로에서 소문자로 저장되는 케이스에 대한 폴백
_TYPE_ID_TO_DISPLAY: dict[str, str] = {
    "head":    "Head",
    "nozzle":  "Nozzle",
    "pump":    "Pump",
    "prv":     "Prv",
    "orifice": "Orifice",
    "wt":      "WT",
    "res":     "Res",
    "valve":        "Valve",
    "hose_stream":  "Hose Stream",
    "base":         "",
}


class NodeManager:
    """
    노드 CRUD, 타입 관리, 서브트리 탐색 등 노드 관련 로직 전담.

    [DR-2 이후] 모든 읽기/쓰기는 NetworkGraph → Node 객체를 통한다.
    """

    def __init__(
        self,
        node_counter: dict,
        graph: NetworkGraph,
        grid_step: float = 0.01,
        type_id_lookup: Callable[[str], str | None] | None = None,
    ):
        self.node_counter = node_counter
        self.graph = graph
        self.grid_step = grid_step
        self.snap_tol = grid_step * 0.45
        self._type_id_lookup = type_id_lookup

    # ------------------------------------------------------------------
    # CRUD
    # ------------------------------------------------------------------

    def add_node(
        self,
        name: str,
        x: float,
        y: float,
        z: float,
        node_type: str = "기본",
    ) -> None:
        x, y, z = snap_point((float(x), float(y), float(z)), self.grid_step)
        if self.graph.has_node(name):
            self.graph.update_node(
                name,
                coords=(x, y, z),
                elevation_m=float(z),
                type=str(node_type),
            )
        else:
            self.graph.add_node(Node.create(name, x, y, z, str(node_type)))
        self._refresh_node_views_from_graph(name)
        logger.debug("[노드 생성] %s = (%s, %s, %s), type=%s", name, x, y, z, node_type)

    def delete_node(self, name: str, pipe_registry: dict) -> bool:
        """
        노드 삭제. 연결된 배관이 2개 이상이면 삭제 거부.
        반환: 삭제 성공 여부
        """
        if not self.graph.has_node(name):
            return False

        connected_pids = [
            pid for pid, p in pipe_registry.items()
            if p.start == name or p.end == name
        ]

        if len(connected_pids) >= 2:
            return False

        self.graph.remove_node(name)
        self._refresh_node_views_from_graph(name)  # node_type_ids 등 캐시 정리

        return True

    def rename_node(self, old_name: str, new_name: str, pipe_registry: dict) -> None:
        if old_name == new_name:
            return
        if self.graph.has_node(new_name):
            raise ValueError(f"노드 이름 {new_name} 이(가) 이미 존재합니다.")
        if not self.graph.has_node(old_name):
            return

        node = self.graph.get_node(old_name)
        if node is None:
            return
        renamed = copy.deepcopy(node)
        renamed.id = new_name
        self.graph.add_node(renamed)

        for pid, p in self.graph.pipes.items():
            updates = {}
            if p.start == old_name:
                updates["start"] = new_name
            if p.end == old_name:
                updates["end"] = new_name
            if updates:
                self.graph.update_pipe(pid, **updates)

        self.graph.invalidate_indices()

        # check_valve_direction은 Node 필드 → graph에서 직접 갱신
        for n, node_obj in self.graph.nodes.items():
            if node_obj.check_valve_direction == old_name:
                self.graph.update_node(n, check_valve_direction=new_name)

        self.graph.remove_node(old_name)
        self._refresh_node_views_from_graph(old_name)
        self._refresh_node_views_from_graph(new_name)

        logger.debug("[노드 이름 변경] %s -> %s", old_name, new_name)

    def find_node_at_position(
        self, x: float, y: float, z: float, tol: float | None = None
    ) -> str | None:
        if tol is None:
            tol = self.snap_tol
        x, y, z = float(x), float(y), float(z)
        for name, node in self.graph.nodes.items():
            nx, ny, nz = node.coords
            if abs(nx - x) <= tol and abs(ny - y) <= tol and abs(nz - z) <= tol:
                return name
        return None

    def generate_node_name(self, kind: str = "N") -> str:
        if "N" not in self.node_counter:
            self.node_counter["N"] = 0
        self.node_counter["N"] += 1
        return f"N{self.node_counter['N']}"

    # ------------------------------------------------------------------
    # 타입 관리
    # ------------------------------------------------------------------

    def get_node_type_id(self, node_name: str) -> str:
        node = self.graph.get_node(node_name)
        if node is None:
            return NodeTypeId.BASE.value

        if node.type_id:
            return str(node.type_id)

        self._infer_and_set_type_id(node_name)
        node = self.graph.get_node(node_name)
        return str(node.type_id) if node and node.type_id else NodeTypeId.BASE.value

    def set_node_type_id(
        self, node_name: str, type_id: str, display_name: str | None = None
    ) -> None:
        if not self.graph.has_node(node_name):
            return

        patch: dict = {"type_id": str(type_id)}
        if display_name is not None:
            patch["type"] = str(display_name)

        self.graph.update_node(node_name, **patch)
        self._refresh_node_views_from_graph(node_name)

    def get_nodes_by_type_id(self, type_id: str) -> list[str]:
        tid = str(type_id)
        return sorted(
            node_name
            for node_name in self.graph.nodes.keys()
            if self.get_node_type_id(node_name) == tid
        )

    def _infer_and_set_type_id(self, node_name: str) -> None:
        node = self.graph.get_node(node_name)
        if node is None:
            return
        display = str(node.type or "")
        # 표시명 직접 추론 우선 → 불가 시 필드값 기반 추론 폴백
        tid = self._type_id_from_display(display) or normalize_node_type_id(display, node.to_dict())
        self.graph.update_node(node_name, type_id=tid)
        self._refresh_node_views_from_graph(node_name)

    def sync_type_ids(self, *, verbose: bool = False) -> int:
        changed = 0
        for name, node in self.graph.nodes.items():
            disp = str(node.type or "")
            # 표시명 직접 추론 우선 → 불가 시 필드값 기반 추론 폴백
            inferred = self._type_id_from_display(disp) or normalize_node_type_id(disp, node.to_dict())

            if node.type_id != inferred:
                self.graph.update_node(name, type_id=inferred)
                changed += 1
                if verbose:
                    logger.debug("[SYNC] %s: type_id -> %s (display=%s)", name, inferred, disp)

        return changed

    # ------------------------------------------------------------------
    # 표시용 헬퍼
    # ------------------------------------------------------------------

    def get_node_type_display(self, node_name: str) -> str:
        """노드 타입의 사람이 읽기 쉬운 표시명을 반환한다.

        Returns:
            "" — "기본" 또는 빈값 (기본 타입)
            "Head", "Pump", ... — 고정 타입
            "Alarm Valve", "Water Spray", ... — 정규화된 카테고리명 그대로
        """
        node = self.graph.get_node(node_name)
        if node is None:
            return ""
        ntype = str(node.type or "기본").strip()
        if not ntype or ntype in ("기본", "base"):
            return ""
        lower = ntype.lower()
        if lower in _TYPE_ID_TO_DISPLAY:
            return _TYPE_ID_TO_DISPLAY[lower]
        # Head 노드 전용 폴백: node.type에 spec name(예: "80(5.6)")이 저장된 경우
        tid = str(getattr(node, "type_id", "") or "").strip().lower()
        if tid == "head":
            return _TYPE_ID_TO_DISPLAY["head"]
        return ntype

    def get_node_display_code(self, node_name: str) -> str:
        tid = str(self.get_node_type_id(node_name) or "").lower().strip()

        _map = {
            NodeTypeId.PUMP.value:    "Pu",
            NodeTypeId.HEAD.value:    "H",
            NodeTypeId.NOZZLE.value:  "Nz",
            NodeTypeId.PRV.value:     "Prv",
            NodeTypeId.ORIFICE.value: "Or",
            NodeTypeId.WT.value:          "Wt",
            NodeTypeId.RES.value:         "Res",
            NodeTypeId.HOSE_STREAM.value: "Hs",
        }
        if tid in _map:
            return _map[tid]

        node = self.graph.get_node(node_name)
        meta_type = str(getattr(node, "type", "") or "").lower().strip()

        if "alarm" in meta_type:     return "Av"
        if "dry" in meta_type:       return "Dv"
        if "check" in meta_type:     return "Ck"
        if "gate" in meta_type:      return "Gv"
        if "butterfly" in meta_type or "but" in meta_type: return "Bv"
        if "pre" in meta_type:       return "Pre"
        if "valve" in meta_type:     return "V"

        return ""

    def get_node_display_label(self, node_name: str) -> str:
        type_display = self.get_node_type_display(node_name)
        if type_display:
            return f"{node_name}({type_display})"
        return str(node_name)

    # ------------------------------------------------------------------
    # 데이터 방역 / 무결성
    # ------------------------------------------------------------------

    def _type_id_from_display(self, display: str) -> str | None:
        """
        표시명(display name)만으로 type_id를 결정합니다.
        필드값에 의존하지 않으므로 타입 변경 직후 stale 필드가 잔존해도 신뢰할 수 있는
        1차 판단입니다. 추론 불가 시 None 반환 → 호출자가 normalize_node_type_id로 폴백.

        우선순위:
          1. 주입된 콜백(JSON 라이브러리 조회 — SSOT)
          2. NodeTypeId 내부 ID 정확 매칭
          3. 특수 노드 한글 표시명 폴백 (§1.6)
        """
        if not display:
            return None

        # 1. 주입된 콜백으로 JSON 라이브러리 조회 (SSOT)
        if self._type_id_lookup:
            result = self._type_id_lookup(display)
            if result:
                return result

        s = display.lower().strip()

        # 2. 통합 룩업 (한글/영문/약어 → NodeTypeId) — SSOT: models._NODE_TYPE_LOOKUP
        looked_up = _NODE_TYPE_LOOKUP.get(s)
        if looked_up:
            return looked_up
        low = s
        if "pump" in low:
            return NodeTypeId.PUMP.value
        if "wt" in low or "tank" in low:
            return NodeTypeId.WT.value
        if "prv" in low:
            return NodeTypeId.PRV.value
        if "orif" in low:
            return NodeTypeId.ORIFICE.value

        return None  # 헤드명("80(5.6)") 등 — 호출자가 normalize_node_type_id 폴백 사용

    def sanitize_node_data(self, node_name: str) -> None:
        """
        노드 타입에 맞지 않는 Node 필드를 기본값으로 초기화한다.
        Graph가 SSOT이므로 graph.update_node()를 직접 호출한다.

        [stale type_id 감지 로직]
        UI가 node.type(표시명)을 변경한 뒤 node.type_id가 갱신되지 않은 경우에도
        _type_id_from_display()로 표시명을 직접 추론하여 불일치를 교정한다.
        """
        node = self.graph.get_node(node_name)
        if node is None:
            return

        display = str(node.type or "")
        stored_tid = str(node.type_id or "").lower().strip()

        # 1단계: 표시명(type)으로부터 type_id를 직접 추론 (필드값 무관, stale 내성)
        display_implied = self._type_id_from_display(display)

        if display_implied and display_implied != stored_tid:
            # 표시명이 명확한데 저장된 type_id와 다르면 → 표시명 우선 적용
            stored_tid = display_implied
            self.graph.update_node(node_name, type_id=stored_tid)

        # 2단계: 여전히 type_id가 없으면 필드값 기반 추론 (헤드명 "80(5.6)" 등 폴백)
        if not stored_tid:
            stored_tid = normalize_node_type_id(display, node.to_dict())
            self.graph.update_node(node_name, type_id=stored_tid)

        type_id = stored_tid

        tid = type_id.lower().strip()
        keep = _KEEP_MAP.get(tid, set())

        # 타입에 속하지 않는 필드를 기본값으로 reset
        patch: dict = {}
        for field_name, default_val in _PURGEABLE_FIELDS.items():
            if field_name not in keep:
                # list 필드는 새 객체를 생성해야 한다
                patch[field_name] = [] if default_val is None and field_name == "pump_curve_data" else default_val

        if patch:
            self.graph.update_node(node_name, **patch)

        # 밸브 타입: 알려진 category는 안전하게 교정한다. 판정 불가 category는
        # 레거시 플래그를 보존하고, fitting 우선 해석은 editor facade가 수행한다.
        if tid == "valve":
            node_after = self.graph.get_node(node_name)
            if node_after is not None:
                patch_valve: dict = {}
                if not hasattr(node_after, "is_active"):
                    patch_valve["is_active"] = True
                category = fitting_category_from_id(node_after.category_id)
                if category is not None:
                    expected_cv = is_check_valve_category(category)
                    if bool(node_after.is_check_valve) != expected_cv:
                        patch_valve["is_check_valve"] = expected_cv
                if patch_valve:
                    self.graph.update_node(node_name, **patch_valve)

    def reset_node_to_base(self, node_name: str, display_name: str = "기본") -> None:
        if not self.graph.has_node(node_name):
            return

        # 타입·상태·모든 타입별 필드를 한 번에 graph에 기록
        self.graph.update_node(
            node_name,
            type=str(display_name),
            type_id="base",
            category_id="",
            is_active=True,
            is_closed=False,
            is_check_valve=False,
            k_factor_si=None,
            head_spec_name=None,
            nozzle_name=None,
            required_pressure_bar=0.0,
            pressure_setting_bar=None,
            fitting_id=None,
            hole_diameter=0.0,
            orifice_discharge_coeff=0.0,
            pipe_dn=0.0,
            loss_coefficient=0.0,
            check_valve_direction="",
            rated_q=0.0,
            rated_p=0.0,
            shutoff_p=0.0,
            peak_q=0.0,
            peak_p=0.0,
            pump_library_id=None,
            pump_curve_data=[],
            water_level=0.0,
            has_check_valve=True,
            hose_stream_flow_lps=400.0 / 60.0,
        )
        self._refresh_node_views_from_graph(node_name)

    def validate_and_fix_integrity(
        self, pipe_registry: dict, *, strict: bool = False
    ) -> dict:
        """
        무결성 검사 및 자동 보정.
        pipe_registry는 PipeManager 소유이지만 검사에 필요하므로 인자로 받는다.
        """
        import re

        report: dict[str, list] = {"fixes": [], "warnings": [], "errors": []}

        if not isinstance(self.node_counter, dict):
            self.node_counter.clear()
            self.node_counter["N"] = 0
            report["fixes"].append("node_counter 초기화")

        self.sync_type_ids(verbose=False)

        name_pat = re.compile(r"^N\d+$")
        node_names = list(self.graph.nodes.keys())

        for n in node_names:
            if not name_pat.match(str(n)):
                msg = f"노드 내부키 규칙 위반: '{n}' (내부키는 N<number>만 허용)"
                if strict:
                    report["errors"].append(msg)
                else:
                    report["warnings"].append(msg)

        for n in node_names:
            node = self.graph.get_node(n)
            if node is None:
                continue
            display = str(node.type or "")

            if not node.type_id:
                tid = normalize_node_type_id(display, node.to_dict())
                self.graph.update_node(n, type_id=tid)
                report["fixes"].append(f"{n}: type_id 보정 → {tid}")

        # 고아 배관 검사
        bad_pids = [
            pid for pid, p in list(pipe_registry.items())
            if not self.graph.has_node(p.start) or not self.graph.has_node(p.end)
        ]
        if bad_pids:
            if strict:
                report["errors"].append(
                    f"끊긴 배관 존재: {len(bad_pids)}개 (파이프 끝점 노드 누락)"
                )
            else:
                for pid in bad_pids:
                    self.graph.remove_pipe(pid)
                report["fixes"].append(f"끊긴 배관 자동 삭제: {len(bad_pids)}개")

        # 배관 구경 0 검사
        zero_diam_pipes = [
            pid for pid, p in pipe_registry.items() if p.diameter <= 0.1
        ]
        if zero_diam_pipes:
            msg = (
                f"구경(Diameter)이 0인 배관 발견: {len(zero_diam_pipes)}개 "
                f"({', '.join(zero_diam_pipes[:3])}...)"
            )
            if strict:
                report["errors"].append(msg + " -> 속성을 수정하세요.")
            else:
                report["warnings"].append(msg)

        # node_counter 보정
        try:
            max_n = 0
            for n in node_names:
                if name_pat.match(str(n)):
                    max_n = max(max_n, int(str(n)[1:]))
            if self.node_counter.get("N", 0) < max_n:
                self.node_counter["N"] = max_n
                report["fixes"].append(f"node_counter['N'] 보정 → {max_n}")
        except Exception:
            pass

        return report

    # ------------------------------------------------------------------
    # 헤드 전용 헬퍼
    # ------------------------------------------------------------------

    def create_head_node(
        self,
        name: str | None = None,
        x: float = 0.0,
        y: float = 0.0,
        z: float = 0.0,
        k_factor_si: float = 80.0,
        active: bool = True,
    ) -> Node:
        if name is None:
            name = f"H{len(self.graph.nodes) + 1}"

        node = Node.create(name, x, y, z, "head")
        node.k_factor_si = float(k_factor_si)
        node.is_active = bool(active)

        self.graph.add_node(node)
        self._refresh_node_views_from_graph(name)

        logger.debug("[create_head_node] Added head %s (K=%s, Active=%s)", name, k_factor_si, active)
        return node

    # ------------------------------------------------------------------
    # 서브트리 / 이웃 탐색
    # ------------------------------------------------------------------

    def get_neighbors(self, name: str, pipe_registry: dict) -> list[str]:
        if not self.graph.has_node(name):
            return []
        neighbors: set[str] = set()
        for p in pipe_registry.values():
            if p.start == name:
                neighbors.add(p.end)
            elif p.end == name:
                neighbors.add(p.start)
        return sorted(neighbors)

    def get_subtree_nodes(
        self, root: str, pipe_registry: dict, blocked_neighbor: str | None = None
    ) -> set[str]:
        if not self.graph.has_node(root):
            return set()
        adj: dict[str, set[str]] = {}
        for p in pipe_registry.values():
            s, e = p.start, p.end
            adj.setdefault(s, set()).add(e)
            adj.setdefault(e, set()).add(s)

        visited: set[str] = {root}
        queue = [root]
        while queue:
            n = queue.pop(0)
            for nb in adj.get(n, []):
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
        self,
        root: str,
        toward_neighbor: str | None,
        pipe_registry: dict,
    ) -> set[str]:
        if not self.graph.has_node(root):
            return set()
        if toward_neighbor is None:
            return self.get_subtree_nodes(root, pipe_registry)

        adj: dict[str, set[str]] = {}
        for p in pipe_registry.values():
            s, e = p.start, p.end
            adj.setdefault(s, set()).add(e)
            adj.setdefault(e, set()).add(s)

        if toward_neighbor not in adj.get(root, set()):
            return self.get_subtree_nodes(root, pipe_registry)

        visited: set[str] = {root, toward_neighbor}
        queue = [toward_neighbor]
        while queue:
            n = queue.pop(0)
            for nb in adj.get(n, []):
                if nb == root:
                    continue
                if nb not in visited:
                    visited.add(nb)
                    queue.append(nb)
        return visited

    # ------------------------------------------------------------------
    # 내부 헬퍼
    # ------------------------------------------------------------------

    def _build_domain_node(self, name: str) -> Node | None:
        """
        Graph가 SSOT이므로 deepcopy만 반환한다.
        과거에는 외부 메타 overlay가 필요했으나 DR-1 이후 불필요.
        """
        graph_node = self.graph.get_node(name)
        if graph_node is None:
            return None
        return copy.deepcopy(graph_node)

    def _sync_graph_node_from_legacy(self, node_name: str) -> None:
        """
        [전환기 아티팩트 — DR-5에서 제거 예정]
        과거에는 외부 메타 → graph 방향 동기화가 있었으나,
        DR-2 전환 후 graph가 항상 최신 상태이므로 아무 작업도 하지 않는다.
        호출부(sanitize_node_data, reset_node_to_base)가 이미 graph를 직접 갱신한다.
        """

    def _refresh_node_views_from_graph(self, node_name: str) -> None:
        """(레거시) Graph 기준으로 node view/cache를 갱신하던 메서드. 이제 아무 일도 하지 않는다."""
        pass

    def _make_node_object(self, name: str) -> Node | None:
        """_build_domain_node의 별칭 (과도기 호환)"""
        return self._build_domain_node(name)
