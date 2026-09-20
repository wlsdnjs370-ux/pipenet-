# 파일명 editor_core.py
"""
PipeEditor — 얇은 파사드(Facade) 클래스

Phase 2-1 리팩토링:
  기존 PipeEditor의 노드/배관 로직을 domain/node_manager.py, domain/pipe_manager.py로 추출.
  이 클래스는 기존 메서드 시그니처를 100% 유지하면서 내부적으로 매니저에 위임한다.
  외부(main.py, view_graphics.py, hydraulic_engine.py 등)에서는 변경 없이 동작.
"""
import copy
import logging
from domain.geometry import snap_point, snap_value
from domain.models import Node
import domain.models as models
from domain.network_graph import NetworkGraph
from domain.node_manager import NodeManager
from domain.pipe_manager import PipeManager
from infra.project_serializer import ProjectSerializer
from protocols import ProjectSerializerProtocol
from services.library_service import LibraryService
from domain.fitting_models import FittingDefinition

logger = logging.getLogger(__name__)

# design_settings 기본값 SSOT — roughness_mm 등 기본값을 한 곳에서 관리
# roughness_mm=0.100: NFPA 습식강관 절대조도(=대치 C120). 신규 프로젝트에만 적용되며,
# 기존 .kfp에 저장된 값은 로드 시 그대로 유지된다(project_serializer가 저장값 우선).
DEFAULT_DESIGN_SETTINGS: dict = {
    "min_required_pressure_bar": 1.0,
    "calculation_method": "H-W",
    "roughness_mm": 0.100,
}


class PipeEditor:
    def __init__(self):
        self.pipe_id_counter = 0

        # 라이브러리 서비스 ([2-2] LibraryService에 위임)
        self._lib = LibraryService()

        self.current_standard = None
        self.current_dn = None
        self.current_C = None
        self.default_dn = "DN25"

        self.subtree_clipboard = None
        self.subtree_paste_pending = False
        self._last_paste_error: str = ""

        self.grid_step = 0.01
        self.snap_tol = self.grid_step * 0.45
        self.intersection_tol = 1e-6

        # 노드 저장 (과도기 구조)
        self.node_counter = {"N": 0}

        # 전역 설계 설정
        self.design_settings = dict(DEFAULT_DESIGN_SETTINGS)

        # Antifreeze analysis exists only as an explicitly created case.
        # None is the sole default and is never inferred from formula selection.
        self.antifreeze_analysis_case = None
        self.antifreeze_report_pdf_path = None

        # 프로젝트 개요 (SSOT: .kfp에 함께 저장)
        self.report_common: dict = {
            "design_area": "",
            "company": "",
            "project": "",
            "date": "",
            "notes": "",
        }

        # [1-3 브릿지] NetworkGraph
        self.graph = NetworkGraph()

        # [2-1] 매니저 초기화 — 공유 상태 참조를 주입
        self._node_mgr = NodeManager(
            self.node_counter,
            self.graph,
            self.grid_step,
            type_id_lookup=lambda name: self._lib.find_category_type_id(name),
        )
        self._pipe_mgr = PipeManager(
            pipe_registry=self.graph.pipes,
            graph=self.graph,
            node_manager=self._node_mgr,
            grid_step=self.grid_step,
            intersection_tol=self.intersection_tol,
        )
        self._pipe_mgr.pipe_id_counter = self.pipe_id_counter
        self._pipe_mgr._design_settings = self.design_settings
        self._pipe_mgr._current_dn = getattr(self, "current_dn", "DN25")
        self._pipe_mgr._pipe_spec_getter = getattr(self, "get_pipe_spec", None)

        # [2-2] 직렬화기
        self._serializer: ProjectSerializerProtocol = ProjectSerializer(self)

    # ==================================================================
    # [2-2] 라이브러리 프로퍼티 (외부에서 editor.pipe_library = ... 호환 유지)
    # ==================================================================

    @property
    def pipe_library(self):
        return self._lib.pipe_library

    @pipe_library.setter
    def pipe_library(self, value):
        self._lib.pipe_library = value

    @property
    def nozzle_library(self):
        return self._lib.nozzle_library

    @nozzle_library.setter
    def nozzle_library(self, value):
        self._lib.nozzle_library = value

    @property
    def pump_library(self):
        return self._lib.pump_library

    @pump_library.setter
    def pump_library(self, value):
        self._lib.pump_library = value

    @property
    def pipe_sizing_library(self):
        return self._lib.pipe_sizing_library

    @pipe_sizing_library.setter
    def pipe_sizing_library(self, value):
        self._lib.pipe_sizing_library = value

    # ==================================================================
    # [2-1] 매니저 동기화 헬퍼
    # ==================================================================

    def _sync_pipe_id_counter(self):
        """pipe_id_counter를 PipeManager와 동기화"""
        self.pipe_id_counter = self._pipe_mgr.pipe_id_counter

    # ==================================================================
    # [1-3 브릿지] NetworkGraph 동기화 헬퍼 (NodeManager에 위임)
    # ==================================================================

    def _build_domain_node(self, name: str) -> Node | None:
        return self._node_mgr._build_domain_node(name)

    def _sync_graph_from_existing(self) -> None:
        # D2-a: pipe_registry는 graph.pipes property이므로, graph 교체 전에 기존 배관을 보존한다.
        existing_nodes = dict(self.graph.nodes)
        existing_pipes = dict(self.graph.pipes)
        self.graph = NetworkGraph()
        # 매니저의 graph 참조도 갱신
        self._node_mgr.graph = self.graph
        self._pipe_mgr.graph = self.graph

        for name, node in existing_nodes.items():
            domain_node = self._node_mgr._build_domain_node(name)
            if domain_node is not None:
                self.graph.add_node(domain_node)

        for pid, pipe in existing_pipes.items():
            self.graph.add_pipe(pipe)
        self._refresh_all_views()

    def _refresh_node_views(self, node_name: str) -> None:
        node = self.graph.get_node(node_name)
        if node is None:
            return

        if node_name.startswith("N") and node_name[1:].isdigit():
            node_no = int(node_name[1:])
            if self.node_counter.get("N", 0) < node_no:
                self.node_counter["N"] = node_no

    def _refresh_node_type_id_cache(self) -> None:
        pass

    def _refresh_all_views(self) -> None:
        self._pipe_mgr._mark_pipe_key_index_dirty()

    # ==================================================================
    # 배관 인덱스 (PipeManager에 위임)
    # ==================================================================

    def _rebuild_pipe_key_index(self):
        self._pipe_mgr._rebuild_pipe_key_index()

    def _mark_pipe_key_index_dirty(self):
        self._pipe_mgr._mark_pipe_key_index_dirty()

    # ==================================================================
    # 노드 관련 (NodeManager에 위임)
    # ==================================================================

    def reset_node_to_base(self, node_name: str, display_name: str = "기본"):
        self.update_node_via_graph(
            node_name,
            node_type=str(display_name),
            type_id="base",
            category_id="",
            is_active=True,
            is_closed=False,
            k_factor_si=None,
            required_pressure_bar=0.0,
            fitting_id=None,
            head_spec_name=None,
            nozzle_name=None,
            rated_q=0.0,
            rated_p=0.0,
            shutoff_p=0.0,
            peak_q=0.0,
            peak_p=0.0,
            pump_library_id=None,
            pump_curve_data=[],
            water_level=0.0,
            pressure_setting_bar=None,
            loss_coefficient=0.0,
            hole_diameter=0.0,
            orifice_discharge_coeff=0.0,
            is_check_valve=False,
        )

    def validate_and_fix_integrity(self, strict: bool = False):
        report = self._node_mgr.validate_and_fix_integrity(
            self.graph.pipes, strict=strict
        )
        # 길이 업데이트
        try:
            self.update_all_pipe_lengths_in_pipe_data()
        except Exception as e:
            from services.i18n_service import _t
            report.setdefault("warnings", []).append(_t("배관 길이 업데이트 실패: {e}").format(e=e))
        self._refresh_all_views()
        return report

    def sanitize_node_data(self, node_name):
        self._node_mgr.sanitize_node_data(node_name)
        node = self.graph.get_node(node_name)
        if node is not None and str(node.type_id or "").strip().lower() == "valve":
            from services.check_valve_category_resolver import (
                resolve_node_check_valve_identity,
            )

            identity = resolve_node_check_valve_identity(
                node,
                self.get_fitting_data_v3,
                warn_unresolved=True,
            )
            if identity.resolved:
                self.graph.update_node(
                    node_name,
                    category_id=identity.category_id,
                    is_check_valve=identity.is_check_valve,
                )
        self._refresh_node_views(node_name)

    def generate_node_name(self, kind="N"):
        return self._node_mgr.generate_node_name(kind)

    def rename_node(self, old_name, new_name):
        return self.rename_node_via_graph(str(old_name), str(new_name))

    def add_node(self, name, x, y, z, node_type="기본"):
        self.add_node_via_graph(
            str(name),
            float(x),
            float(y),
            float(z),
            node_type=str(node_type),
            split_existing_pipes=True,
        )

    def add_node_via_graph(
        self,
        name: str,
        x: float,
        y: float,
        z: float,
        node_type: str = "기본",
        *,
        meta: dict | None = None,
        split_existing_pipes: bool = True,
    ) -> Node:
        if self.graph.has_node(name):
            from services.i18n_service import _t
            raise ValueError(_t("노드 이름 {name} 이(가) 이미 존재합니다.").format(name=name))

        sx, sy, sz = snap_point((float(x), float(y), float(z)), self.grid_step)
        node = Node.create(name, sx, sy, sz, node_type=node_type)
        if isinstance(meta, dict):
            node.update_from_meta(copy.deepcopy(meta))

        self.graph.add_node(node)
        self._refresh_node_views(name)

        if split_existing_pipes:
            self._pipe_mgr._split_existing_pipes_at_node(name)
            self._sync_pipe_id_counter()
        return node

    def update_node_via_graph(
        self,
        name: str,
        *,
        x: float | None = None,
        y: float | None = None,
        z: float | None = None,
        node_type: str | None = None,
        meta: dict | None = None,
        **updates,
    ) -> bool:
        node = self.graph.get_node(name)
        if node is None:
            return False

        patch: dict = {}
        cur_x, cur_y, cur_z = node.coords
        nx = cur_x if x is None else float(x)
        ny = cur_y if y is None else float(y)
        nz = cur_z if z is None else float(z)
        if (nx, ny, nz) != (cur_x, cur_y, cur_z):
            sx, sy, sz = snap_point((nx, ny, nz), self.grid_step)
            patch["coords"] = (sx, sy, sz)
            patch["elevation_m"] = sz

        if node_type is not None:
            patch["type"] = str(node_type)
        if "k_factor" in updates and "k_factor_si" not in updates:
            updates["k_factor_si"] = updates.pop("k_factor")

        for key, value in updates.items():
            if hasattr(node, key):
                patch[key] = value

        if isinstance(meta, dict):
            meta_update = dict(meta)
            if "k_factor" in meta_update and "k_factor_si" not in meta_update:
                meta_update["k_factor_si"] = meta_update["k_factor"]
            for key, value in meta_update.items():
                if hasattr(node, key):
                    patch[key] = value

        if patch:
            self.graph.update_node(name, **patch)
        # 준비작동식 FDT 기본값/레거시 보정은 도메인 SSOT 한 곳에서 — 파일 로드(update_from_meta)와
        # 동일 규칙을 앱 내 편집 경로에도 적용해 화면값·저장값·계산값을 일치시킨다.
        if str(getattr(node, "category_id", "") or "") == "preaction_valve":
            node.apply_preaction_fdt_defaults(meta if isinstance(meta, dict) else {})
        self._refresh_node_views(name)
        return True

    def delete_node_via_graph(self, name: str) -> bool:
        if not self.graph.has_node(name):
            return False

        connected_pids = [
            pid for pid, p in self.graph.pipes.items()
            if p.start == name or p.end == name
        ]
        if len(connected_pids) >= 2:
            return False

        for pid in connected_pids:
            p = self.graph.pipes[pid]
            self.delete_pipe(p.start, p.end)

        self.graph.remove_node(name)
        self._refresh_node_views(name)
        return True

    def rename_node_via_graph(self, old_name: str, new_name: str) -> bool:
        if old_name == new_name:
            return True
        if not self.graph.has_node(old_name):
            return False
        if self.graph.has_node(new_name):
            from services.i18n_service import _t
            raise ValueError(_t("노드 이름 {name} 이(가) 이미 존재합니다.").format(name=new_name))

        src = self.graph.get_node(old_name)
        if src is None:
            return False

        renamed = copy.deepcopy(src)
        renamed.id = new_name
        self.graph.add_node(renamed)

        for pipe in self.graph.pipes.values():
            patch: dict = {}
            if pipe.start == old_name:
                patch["start"] = new_name
            if pipe.end == old_name:
                patch["end"] = new_name
            if patch:
                self.graph.update_pipe(pipe.id, **patch)

        self.graph.remove_node(old_name)

        for node_name in list(self.graph.nodes.keys()):
            candidate = self.graph.get_node(node_name)
            if candidate is None:
                continue
            if str(getattr(candidate, "check_valve_direction", "") or "") == old_name:
                self.graph.update_node(node_name, check_valve_direction=new_name)

        self._refresh_all_views()
        return True

    def delete_node(self, name):
        return self.delete_node_via_graph(str(name))

    def find_node_at_position(self, x, y, z, tol=None):
        return self._node_mgr.find_node_at_position(x, y, z, tol)

    def get_node_type_id(self, node_name: str) -> str:
        node = self.graph.get_node(node_name)
        if node is None:
            return "base"
        if node.type_id:
            return str(node.type_id)

        inferred = models.normalize_node_type_id(node.type, node.to_dict())
        self.graph.update_node(node_name, type_id=inferred)
        self._refresh_node_views(node_name)
        return inferred

    def set_node_type_id(self, node_name: str, type_id: str, display_name: str | None = None):
        node = self.graph.get_node(node_name)
        if node is None:
            return

        patch = {"type_id": str(type_id)}
        if display_name is not None:
            patch["type"] = str(display_name)
        self.graph.update_node(node_name, **patch)
        self._refresh_node_views(node_name)

    def get_nodes_by_type_id(self, type_id: str) -> list:
        target = str(type_id)
        return sorted(
            node.id for node in self.graph.get_all_nodes()
            if str(node.type_id) == target
        )

    def sync_type_ids(self, *, verbose=False):
        changed = 0
        for node_name in list(self.graph.nodes.keys()):
            node = self.graph.get_node(node_name)
            if node is None:
                continue

            inferred = models.normalize_node_type_id(node.type, node.to_dict())
            if node.type_id != inferred:
                self.graph.update_node(node_name, type_id=inferred)
                changed += 1
                if verbose:
                    logger.debug("[SYNC] %s: type_id -> %s (display=%s)", node_name, inferred, node.type)

        if changed:
            self._refresh_all_views()
        return changed

    def _infer_and_set_type_id(self, node_name: str):
        node = self.graph.get_node(node_name)
        if node is None:
            return

        inferred = models.normalize_node_type_id(node.type, node.to_dict())
        self.graph.update_node(node_name, type_id=inferred)
        self._refresh_node_views(node_name)

    def get_node_display_code(self, node_name: str) -> str:
        return self._node_mgr.get_node_display_code(node_name)

    def get_node_type_display(self, node_name: str) -> str:
        return self._node_mgr.get_node_type_display(node_name)

    def get_node_display_label(self, node_name: str) -> str:
        return self._node_mgr.get_node_display_label(node_name)

    def create_head_node(self, name=None, x=0.0, y=0.0, z=0.0, k_factor_si=80.0, active=True):
        node_name = str(name) if name is not None else f"H{len(self.graph.nodes) + 1}"
        self.add_node_via_graph(
            node_name,
            float(x),
            float(y),
            float(z),
            node_type="head",
            meta={
                "type_id": "head",
                "k_factor_si": float(k_factor_si),
                "is_active": bool(active),
            },
            split_existing_pipes=False,
        )
        node = self.graph.get_node(node_name)
        if node is None:
            from services.i18n_service import _t
            raise ValueError(_t("헤드 노드 생성 실패: {node_name}").format(node_name=node_name))
        return node

    def get_neighbors(self, name):
        return self._node_mgr.get_neighbors(name, self.graph.pipes)

    def get_subtree_nodes(self, root, blocked_neighbor=None):
        return self._node_mgr.get_subtree_nodes(root, self.graph.pipes, blocked_neighbor)

    def get_subtree_one_direction(self, root, toward_neighbor=None):
        return self._node_mgr.get_subtree_one_direction(root, toward_neighbor, self.graph.pipes)

    def delete_subtree_one_direction(self, root_name, toward_neighbor):
        """root에서 toward_neighbor 방향으로 뻗은 서브트리(배관+노드)를 삭제.

        Root 노드는 보존한다. 고립되더라도 삭제하지 않아 cut 후 붙여넣기
        앵커와 선택 상태를 유지한다. 단, 그래프가 완전히 비는 극단적
        엣지케이스에서는 이후 편집이 불가능해지므로 DeletePipeCommand와
        동일하게 N1을 부활시킨다.
        """
        subtree_nodes = self.get_subtree_one_direction(root_name, toward_neighbor)

        pipes_to_delete = []
        for p in self.graph.pipes.values():
            s, e = p.start, p.end
            is_internal = (s in subtree_nodes and e in subtree_nodes)
            is_link = ((s == root_name and e in subtree_nodes) or
                       (e == root_name and s in subtree_nodes))
            if is_internal or is_link:
                pipes_to_delete.append((s, e))

        for s, e in pipes_to_delete:
            self.delete_pipe(s, e)

        for n in subtree_nodes:
            if n == root_name:
                continue
            if self.graph.has_node(n):
                self.delete_node_via_graph(n)

        if not self.graph.nodes:
            self.add_node("N1", 0, 0, 0)
            self.node_counter["N"] = 1

    def _make_node_object(self, name):
        return self._node_mgr._make_node_object(name)

    # ==================================================================
    # 배관 관련 (PipeManager에 위임)
    # ==================================================================

    def _pipe_key(self, start, end):
        return PipeManager._pipe_key(start, end)

    def _snap_point(self, x, y, z):
        return snap_point((x, y, z), self.grid_step)

    def _length_to_snapped(self, length):
        return snap_value(float(length), self.grid_step)

    def get_pipe_id_by_nodes(self, n1, n2):
        return self._pipe_mgr.get_pipe_id_by_nodes(n1, n2)

    def update_pipe_data(self, pid, **kwargs):
        self._pipe_mgr.update_pipe_data(pid, **kwargs)

    def _resolve_pipe_attributes(self, override_attrs: dict = None):
        return self._pipe_mgr._resolve_pipe_attributes(
            override_attrs=override_attrs,
            design_settings=getattr(self, "design_settings", {}),
            current_dn=getattr(self, "current_dn", "DN25"),
            pipe_spec_getter=self.get_pipe_spec,
        )

    def create_pipe(self, start, end, attributes=None, silent=False) -> tuple:
        result = self._pipe_mgr.create_pipe(
            start, end, attributes=attributes, silent=silent,
            design_settings=getattr(self, "design_settings", {}),
            current_dn=getattr(self, "current_dn", "DN25"),
            pipe_spec_getter=self.get_pipe_spec,
        )
        self._sync_pipe_id_counter()
        return result

    def add_pipe(self, start, end, silent=False):
        success, msg = self.create_pipe(start, end, attributes=None, silent=silent)
        return success

    def delete_pipe(self, start, end):
        self._pipe_mgr.delete_pipe(start, end)

    def delete_pipe_data(self, pid):
        pass

    def split_pipe(self, start, end, dist_from_start, new_name):
        self._pipe_mgr.split_pipe(start, end, dist_from_start, new_name)
        self._sync_pipe_id_counter()

    def validate_new_segment(self, start, end, ignore_pipes=None):
        return self._pipe_mgr.validate_new_segment(start, end, ignore_pipes)

    def _check_collision(self, start, end, axis, ignore_pipes=None):
        return self._pipe_mgr._check_collision(start, end, axis, ignore_pipes)

    def _has_axis_overlap(self, start, end, ignore_pipe=None):
        return self._pipe_mgr._has_axis_overlap(start, end, ignore_pipe)

    def _infer_pipe_axis(self, start, end):
        return self._pipe_mgr._infer_pipe_axis(start, end)

    def _find_segment_intersections(self, p1, p2, axis_new):
        return self._pipe_mgr._find_segment_intersections(p1, p2, axis_new)

    def _split_new_pipe_on_existing_nodes(self, start, end):
        self._pipe_mgr._split_new_pipe_on_existing_nodes(start, end)
        self._sync_pipe_id_counter()

    def _split_existing_pipes_at_node(self, node_name):
        self._pipe_mgr._split_existing_pipes_at_node(node_name)
        self._sync_pipe_id_counter()

    def get_pipe_props(self, start, end):
        return self._pipe_mgr.get_pipe_props(start, end)

    def _find_pipe_index(self, start, end):
        return self._pipe_mgr._find_pipe_index(start, end)

    def get_pipe_length(self, start, end):
        return self._pipe_mgr.get_pipe_length(start, end)

    def compute_length(self, n1, n2):
        return self._pipe_mgr.compute_length(n1, n2)

    def update_all_pipe_lengths_in_pipe_data(self):
        self._pipe_mgr.update_all_pipe_lengths()

    def is_pipe_in_loop(self, start_node, end_node):
        return self._pipe_mgr.is_pipe_in_loop(start_node, end_node)

    def _detect_fitting_at_node(self, node_name):
        return self._pipe_mgr._detect_fitting_at_node(node_name)

    def apply_pipe_attributes(self, pipe_id, attrs):
        self._pipe_mgr.apply_pipe_attributes(pipe_id, attrs)

    def rebuild_pipe_data_from_pipes(self):
        pass

    def add_branch_from_node_axis(self, node, axis, length, name, node_type="기본", pipe_attrs=None) -> tuple:
        result = self._pipe_mgr.add_branch_from_node_axis(
            node, axis, length, name, node_type, pipe_attrs
        )
        self._sync_pipe_id_counter()
        return result

    def add_branch_from_node_diagonal(self, node, direction, run, name, node_type="기본", pipe_attrs=None) -> tuple:
        result = self._pipe_mgr.add_branch_from_node_diagonal(
            node, direction, run, name, node_type, pipe_attrs
        )
        self._sync_pipe_id_counter()
        return result

    def auto_route_v1(self, start, end):
        result = self._pipe_mgr.auto_route_v1(start, end)
        self._sync_pipe_id_counter()
        return result

    # ==================================================================
    # 중간 노드 병합 삭제
    # ==================================================================

    # 병합 삭제가 불가능한 노드 타입 (수원·펌프만 차단)
    _MERGE_DELETE_BLOCKED_TYPES: set[str] = {
        "pump", "wt", "reservoir",
    }

    def preview_merge_delete_node(self, node_name: str) -> tuple[bool, str]:
        """중간 노드 병합 삭제 가능 여부를 판정한다.

        Returns:
            (가능 여부, 불가 시 사유 메시지)
        """
        from services.i18n_service import _t

        node = self.graph.get_node(node_name)
        if node is None:
            return False, _t("노드를 찾을 수 없습니다")

        # 1. 노드 타입 검사 (get_node_type_id 게이트웨이 — 구버전 데이터 추론 포함)
        resolved_type_id = self.get_node_type_id(node_name)
        if resolved_type_id in self._MERGE_DELETE_BLOCKED_TYPES:
            return False, _t("이 노드는 삭제할 수 없는 타입({node_type})입니다").format(node_type=node.type)

        # 2. 연결 배관 수 검사
        connected_pids = self.graph.get_connected_pipes(node_name)
        if len(connected_pids) != 2:
            return False, _t("연결 배관이 {count}개여서 삭제할 수 없습니다").format(count=len(connected_pids))

        pipe_a = self.graph.get_pipe(connected_pids[0])
        pipe_b = self.graph.get_pipe(connected_pids[1])
        if pipe_a is None or pipe_b is None:
            return False, _t("배관 정보를 찾을 수 없습니다")

        # 3. 설계 속성 동일 여부
        for attr in ("type", "diameter", "nominal_mm", "C", "roughness_mm"):
            if getattr(pipe_a, attr) != getattr(pipe_b, attr):
                return False, _t("양쪽 배관의 규격이 달라 합칠 수 없습니다")

        # 4. 동축(직선) 검사
        node_a = pipe_a.end if pipe_a.start == node_name else pipe_a.start
        node_c = pipe_b.end if pipe_b.start == node_name else pipe_b.start
        axis_ab = self._infer_pipe_axis(node_a, node_name)
        axis_bc = self._infer_pipe_axis(node_name, node_c)
        if not axis_ab or not axis_bc or axis_ab != axis_bc:
            return False, _t("배관이 일직선이 아니어서 합칠 수 없습니다")

        # 5. 병합 배관 A→C 유효성 (preflight)
        valid, msg = self.validate_new_segment(
            node_a, node_c,
            ignore_pipes=[(pipe_a.start, pipe_a.end), (pipe_b.start, pipe_b.end)],
        )
        if not valid:
            return False, _t("병합 배관을 생성할 수 없습니다: {msg}").format(msg=msg)

        return True, ""

    def merge_delete_node(self, node_name: str) -> tuple[bool, str]:
        """중간 노드를 삭제하고 양쪽 배관을 하나로 합친다.

        preview_merge_delete_node()로 사전 검증한 뒤 호출해야 한다.

        Returns:
            (성공 여부, 실패 시 메시지)
        """
        # 재검증 (호출 시점과 preview 사이에 상태가 바뀔 수 있으므로)
        ok, reason = self.preview_merge_delete_node(node_name)
        if not ok:
            return False, reason

        connected_pids = self.graph.get_connected_pipes(node_name)
        pipe_a = self.graph.get_pipe(connected_pids[0])
        pipe_b = self.graph.get_pipe(connected_pids[1])

        # 양쪽 끝 노드 결정
        node_a = pipe_a.end if pipe_a.start == node_name else pipe_a.start
        node_c = pipe_b.end if pipe_b.start == node_name else pipe_b.start

        # 합산 속성 계산
        merged_length = pipe_a.length_m + pipe_b.length_m
        merged_eq_length = pipe_a.equivalent_length + pipe_b.equivalent_length
        merged_fittings = list(pipe_a.fittings) + list(pipe_b.fittings)

        # 실패 시 지운 노드·배관만 되돌린다. 그래프 전체 deepcopy 는
        # CAD 변환 일괄 정리에서 노드마다 수 초씩 붙는다.
        node_obj = self.graph.get_node(node_name)

        self.delete_pipe(pipe_a.start, pipe_a.end)
        self.delete_pipe(pipe_b.start, pipe_b.end)
        self.delete_node_via_graph(node_name)

        success, msg = self.create_pipe(node_a, node_c, attributes={
            "type": pipe_a.type,
            "diameter": pipe_a.diameter,
            "nominal_mm": pipe_a.nominal_mm,
            "C": pipe_a.C,
            "roughness_mm": pipe_a.roughness_mm,
        })
        if not success:
            if node_obj is not None and not self.graph.has_node(node_name):
                self.graph.add_node(node_obj)
            if not self.graph.has_pipe(pipe_a.id):
                self.graph.add_pipe(pipe_a)
            if not self.graph.has_pipe(pipe_b.id):
                self.graph.add_pipe(pipe_b)
            self._mark_pipe_key_index_dirty()
            from services.i18n_service import _t
            return False, _t("병합 배관 생성 실패: {msg}").format(msg=msg)

        # create_pipe가 length를 좌표 기반으로 계산하므로, 합산값으로 보정
        new_pid = self.get_pipe_id_by_nodes(node_a, node_c)
        if new_pid:
            self._pipe_mgr.apply_pipe_attributes(new_pid, {
                "length_m": merged_length,
                "equivalent_length": merged_eq_length,
                "fittings": merged_fittings,
            })

        self._sync_pipe_id_counter()
        return True, ""

    def collect_collinear_merge_candidates(self) -> list[str]:
        """일직선 중간 노드(병합 삭제 가능) 후보를 모은다.

        기존 preview_merge_delete_node 판정을 재사용한다.
        헤드·노즐 등 비기본(type_id ≠ base) 노드는 일괄 정리에서 제외한다.
        """
        from domain.node_predicates import is_base_type_id

        out: list[str] = []
        for node_name in sorted(self.graph.nodes.keys(), key=str):
            if not is_base_type_id(self.get_node_type_id(node_name)):
                continue
            ok, _reason = self.preview_merge_delete_node(node_name)
            if ok:
                out.append(node_name)
        return out

    def cleanup_collinear_intermediate_nodes(self) -> dict:
        """일직선 중간 노드를 없을 때까지 병합 삭제한다.

        후보 수집(전 노드 preview)이 가장 비싸므로 라운드당 한 번만 하고,
        그 라운드 안에서 후보를 순서대로 전부 병합 시도한다. 병합 자체는
        단건 merge_delete_node(SSOT)를 그대로 쓴다. 같은 라운드의 앞선
        병합으로 조건이 바뀐 후보는 preview 재검증에서 걸러 다음 라운드
        재수집에 맡긴다 (단건 반복과 같은 결과).

        Returns:
            {"removed": [삭제된 노드 id…], "failed": [{"id", "reason"}, …]}
        """
        removed: list[str] = []
        failed: list[dict] = []
        skip: set[str] = set()
        max_rounds = max(len(self.graph.nodes), 1) * 2
        for _ in range(max_rounds):
            candidates = [
                n for n in self.collect_collinear_merge_candidates()
                if n not in skip
            ]
            if not candidates:
                break
            for name in candidates:
                ok, _reason = self.preview_merge_delete_node(name)
                if not ok:
                    continue
                ok, reason = self.merge_delete_node(name)
                if ok:
                    removed.append(name)
                else:
                    failed.append({"id": name, "reason": reason})
                    skip.add(name)
        return {"removed": removed, "failed": failed}

    # ==================================================================
    # 라이브러리 조회 ([2-2] LibraryService에 위임)
    # ==================================================================

    def _robust_get_library_item(self, standard, dn_mm):
        return self._lib._robust_get_library_item(standard, dn_mm)

    def get_pipe_spec(self, standard, dn_mm):
        return self._lib.get_pipe_spec(standard, dn_mm)

    def get_nozzle_spec(self, nozzle_name):
        return self._lib.get_nozzle_spec(nozzle_name)

    def get_head_spec(self, head_name):
        return self._lib.get_head_spec(head_name)

    def get_fitting_data_v3(self, fitting_id: str) -> FittingDefinition | None:
        """v3 라이브러리에서 FittingDefinition 조회. 없으면 None 반환."""
        result = self._lib.get_fitting_data_v3(fitting_id)
        if result is None:
            logger.warning("[LIB] v3 fitting miss: id=%s", fitting_id)
        return result

    def get_fittings_v3_flat(self) -> list[FittingDefinition]:
        """v3 라이브러리 전체 항목 반환."""
        return self._lib.get_fittings_v3_flat()

    def get_pipe_inner_diameter_mm(self, dn_mm):
        return self._lib.get_pipe_inner_diameter_mm(dn_mm, self.current_standard)

    def find_nozzle_category_by_display(self, display_name: str):
        return self._lib.find_nozzle_category_by_display(display_name)

    def find_category_type_id(self, display_name: str):
        return self._lib.find_category_type_id(display_name)

    def get_nozzle_items_by_category(self, category_id: str):
        return self._lib.get_nozzle_items_by_category(category_id)

    def get_nozzle_categories(self):
        return self._lib.get_nozzle_categories()

    def get_nozzle_category_color(self, category_id: str) -> tuple[int, int, int] | None:
        """카테고리에 저장된 color_rgb 튜플 반환. 없으면 None (호출부가 해시 폴백 사용)."""
        return self._lib.get_nozzle_category_color(category_id)

    def get_head_items(self):
        return self._lib.get_head_items()

    def get_node_type_combo_items(self) -> list[str | None]:
        """제어판/팝업 공용 — 노드속성 콤보박스 항목 목록.
        None 은 구분선(separator)을 의미한다.
        섹션: 기본 | Head/노즐 | 방재밸브 | 체크밸브 | 개폐밸브 | 특수밸브 | Prv/Orifice | Pump/WT | Hose Stream

        밸브 카테고리는 domain.constants.NODE_PLACEABLE_VALVE_CATEGORIES SSOT를 따른다.
        표시명은 본 함수 내부 매핑(_VALVE_DISPLAY_EN)으로 결정한다 — i18n은 호출자(메뉴/콤보)가 처리.
        """
        from services.i18n_service import _t
        from domain.constants import NODE_PLACEABLE_VALVE_CATEGORIES, FittingCategory

        # 본 메서드 전용 영문 표시명 매핑 (D6 — domain 계층에 두지 않음)
        # 새 카테고리 추가 시 NODE_PLACEABLE_VALVE_CATEGORIES tuple과 본 dict 양쪽 갱신 필요.
        # 누락은 Stage 2 회귀 테스트가 검출한다.
        _VALVE_DISPLAY_EN: dict[FittingCategory, str] = {
            FittingCategory.ALARM_VALVE:        "Alarm Valve",
            FittingCategory.FLOW_SWITCH_VANE:   "Vane Type Flow Switch",
            FittingCategory.PREACTION_VALVE:    "Preaction Valve",
            FittingCategory.DRY_PIPE_VALVE:     "Dry Pipe Valve",
            FittingCategory.CHECK_VALVE_SWING:  "Swing Check",
            FittingCategory.SMOLENSKY_CHECK:    "Smolensky Check",
            FittingCategory.GATE_VALVE:         "Gate Valve",
            FittingCategory.BUTTERFLY_VALVE:    "Butterfly Valve",
            FittingCategory.GLOBE_VALVE:        "Globe Valve",
            FittingCategory.ANGLE_VALVE:        "Angle Valve",
            FittingCategory.BALL_VALVE:         "Ball Valve",
            FittingCategory.STRAINER:           "Strainer",
        }

        nozzle_names = [
            cat.get("display_name", "")
            for cat in self._lib.get_nozzle_categories()
            if cat.get("category_id") != "head" and cat.get("display_name")
        ]

        # 섹션 분할: 방재(0~3) | 체크(4~5) | 개폐(6~10) | 특수(11)
        valve_displays = [_VALVE_DISPLAY_EN[c] for c in NODE_PLACEABLE_VALVE_CATEGORIES]
        sec_alarm   = valve_displays[0:4]
        sec_check   = valve_displays[4:6]
        sec_open    = valve_displays[6:11]
        sec_special = valve_displays[11:]

        result: list[str | None] = []
        result.append(_t("기본"))
        result.append(None)              # ── Default 아래
        result.append("Head")
        result.extend(nozzle_names)
        result.append(None)              # ── Hose Nozzle 아래
        result.extend(sec_alarm)
        result.append(None)              # ── 방재 아래
        result.extend(sec_check)
        result.extend(sec_open)
        result.append(None)              # ── 개폐 아래
        result.extend(sec_special)
        result.append(None)              # ── 특수 아래
        result.append("Prv")
        result.append("Orifice")
        result.append(None)              # ── Orifice 아래
        result.append("Pump")
        result.append("WT")
        result.append(None)              # ── WT 아래
        result.append("Hose Stream")
        return result

    def save_pipe_library(self, to_factory: bool = False):
        return self._lib.save_pipe_library(to_factory=to_factory)

    def save_nozzle_library(self, to_factory: bool = False):
        return self._lib.save_nozzle_library(to_factory=to_factory)

    def save_pump_library(self, to_factory: bool = False):
        return self._lib.save_pump_library(to_factory=to_factory)

    def save_fittings_library_v3(self, to_factory: bool = False):
        return self._lib.save_fittings_library_v3(to_factory=to_factory)

    def save_pipe_sizing_library(self, to_factory: bool = False):
        return self._lib.save_pipe_sizing_library(to_factory=to_factory)

    def load_pipe_sizing_library_from_disk(self):
        return self._lib.load_pipe_sizing_library_from_disk()

    def get_head_range_presets(self):
        return self._lib.get_head_range_presets()

    def get_pump_list(self) -> list[dict]:
        return self._lib.get_pump_list()

    def get_pump_spec(self, pump_id: str) -> dict | None:
        return self._lib.get_pump_spec(pump_id)

    # ==================================================================
    # 서브트리 복사/붙여넣기
    # ==================================================================

    def copy_subtree_from_node(self, root_name, toward_neighbor_name, silent=False):
        if not self.graph.has_node(root_name):
            return False

        if toward_neighbor_name:
            subtree_nodes = self.get_subtree_one_direction(root_name, toward_neighbor_name)
        else:
            subtree_nodes = self.get_subtree_nodes(root_name)

        if not subtree_nodes:
            return False

        root_node = self.graph.get_node(root_name)
        if root_node is None:
            return False
        root_pos = tuple(float(v) for v in root_node.coords)

        nodes_offsets = {}
        node_types_clip = {}
        clip_node_state = {}

        for n in subtree_nodes:
            node_obj = self.graph.get_node(n)
            if node_obj is None:
                continue
            x, y, z = tuple(float(v) for v in node_obj.coords)
            dx = x - root_pos[0]
            dy = y - root_pos[1]
            dz = z - root_pos[2]
            nodes_offsets[n] = (dx, dy, dz)
            node_types_clip[n] = str(node_obj.type or "기본")

            clip_node_state[n] = node_obj.to_dict()

        pipes = []
        clip_pipe_data = {}

        for pid, p_obj in self.graph.pipes.items():
            s0 = p_obj.start
            e0 = p_obj.end
            if s0 in subtree_nodes and e0 in subtree_nodes:
                pipes.append((s0, e0))
                k = self._pipe_key(s0, e0)
                data_to_save = {
                    "type": p_obj.type,
                    "diameter": p_obj.diameter,
                    "nominal_mm": p_obj.nominal_mm,
                    "C": p_obj.C,
                    "equivalent_length": p_obj.equivalent_length,
                    "roughness_mm": p_obj.roughness_mm
                }
                clip_pipe_data[k] = data_to_save

        self.subtree_clipboard = {
            "root": root_name,
            "nodes_offsets": nodes_offsets,
            "node_types": node_types_clip,
            "node_state": clip_node_state,
            "pipes": pipes,
            "pipe_data": clip_pipe_data,
        }

        self.subtree_paste_pending = True
        return True

    def paste_subtree_to_node(self, target_root_name):
        from services.i18n_service import _t
        self.subtree_paste_pending = False
        if self.subtree_clipboard is None:
            return False
        if not self.graph.has_node(target_root_name):
            return False

        clip = self.subtree_clipboard
        root_orig = clip.get("root") or clip.get("root_node")
        if not root_orig:
            return False

        nodes_offsets = clip.get("nodes_offsets", {})
        node_types_clip = clip.get("node_types", {})
        clip_node_state = clip.get("node_state", {})
        pipes = clip.get("pipes", [])
        clip_pipe_data = clip.get("pipe_data", {})

        backup_snapshot = self.to_dict()
        is_success = True
        fail_reason = ""

        try:
            target_root = self.graph.get_node(target_root_name)
            if target_root is None:
                return False
            root_target_pos = tuple(float(v) for v in target_root.coords)
            mapping = {}
            mapping[root_orig] = target_root_name

            for orig_name, offset in nodes_offsets.items():
                if orig_name == root_orig:
                    continue

                dx, dy, dz = offset
                x, y, z = self._snap_point(
                    root_target_pos[0] + dx,
                    root_target_pos[1] + dy,
                    root_target_pos[2] + dz,
                )
                existing = self.find_node_at_position(x, y, z)
                attr = node_types_clip.get(orig_name, "기본")
                saved_state = clip_node_state.get(orig_name, {})

                if existing is not None:
                    new_name = existing
                else:
                    new_name = self.generate_node_name("N")
                    self.add_node(new_name, x, y, z, node_type=attr)
                    if saved_state:
                        # 붙여넣기 배치는 root 기준 평행이동 결과가 좌표 SSOT여야 하므로
                        # 원본 절대좌표/ID를 메타 복원에서 제외한다.
                        state_meta = {
                            k: copy.deepcopy(v) for k, v in saved_state.items()
                            if k not in {"id", "coords", "elevation_m"}
                        }
                        self.update_node_via_graph(new_name, meta=state_meta)
                        self.sanitize_node_data(new_name)

                mapping[orig_name] = new_name

            for s_orig, e_orig in pipes:
                s_new = mapping.get(s_orig)
                e_new = mapping.get(e_orig)
                if not s_new or not e_new or s_new == e_new:
                    continue

                orig_key = self._pipe_key(s_orig, e_orig)
                pdata = clip_pipe_data.get(orig_key, {})

                # 비평행 교차 사전 검사: 분할 없이 교차하는 경우 차단
                axis = self._infer_pipe_axis(s_new, e_new)
                if axis:
                    n1 = self.graph.get_node(s_new)
                    n2 = self.graph.get_node(e_new)
                    if n1 is None or n2 is None:
                        is_success = False
                        fail_reason = _t("붙여넣기 불가 ({s}↔{e}): 노드 복원 실패").format(s=s_new, e=e_new)
                        break
                    p1 = tuple(float(v) for v in n1.coords)
                    p2 = tuple(float(v) for v in n2.coords)
                    if self._find_segment_intersections(p1, p2, axis):
                        is_success = False
                        fail_reason = _t("붙여넣기 불가 ({s}↔{e}): 기존 배관과 교차합니다.\n교차 위치에 노드를 만들고 붙여넣어 주세요.").format(s=s_new, e=e_new)
                        break

                success, msg = self.create_pipe(s_new, e_new, attributes=pdata, silent=True)
                if not success:
                    is_success = False
                    fail_reason = _t("배관 생성 실패 ({s}↔{e}): {msg}").format(s=s_new, e=e_new, msg=msg)
                    break

            if is_success:
                rep = self.validate_and_fix_integrity(strict=True)
                if rep.get("errors"):
                    is_success = False
                    fail_reason = _t("무결성 검사 실패:") + "\n" + "\n".join(rep["errors"][:5])

        except Exception as e:
            is_success = False
            fail_reason = _t("시스템 오류: {e}").format(e=str(e))

        if not is_success:
            self.from_dict(backup_snapshot)
            self._last_paste_error = fail_reason
            return False

        return True

    # ==================================================================
    # 내부 편의 함수
    # ==================================================================

    def _show_msg(self, title, icon, text):
        logger.debug("[Engine Message Blocked] %s: %s", title, text)

    # ==================================================================
    # 직렬화 ([2-2] ProjectSerializer에 위임)
    # ==================================================================

    def to_dict(self):
        return self._serializer.to_dict()

    def to_runtime_snapshot(self):
        return self._serializer.to_runtime_snapshot()

    def from_runtime_snapshot(self, snapshot: dict):
        self._serializer.from_runtime_snapshot(snapshot)

    def to_json_serializable(self):
        return self._serializer.to_json_serializable()

    def from_dict(self, data: dict):
        self._serializer.from_dict(data)

    def _rebind_managers(self):
        """
        from_dict / from_runtime_snapshot 후 매니저의 참조를 갱신.
        (dict 객체가 교체되었을 수 있으므로)
        """
        self._node_mgr.node_counter = self.node_counter
        self._node_mgr.graph = self.graph
        self._node_mgr.grid_step = self.grid_step
        self._node_mgr.snap_tol = self.grid_step * 0.45

        self._pipe_mgr.pipe_registry = self.graph.pipes
        self._pipe_mgr.graph = self.graph
        self._pipe_mgr.node_mgr = self._node_mgr
        self._pipe_mgr.grid_step = self.grid_step
        self._pipe_mgr.intersection_tol = self.intersection_tol
        self._pipe_mgr.pipe_id_counter = self.pipe_id_counter
        self._pipe_mgr._design_settings = getattr(self, "design_settings", {})
        self._pipe_mgr._current_dn = getattr(self, "current_dn", "DN25")
        self._pipe_mgr._pipe_spec_getter = getattr(self, "get_pipe_spec", None)

    def save_json(self, filepath: str):
        self._serializer.save_json(filepath)

    def load_json(self, filepath: str, strict: bool = False):
        return self._serializer.load_json(filepath, strict=strict)
