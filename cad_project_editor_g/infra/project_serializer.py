"""
ProjectSerializer — PipeEditor 직렬화/역직렬화 담당

Phase 2-2 리팩토링: editor_core.py에서 직렬화 메서드 추출.
PipeEditor 인스턴스를 참조하여 상태를 읽고 씁니다.
PySide6 의존성 없음 (순수 infra 계층).
"""
import base64
import copy
import json
import logging
import zlib

from version import ANTIFREEZE_PROJECT_SCHEMA_VERSION, PROJECT_SCHEMA_VERSION

_log = logging.getLogger(__name__)

import domain.models as models
from domain.component_spec_adapter import COMPONENT_SPEC_KEYS, normalize_node_meta_for_load

NODE_STATE_KEY = "nodes" + "_meta"
NODE_STATE_RUNTIME_KEY = "nodes" + "_meta_runtime"

# 지원하는 최소 프로젝트 스키마 major 버전.
# 역사적으로 존재한 포맷은 "2.2-NODE-SCHEMA"(폐기 대상)와 "3.1-NFPA13-EQL"(현행)뿐이다.
# 정책(ADR-004): 2.x 이하는 지원하지 않고, 3.x 이상만 로드를 허용한다.
MIN_SUPPORTED_SCHEMA_MAJOR = 3

# [migration S2] 알려진 최대 스키마 major. 이보다 높은 major(미래 버전)는
# "더 신버전 파일" 로 명확히 거부한다(상위 게이트). 과거 3.1 빌드에는 이 게이트가
# 없어 4.0 파일을 빈 노드로 오독하지만(ADR에서 수용), 신 빌드는 미래 bump를 깔끔히 막는다.
KNOWN_MAX_SCHEMA_MAJOR = 5


class ProjectSerializer:
    """PipeEditor 상태를 JSON 직렬화/역직렬화하는 클래스."""

    def __init__(self, editor) -> None:
        self._editor = editor

    # =========================================================
    # 직렬화 (editor → dict)
    # =========================================================

    def to_dict(self) -> dict:
        e = self._editor
        e.sync_type_ids(verbose=False)

        serialized_pipes = {
            pid: p.to_legacy_dict() for pid, p in e.graph.pipes.items()
        }

        # [migration S2 / 4.0] 단일 노드 블록(SSOT): nodes_meta_runtime 하나만 기록한다.
        # node.to_dict()가 id/coords/elevation_m/type/type_id + 모든 장치 필드를 담으므로
        # 구 4-블록(nodes·node_types·nodes_meta)이 보유하던 정보를 전부 포섭한다.
        runtime_meta_copy: dict = {
            name: node_obj.to_dict()
            for name, node_obj in e.graph.nodes.items()
        }

        design_settings: dict = {}
        if isinstance(getattr(e, "design_settings", None), dict):
            try:
                design_settings = copy.deepcopy(e.design_settings)
            except Exception:
                design_settings = dict(e.design_settings)

        rc = getattr(e, "report_common", None)
        antifreeze_case = getattr(e, "antifreeze_analysis_case", None)
        project_version = (
            ANTIFREEZE_PROJECT_SCHEMA_VERSION
            if antifreeze_case is not None
            else PROJECT_SCHEMA_VERSION
        )

        payload = {
            "pipe_data": serialized_pipes,
            "pipe_id_counter": e.pipe_id_counter,
            "node_counter": e.node_counter.copy(),
            NODE_STATE_RUNTIME_KEY: runtime_meta_copy,
            "design_settings": design_settings,
            "report_common": copy.deepcopy(rc) if isinstance(rc, dict) else {},
            "pipe_library": copy.deepcopy(e.pipe_library) if e.pipe_library else {},
            "nozzle_library": copy.deepcopy(e.nozzle_library) if e.nozzle_library else {},
            "pump_library": copy.deepcopy(e.pump_library) if e.pump_library else {},
            "pipe_sizing_library": (
                copy.deepcopy(e.pipe_sizing_library) if e.pipe_sizing_library else {}
            ),
            "fittings_library_v3": [
                item.to_dict() for item in (e._lib.fittings_library_v3 or [])
            ],
            "version": project_version,
        }
        if antifreeze_case is not None:
            payload["antifreeze_analysis_case"] = antifreeze_case.to_dict()
            report_path = getattr(e, "antifreeze_report_pdf_path", None)
            if report_path:
                payload["antifreeze_report_pdf_path"] = str(report_path)
        return payload

    def to_runtime_snapshot(self) -> dict:
        e = self._editor
        # graph가 SSOT이므로 런타임 스냅샷은 graph 중심으로 저장한다.
        return {
            "graph": copy.deepcopy(e.graph.to_dict()),
            "pipe_id_counter": int(getattr(e, "pipe_id_counter", 0)),
            "node_counter": copy.deepcopy(getattr(e, "node_counter", {})),
            "design_settings": copy.deepcopy(getattr(e, "design_settings", {})),
        }

    @staticmethod
    def _validate_and_normalize_runtime_snapshot(snapshot: dict) -> dict:
        """
        Runtime snapshot 무결성 검증 + 정규화.
        - 필수 키/타입 검사
        - 노드 좌표 길이(3) 및 숫자형 강제
        - coords를 tuple로 정규화
        """
        if not isinstance(snapshot, dict):
            raise TypeError("runtime snapshot must be dict")

        graph = snapshot.get("graph")
        if not isinstance(graph, dict):
            raise ValueError("runtime snapshot missing 'graph' dict")

        nodes = graph.get("nodes")
        pipes = graph.get("pipes")
        if not isinstance(nodes, dict):
            raise ValueError("'graph.nodes' must be dict")
        if not isinstance(pipes, dict):
            raise ValueError("'graph.pipes' must be dict")

        normalized_nodes: dict = {}
        for node_id, node_data in nodes.items():
            if not isinstance(node_data, dict):
                raise ValueError(f"node '{node_id}' payload must be dict")

            coords = node_data.get("coords")
            if not isinstance(coords, (list, tuple)) or len(coords) != 3:
                raise ValueError(f"node '{node_id}' coords must be length-3 list/tuple")

            try:
                x = float(coords[0])
                y = float(coords[1])
                z = float(coords[2])
            except (TypeError, ValueError) as exc:
                raise ValueError(f"node '{node_id}' coords must be numeric") from exc

            node_copy = dict(node_data)
            node_copy["coords"] = (x, y, z)  # tuple 정규화
            normalized_nodes[str(node_id)] = node_copy

        normalized_pipes: dict = {}
        for pipe_id, pipe_data in pipes.items():
            if not isinstance(pipe_data, dict):
                raise ValueError(f"pipe '{pipe_id}' payload must be dict")
            start_node = pipe_data.get("start")
            end_node = pipe_data.get("end")
            if (
                not isinstance(start_node, str)
                or not isinstance(end_node, str)
                or not start_node
                or not end_node
            ):
                raise ValueError(f"pipe '{pipe_id}' must have non-empty string start/end")
            normalized_pipes[str(pipe_id)] = dict(pipe_data)

        node_counter = snapshot.get("node_counter", {})
        design_settings = snapshot.get("design_settings", {})
        if not isinstance(node_counter, dict):
            raise ValueError("'node_counter' must be dict")
        if not isinstance(design_settings, dict):
            raise ValueError("'design_settings' must be dict")

        try:
            pipe_id_counter = int(snapshot.get("pipe_id_counter", 0))
        except (TypeError, ValueError) as exc:
            raise ValueError("'pipe_id_counter' must be int-castable") from exc

        return {
            "graph": {
                "nodes": normalized_nodes,
                "pipes": normalized_pipes,
            },
            "pipe_id_counter": pipe_id_counter,
            "node_counter": dict(node_counter),
            "design_settings": dict(design_settings),
        }

    @staticmethod
    def encode_runtime_snapshot_blob(snapshot: dict, as_text: bool = False) -> bytes | str:
        """
        snapshot(dict) -> compressed blob
        - as_text=False: bytes(zlib)
        - as_text=True : base64 str
        """
        normalized = ProjectSerializer._validate_and_normalize_runtime_snapshot(snapshot)
        raw = json.dumps(normalized, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        packed = zlib.compress(raw, level=6)
        if as_text:
            return base64.b64encode(packed).decode("ascii")
        return packed

    @staticmethod
    def decode_runtime_snapshot_blob(blob: bytes | str | dict) -> dict:
        """
        blob(bytes/str/dict) -> normalized snapshot(dict).
        dict 입력은 하위호환을 위해 그대로 검증 후 반환.
        """
        if isinstance(blob, dict):
            return ProjectSerializer._validate_and_normalize_runtime_snapshot(blob)

        if isinstance(blob, str):
            try:
                packed = base64.b64decode(blob.encode("ascii"))
            except Exception as exc:
                raise ValueError("invalid base64 runtime snapshot blob") from exc
        elif isinstance(blob, (bytes, bytearray)):
            packed = bytes(blob)
        else:
            raise TypeError("runtime snapshot blob must be bytes|str|dict")

        try:
            raw = zlib.decompress(packed)
        except Exception as exc:
            raise ValueError("invalid compressed runtime snapshot blob") from exc

        try:
            payload = json.loads(raw.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise ValueError("invalid runtime snapshot JSON payload") from exc

        return ProjectSerializer._validate_and_normalize_runtime_snapshot(payload)

    def to_runtime_snapshot_blob(self, as_text: bool = False) -> bytes | str:
        """현재 editor 상태를 즉시 blob으로 생성."""
        snapshot = self.to_runtime_snapshot()
        return self.encode_runtime_snapshot_blob(snapshot, as_text=as_text)

    def from_runtime_snapshot_blob(self, blob: bytes | str | dict) -> None:
        """
        blob 복원 진입점.
        검증 실패 시 예외를 그대로 올려 Undo/Redo를 중단한다.
        """
        snapshot = self.decode_runtime_snapshot_blob(blob)
        self.from_runtime_snapshot(snapshot)

    def to_json_serializable(self) -> dict:
        e = self._editor
        # [DR-3] graph가 SSOT이므로 node.to_dict()를 직접 사용 (side-effect 없음)
        serialized_pipes = {
            pid: p.to_legacy_dict() for pid, p in e.graph.pipes.items()
        }
        return {
            "nodes": {name: list(node.coords) for name, node in e.graph.nodes.items()},
            "node_types": {name: str(node.type) for name, node in e.graph.nodes.items()},
            NODE_STATE_KEY: {name: node.to_dict() for name, node in e.graph.nodes.items()},
            "pipe_data": serialized_pipes,
            "design_settings": dict(getattr(e, "design_settings", {}) or {}),
        }

    # =========================================================
    # 역직렬화 (dict → editor 상태)
    # =========================================================

    @staticmethod
    def _parse_schema_major(version: str) -> int | None:
        """버전 문자열의 선두 major 정수를 추출. 파싱 불가 시 None.

        예) "3.1-NFPA13-EQL" -> 3,  "2.2-NODE-SCHEMA" -> 2
        """
        head = str(version).strip().split(".", 1)[0]
        digits = ""
        for ch in head:
            if ch.isdigit():
                digits += ch
            else:
                break
        return int(digits) if digits else None

    def from_dict(self, data: dict) -> None:
        # 구 버전 파일 감지 (v2.0 이전: version 키 없음)
        version = data.get("version", "")
        if not version:
            from services.i18n_service import _t
            raise ValueError(
                _t("이 파일은 K-Fire v2.0 이전 버전 형식입니다.\nv2.0에서는 열 수 없습니다.\n\nv2.0에서 새로 작성해 주세요.")
            )

        # [ADR-004] 3.x-only 버전 게이트: 2.x 이하(및 파싱 불가 버전)는 명시적으로 거부.
        major = self._parse_schema_major(version)
        if major is None or major < MIN_SUPPORTED_SCHEMA_MAJOR:
            from services.i18n_service import _t
            raise ValueError(
                _t("지원하지 않는 프로젝트 버전입니다: {version}\n현재 버전은 3.x 이상 형식만 열 수 있습니다.").format(version=version)
            )

        # [migration S2] 상위 게이트: 알려진 최대 major보다 높은 파일은 명확히 거부한다.
        # (구 빌드가 신버전 파일을 빈 노드로 오독하는 사고를 미래 bump부터 차단)
        if major > KNOWN_MAX_SCHEMA_MAJOR:
            from services.i18n_service import _t
            raise ValueError(
                _t("더 신버전에서 저장된 프로젝트입니다: {version}\n프로그램을 최신 버전으로 업데이트한 뒤 열어 주세요.").format(version=version)
            )

        e = self._editor

        raw_antifreeze_case = data.get("antifreeze_analysis_case")
        if major >= 5:
            if not isinstance(raw_antifreeze_case, dict):
                raise ValueError("major-5 project missing antifreeze_analysis_case")
            from domain.antifreeze import AntifreezeAnalysisCase

            e.antifreeze_analysis_case = AntifreezeAnalysisCase.from_dict(
                raw_antifreeze_case
            )
            raw_report_path = data.get("antifreeze_report_pdf_path")
            e.antifreeze_report_pdf_path = (
                str(raw_report_path).strip()
                if isinstance(raw_report_path, str) and raw_report_path.strip()
                else None
            )
        else:
            if raw_antifreeze_case is not None:
                raise ValueError("antifreeze_analysis_case requires project major 5")
            e.antifreeze_analysis_case = None
            e.antifreeze_report_pdf_path = None

        base_design = getattr(
            e,
            "design_settings",
            {"min_required_pressure_bar": 1.0, "calculation_method": "H-W"},
        )
        loaded_design = data.get("design_settings")
        if isinstance(loaded_design, dict):
            tmp = base_design.copy()
            tmp.update(loaded_design)
            e.design_settings = tmp
        else:
            e.design_settings = base_design

        loaded_rc = data.get("report_common")
        if isinstance(loaded_rc, dict):
            e.report_common = loaded_rc
        else:
            e.report_common = {"design_area": "", "company": "", "project": "", "date": "", "notes": ""}

        # 노드 로드용 로컬 스테이징 dict
        _staged_meta: dict = {}

        raw_runtime_meta = data.get(NODE_STATE_RUNTIME_KEY) or {}
        if isinstance(raw_runtime_meta, dict):
            for k, v in raw_runtime_meta.items():
                try:
                    _staged_meta[str(k)] = copy.deepcopy(normalize_node_meta_for_load(v))
                except Exception:
                    _staged_meta[str(k)] = normalize_node_meta_for_load(dict(v))

        # 구조화 노드 상태로 누락 필드 최소 보충 (type_id, type, required_pressure_bar)
        raw_struct_meta = data.get(NODE_STATE_KEY) or {}
        if isinstance(raw_struct_meta, dict):
            for name, m in raw_struct_meta.items():
                if not isinstance(m, dict):
                    continue
                normalized_struct_meta = normalize_node_meta_for_load(m)
                n = str(name)
                if n not in _staged_meta or not isinstance(_staged_meta.get(n), dict):
                    _staged_meta[n] = {}

                if any(key in m for key in COMPONENT_SPEC_KEYS):
                    for key, value in normalized_struct_meta.items():
                        _staged_meta[n].setdefault(key, copy.deepcopy(value))

                if "required_pressure_bar" not in _staged_meta[n]:
                    req = normalized_struct_meta.get("required_pressure_bar")
                    if req is not None:
                        try:
                            _staged_meta[n]["required_pressure_bar"] = float(req)
                        except Exception:
                            pass

                tid = normalized_struct_meta.get("type_id")
                if tid is not None and str(tid).strip():
                    current_tid = str(_staged_meta[n].get("type_id") or "").strip()
                    if not current_tid:
                        _staged_meta[n]["type_id"] = str(tid).strip()

                t = normalized_struct_meta.get("type")
                if t is not None and str(t).strip():
                    if _staged_meta[n].get("type") in (None, "", "기본"):
                        _staged_meta[n]["type"] = str(t).strip()

        e.grid_step = float(data.get("grid_step", getattr(e, "grid_step", 0.1)))
        e.snap_tol = float(data.get("snap_tol", e.grid_step * 0.45))
        e.intersection_tol = float(
            data.get("intersection_tol", getattr(e, "intersection_tol", 1e-6))
        )

        raw_nodes = data.get("nodes", {})
        raw_types = data.get("node_types", {})
        if not isinstance(raw_nodes, dict):
            raw_nodes = {}
        if not isinstance(raw_types, dict):
            raw_types = {}

        # D2-a: 복원은 Graph를 우선 소스로 구성한다.
        e.graph.clear()
        e._load_skipped_nodes = []

        # [migration S2] 노드 반복 소스는 버전 major로 분기한다.
        #  · major == 3 (레거시): 4-블록 포맷 — 노드 존재성은 nodes(coords) 블록이 키잉.
        #    기존 사용자 .kfp는 전부 3.x이므로 이 경로는 영구 보존(S1 오라클이 동결).
        #  · major >= 4 (신포맷): 단일 nodes_meta_runtime 블록이 유일한 노드 소스.
        #    coords/type을 각 엔트리 자체에서 읽는다(nodes/node_types 조회 없음).
        if major >= 4:
            v4_node_source = raw_runtime_meta if isinstance(raw_runtime_meta, dict) else {}
            for name, raw_meta in v4_node_source.items():
                if not isinstance(raw_meta, dict):
                    continue
                coord = raw_meta.get("coords")
                try:
                    x, y, z = coord
                except Exception as _e:
                    _log.error("노드 좌표 파싱 실패 [노드=%r, coord=%r]: 이 노드는 로드에서 제외됩니다. 원인: %s", name, coord, _e)
                    e._load_skipped_nodes.append(str(name))
                    continue

                node_name = str(name)
                node_type = str(raw_meta.get("type", "기본"))
                node_obj = models.Node.create(node_name, float(x), float(y), float(z), node_type=node_type)

                staged = _staged_meta.get(node_name, {})
                if isinstance(staged, dict):
                    node_obj.update_from_meta(staged)

                e.graph.add_node(node_obj)
        else:
            for name, coord in raw_nodes.items():
                if coord is None:
                    continue
                try:
                    x, y, z = coord
                except Exception as _e:
                    _log.error("노드 좌표 파싱 실패 [노드=%r, coord=%r]: 이 노드는 로드에서 제외됩니다. 원인: %s", name, coord, _e)
                    e._load_skipped_nodes.append(str(name))
                    continue

                node_name = str(name)
                node_type = str(raw_types.get(node_name, "기본"))
                node_obj = models.Node.create(node_name, float(x), float(y), float(z), node_type=node_type)

                staged = _staged_meta.get(node_name, {})
                if isinstance(staged, dict):
                    node_obj.update_from_meta(staged)

                e.graph.add_node(node_obj)

        e.node_counter = dict(data.get("node_counter", {}))

        # [보강] type_id 강제 복원 ─ 두 가지 경로로 방어
        # ① 레거시 node_type_ids 딕셔너리 (구 포맷 호환)
        legacy_tids = data.get("node_type_ids") or {}
        if isinstance(legacy_tids, dict):
            for _n, _tid in legacy_tids.items():
                _n = str(_n)
                if _tid and str(_tid).strip() not in ("", "base"):
                    _g = e.graph.get_node(_n)
                    if _g and str(_g.type_id or "base").lower() == "base":
                        e.graph.update_node(_n, type_id=str(_tid).strip())

        # ② staged_meta에 헤드/노즐 고유 필드가 있는데 type_id가 여전히 base이면 재추론
        for _n, _staged in _staged_meta.items():
            if not isinstance(_staged, dict):
                continue
            _g = e.graph.get_node(_n)
            if _g is None:
                continue
            if str(_g.type_id or "base").lower() in ("", "base"):
                _disp = str(_g.type or "")
                _inferred = models.normalize_node_type_id(_disp, _staged)
                if _inferred and _inferred != "base":
                    e.graph.update_node(_n, type_id=_inferred)

        loaded_pipe_data = data.get("pipe_data")
        if not isinstance(loaded_pipe_data, dict):
            from services.i18n_service import _t
            raise ValueError(
                _t("지원하지 않는 프로젝트 포맷입니다: 'pipe_data'가 필요합니다.")
            )

        e.pipe_id_counter = int(data.get("pipe_id_counter", 0))
        for pid, pdata in loaded_pipe_data.items():
            s = pdata.get("start")
            end_node = pdata.get("end")
            if not s or not end_node:
                continue
                
            # [버그 수정] 존재하지 않는 노드를 참조하는 유령 배관(Dangling Pipe) 생성 방지
            if not e.graph.has_node(s) or not e.graph.has_node(end_node):
                continue

            new_pipe = models.Pipe(pid, s, end_node)
            new_pipe.type = pdata.get("type") or ""
            new_pipe.diameter = float(pdata.get("diameter") or 0.0)
            new_pipe.nominal_mm = int(pdata.get("nominal_mm") or 0)
            new_pipe.length_m = float(pdata.get("length_m") or 0.0)
            new_pipe.equivalent_length = float(pdata.get("equivalent_length") or 0.0)
            new_pipe.C = float(pdata.get("C") or 120.0)
            new_pipe.flow_lpm = float(pdata.get("flow_lpm") or 0.0)
            new_pipe.velocity_mps = float(pdata.get("velocity_mps") or 0.0)
            new_pipe.headloss_m = float(pdata.get("headloss_m") or 0.0)
            new_pipe.fittings = pdata.get("fittings") or []
            new_pipe.roughness_mm = float(pdata.get("roughness_mm") or 0.085)

            e.graph.add_pipe(new_pipe)

        e._rebind_managers()
        e._mark_pipe_key_index_dirty()

        try:
            e.update_all_pipe_lengths_in_pipe_data()
        except Exception as _e:
            _log.error("배관 길이 재계산 실패 (로드 후처리): %s", _e)

        if hasattr(e, "_sync_all_legacy_views_from_graph"):
            e._sync_all_legacy_views_from_graph()

        # 프로젝트 파일 내장 라이브러리로 메모리 교체
        e.pipe_library = data["pipe_library"]

        # standard_meta 정규화 — 구버전 .kfp 또는 메타 없는 임베드 라이브러리 1회 보완
        # (ui.tables.pipe_library_grid_helpers는 PySide6 의존 0 — infra 계층에서 import 허용)
        _pl = e.pipe_library
        if isinstance(_pl, dict) and not (isinstance(_pl.get("standard_meta"), dict) and _pl.get("standard_meta")):
            try:
                from ui.tables.pipe_library_grid_helpers import normalize_standard_meta as _nsm
                _stds = list({r.get("standard", "") for r in _pl.get("pipe_types", []) if r.get("standard")})
                _pl["standard_meta"] = _nsm(None, _stds)
            except Exception as _e:
                _log.warning("[LIB] pipe_library standard_meta 정규화 실패: %s", _e)

        e.nozzle_library = data["nozzle_library"]
        e.pump_library = data.get("pump_library") or {"pumps": []}

        # 배관구경 산정(헤드갯수기준) 라이브러리 복원
        # - .kfp에 키가 있으면 그걸 메모리에 로드 (타 유저 기준 공유)
        # - 없으면(구버전 .kfp) 전역 파일→내장 폴백 (v3 부속과 동일 규칙)
        ps_payload = data.get("pipe_sizing_library")
        if isinstance(ps_payload, dict) and ps_payload.get("presets"):
            e.pipe_sizing_library = ps_payload
        else:
            e.load_pipe_sizing_library_from_disk()

        # v3 피팅 라이브러리 복원
        # - .kfp에 fittings_library_v3 키가 있으면 그걸 메모리에 로드
        # - 없으면 (구버전 .kfp 또는 v3.0 초기 저장본) 시드 디폴트 fallback
        v3_payload = data.get("fittings_library_v3")
        if isinstance(v3_payload, list) and v3_payload:
            try:
                from domain.constants import DN_STANDARD
                from domain.fitting_models import (
                    FittingDefinition,
                    make_user_null_cell,
                    validate_library,
                )
                restored_v3 = [
                    FittingDefinition.from_dict(item) for item in v3_payload
                ]
                for fitting in restored_v3:
                    if fitting.source.get("kind") == "USER":
                        for dn in DN_STANDARD:
                            fitting.equivalent_lengths.setdefault(dn, make_user_null_cell())
                validate_library(restored_v3)
                e._lib.fittings_library_v3 = restored_v3
            except Exception as _e:
                _log.warning("[LIB] .kfp embed v3 라이브러리 복원 실패 → 시드 디폴트로 fallback: %s", _e)
                v3_payload = None  # fallback 분기로 진입
        if not (isinstance(v3_payload, list) and v3_payload):
            try:
                _v3_path = e._lib._base_path() / "fittings_library_v3.json"
                e._lib.load_fittings_v3(_v3_path)
            except Exception as _e:
                _log.warning("[LIB] v3 부속 라이브러리 로딩 실패 (from_dict): %s", _e)

        # [Layer 1] 내장 라이브러리 시스템 카테고리 ID 정규화
        # 구버전 .kfp 내장 라이브러리의 category_id가 현행 명명 규칙과 다를 수 있음
        # (예: Water Spray의 "nozzle" → "water_spray").
        # is_system_protected=True 카테고리만 대상 — 사용자 카테고리는 보호.
        # 노드는 아래 복원 루프에서 display_name 기준으로 별도 처리.
        e._lib.normalize_embedded_library_category_ids()

        # [Layer 1-b] 사용자 카테고리 color_rgb 자동 배정
        # color_rgb 필드 없는 사용자 노즐 카테고리(구버전 .kfp)에 순차 배정.
        # 이미 필드가 있는 카테고리는 건너뜀 — 덮어쓰기 없음.
        e._lib.assign_colors_to_user_categories()

        # [Phase II-4] category_id 복원 — 표시명 직접 매핑 (라이브러리 버전 무관)
        # 매핑 테이블은 services.library_service._SYSTEM_DISPLAY_TO_CATEGORY 로 SSOT 일원화.
        from services.library_service import _SYSTEM_DISPLAY_TO_CATEGORY
        for _n, _node in e.graph.nodes.items():
            _tid = str(_node.type_id or "base").lower()
            _cid = str(_node.category_id or "")
            # 이미 유효한 세부 카테고리가 있으면 건너뜀
            if _cid and _cid != _tid:
                continue
            _disp_lower = str(_node.type or "").lower().strip()
            _mapped = _SYSTEM_DISPLAY_TO_CATEGORY.get(_disp_lower)
            if _mapped:
                e.graph.update_node(_n, category_id=_mapped)
            elif _tid not in ("base", ""):
                e.graph.update_node(_n, category_id=_tid)

        # [Check valve SSOT] 라이브러리 복원·category 마이그레이션이 끝난 뒤
        # canonical category와 파생 플래그를 원자적으로 정규화한다.
        from services.check_valve_category_resolver import (
            resolve_node_check_valve_identity,
        )
        for _n, _node in e.graph.nodes.items():
            if str(_node.type_id or "").strip().lower() != "valve":
                continue
            _identity = resolve_node_check_valve_identity(
                _node,
                e.get_fitting_data_v3,
                warn_unresolved=True,
            )
            if not _identity.resolved:
                continue
            e.graph.update_node(
                _n,
                category_id=_identity.category_id,
                is_check_valve=_identity.is_check_valve,
            )

    def from_runtime_snapshot(self, snapshot: dict) -> None:
        if not isinstance(snapshot, dict):
            return

        e = self._editor
        graph_data = snapshot.get("graph")
        if not isinstance(graph_data, dict):
            from services.i18n_service import _t
            raise ValueError(
                _t("지원하지 않는 runtime snapshot 포맷입니다: 'graph'가 필요합니다.")
            )
        e.graph = type(e.graph).from_dict(graph_data)

        e.pipe_id_counter = int(snapshot.get("pipe_id_counter", 0))
        e.node_counter = copy.deepcopy(snapshot.get("node_counter", {}))
        e.design_settings = copy.deepcopy(snapshot.get("design_settings", {}))

        e._rebind_managers()
        e._mark_pipe_key_index_dirty()
        e.sync_type_ids(verbose=False)
        if hasattr(e, "_sync_all_legacy_views_from_graph"):
            e._sync_all_legacy_views_from_graph()

    # =========================================================
    # 파일 I/O
    # =========================================================

    def save_json(self, filepath: str) -> None:
        e = self._editor
        report = e.validate_and_fix_integrity(strict=True)
        e.last_integrity_report = report

        if report.get("errors"):
            from services.i18n_service import _t
            raise ValueError(
                _t("저장 차단: 무결성 오류") + "\n" + "\n".join(report["errors"][:15])
            )

        data = self.to_dict()
        with open(filepath, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
            _log.info("[JSON] 저장 완료: %s", filepath)

    def load_json(self, filepath: str, strict: bool = False) -> dict:
        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)

            self.from_dict(data)

            e = self._editor
            report = e.validate_and_fix_integrity(strict=strict)
            e.last_integrity_report = report

            skipped = getattr(e, "_load_skipped_nodes", [])
            if skipped:
                if not isinstance(report.get("warnings"), list):
                    report["warnings"] = []
                node_list = ", ".join(skipped[:10])
                from services.i18n_service import _t
                suffix = _t(" 외 {extra}개").format(extra=len(skipped) - 10) if len(skipped) > 10 else ""
                report["warnings"].insert(
                    0,
                    _t("로드 시 {count}개 노드가 좌표 오류로 제외됨: {nodes}{suffix}").format(
                        count=len(skipped), nodes=node_list, suffix=suffix),
                )

            if strict and report.get("errors"):
                raise ValueError("\n".join(report["errors"]))

            _log.info("[JSON] 로드 및 보정 완료: %s", filepath)
            return report

        except Exception as exc:
            _log.error("[Load Error] %s", exc)
            raise exc
