"""
domain/pipe_manager.py

배관 관련 비즈니스 로직을 담당하는 매니저 클래스.
PySide6 의존성 없음 — 순수 도메인 계층.

PipeEditor에서 배관 관련 메서드를 추출하여 단일 책임 원칙(SRP)을 적용.
"""
from __future__ import annotations

import logging
import math
from typing import TYPE_CHECKING, Any

_log = logging.getLogger(__name__)

from .models import Pipe
from .geometry import (
    TOLERANCE, is_close, is_same_point, dist_3d,
    snap_value, snap_point, get_vector, angle_between_vectors,
)

if TYPE_CHECKING:
    from .network_graph import NetworkGraph
    from .node_manager import NodeManager


class PipeManager:
    """
    배관 CRUD, 분할, 교차 감지, 속성 관리 등 배관 관련 로직 전담.

    PipeEditor가 소유하는 공유 상태(pipe_registry, nodes 등)에 대한
    참조를 받아서 동작한다.
    """

    def __init__(
        self,
        pipe_registry: dict[str, Pipe],
        graph: NetworkGraph,
        node_manager: NodeManager,
        grid_step: float = 0.01,
        intersection_tol: float = 1e-6,
    ):
        self.pipe_registry = pipe_registry
        self.graph = graph
        self.node_mgr = node_manager
        self.grid_step = grid_step
        self.intersection_tol = intersection_tol

        self.pipe_id_counter: int = 0
        self._pipe_key_to_pid: dict[tuple[str, str], str] = {}
        self._pipe_key_index_dirty: bool = True

        self._design_settings: dict = {}
        self._current_dn: str = "DN25"
        self._pipe_spec_getter = None

        # 겹침 검사 색인 — frozen_geometry() 컨텍스트 안에서만 산다(아래).
        self._geom_index: dict | None = None

    # ------------------------------------------------------------------
    # 인덱스 관리
    # ------------------------------------------------------------------

    @staticmethod
    def _pipe_key(start: str, end: str) -> tuple[str, str]:
        return tuple(sorted((start, end)))

    def _rebuild_pipe_key_index(self) -> None:
        self._pipe_key_to_pid = {}
        for pid, p in (self.graph.pipes or {}).items():
            try:
                key = self._pipe_key(p.start, p.end)
                self._pipe_key_to_pid[key] = pid
            except Exception as _e:
                _log.warning("배관 키 인덱스 빌드 실패 [pid=%s, start=%r, end=%r]: %s", pid, getattr(p, 'start', '?'), getattr(p, 'end', '?'), _e)
                continue
        self._pipe_key_index_dirty = False

    def _mark_pipe_key_index_dirty(self) -> None:
        self._pipe_key_index_dirty = True

    def get_pipe_id_by_nodes(self, n1: str, n2: str) -> str | None:
        # ★그래프의 양방향 키 색인을 그대로 쓴다. 종전에는 여기 사본
        #   (_pipe_key_to_pid)을 두고 변경 한 건마다 dirty → 다음 조회 때
        #   전체 재구축했다 — CAD 노드정리(병합 7,204건)에서 병합마다 서너 번
        #   불려 O(P) 재구축이 그만큼 반복됐다. 그래프 색인은 증분 유지되고
        #   (network_graph.add_pipe/_unindex_pipe), 평행 배관의 승자 규칙
        #   (나중 것이 이긴다)도 옛 재구축과 같다.
        if not n1 or not n2:
            return None
        return self.graph.find_pipe_by_nodes(n1, n2)

    def _refresh_pipe_indices(self, pipe_id: str) -> None:
        # pipe SSOT는 graph.pipes. 인덱스만 dirty 처리한다.
        self._mark_pipe_key_index_dirty()

    # ------------------------------------------------------------------
    # 겹침 검사 색인 — «좌표가 안 움직이는 구간» 전용
    # ------------------------------------------------------------------
    #
    # _check_collision / _check_diagonal_overlap 은 배관 «전수» 를 훑는다.
    # 편집 화면의 클릭 한 번에는 그만해도 되지만, CAD 변환의 노드정리는
    # 병합 7,204건 × 건마다 서너 번 이 검사를 부른다 — 전수 훑기가 실측
    # 937초의 몸통이었다.
    #
    # 색인 원리: 검사 대상이 될 수 있는 배관은 «같은 축 + 고정좌표 두 개가
    # is_close» (직교) 또는 «같은 대각 계열 + z·직선상수 is_close» (대각)
    # 뿐이다. 고정좌표를 2×TOLERANCE 칸으로 양자화해 담아 두면, is_close 한
    # 값은 반드시 이웃 ±1칸 안에 있다(|a−b|≤T, 칸=2T → 몫 차 ≤ 0.5+ε < 1).
    # 조회는 3×3칸의 «후보» 만 꺼내 **종전과 똑같은 판정식** 을 돌린다 —
    # 색인은 superset 필터일 뿐이라 결과가 전수 훑기와 정확히 같다.
    #
    # ★계약: 컨텍스트 동안 노드 «좌표» 가 움직이면 안 된다. 배관 추가·삭제는
    #   그래프 이벤트(on_pipe_added/on_pipe_removed)로 따라가지만, 좌표 이동은
    #   이벤트로 못 따라간다(직접 대입이 섞여 있다). 노드정리는 좌표를 안
    #   움직인다(삭제·생성뿐) — 그래서 이 컨텍스트가 안전하다.

    _GEOM_Q = 2.0 * TOLERANCE

    def _geom_key(self, pipe) -> tuple | None:
        """배관 하나의 색인 키. 검사식이 거르는 분류와 정확히 같은 기준."""
        po1 = self._get_node_coord(pipe.start)
        po2 = self._get_node_coord(pipe.end)
        if po1 is None or po2 is None:
            return None
        q = self._GEOM_Q
        axis = self._infer_pipe_axis(pipe.start, pipe.end)
        if axis == "X":
            return ("A", "X", math.floor(po1[1] / q), math.floor(po1[2] / q))
        if axis == "Y":
            return ("A", "Y", math.floor(po1[0] / q), math.floor(po1[2] / q))
        if axis == "Z":
            return ("A", "Z", math.floor(po1[0] / q), math.floor(po1[1] / q))
        if self._is_xy_diagonal(pipe.start, pipe.end):
            dx = po2[0] - po1[0]
            dy = po2[1] - po1[1]
            same_sign = (dx > 0) == (dy > 0)
            c = (po1[0] - po1[1]) if same_sign else (po1[0] + po1[1])
            return ("D", same_sign, math.floor(po1[2] / q),
                    math.floor(c / q))
        return None    # 축평행도 대각도 아님 — 어떤 검사식에도 안 걸린다

    def _geom_on_added(self, data) -> None:
        pipe_id = self._geom_event_pid(data)
        pipe = self.graph.get_pipe(pipe_id) if pipe_id else None
        if pipe is None or self._geom_index is None:
            return
        key = self._geom_key(pipe)
        if key is not None:
            self._geom_index["buckets"].setdefault(key, []).append(pipe.id)
            self._geom_index["pid_key"][pipe.id] = key

    def _geom_on_removed(self, data) -> None:
        pipe_id = self._geom_event_pid(data)
        if not pipe_id or self._geom_index is None:
            return
        key = self._geom_index["pid_key"].pop(pipe_id, None)
        if key is not None:
            bucket = self._geom_index["buckets"].get(key)
            if bucket is not None:
                try:
                    bucket.remove(pipe_id)
                except ValueError:
                    pass

    @staticmethod
    def _geom_event_pid(data) -> str | None:
        """이벤트 payload 에서 배관 id — emit 방식(단일/튜플)에 안 매인다."""
        if isinstance(data, tuple):
            for item in data:
                if isinstance(item, dict) and "id" in item:
                    return item["id"]
                if isinstance(item, str):
                    return item
            return None
        if isinstance(data, dict):
            return data.get("id")
        return data if isinstance(data, str) else None

    def frozen_geometry(self):
        """겹침 검사를 색인으로 돌리는 구간(계약은 위 블록 주석).

        중첩 진입은 바깥 색인을 그대로 쓴다. contextmanager 반환."""
        from contextlib import contextmanager

        @contextmanager
        def _ctx():
            if self._geom_index is not None:     # 중첩 — 바깥 것 재사용
                yield
                return
            buckets: dict[tuple, list[str]] = {}
            pid_key: dict[str, tuple] = {}
            for pid, p in self.graph.pipes.items():
                key = self._geom_key(p)
                if key is not None:
                    buckets.setdefault(key, []).append(pid)
                    pid_key[pid] = key
            self._geom_index = {"buckets": buckets, "pid_key": pid_key}
            self.graph.on_pipe_added.subscribe(self._geom_on_added)
            self.graph.on_pipe_removed.subscribe(self._geom_on_removed)
            try:
                yield
            finally:
                self.graph.on_pipe_added.unsubscribe(self._geom_on_added)
                self.graph.on_pipe_removed.unsubscribe(self._geom_on_removed)
                self._geom_index = None
        return _ctx()

    def _geom_candidates(self, kind, tag, f1, f2):
        """(kind, tag) 분류에서 고정값 두 개의 3×3 이웃칸 후보 배관 id."""
        q = self._GEOM_Q
        b1 = math.floor(f1 / q)
        b2 = math.floor(f2 / q)
        buckets = self._geom_index["buckets"]
        out: list[str] = []
        for d1 in (-1, 0, 1):
            for d2 in (-1, 0, 1):
                out.extend(buckets.get((kind, tag, b1 + d1, b2 + d2), ()))
        return out

    # ------------------------------------------------------------------
    # CRUD
    # ------------------------------------------------------------------

    def create_pipe(
        self,
        start: str,
        end: str,
        attributes: dict | None = None,
        silent: bool = False,
        *,
        design_settings: dict | None = None,
        current_dn: str = "DN25",
        pipe_spec_getter=None,
    ) -> tuple[bool, str]:
        """
        배관 생성 (SSOT).
        반환값: (성공여부, 메시지)

        design_settings, current_dn, pipe_spec_getter는 PipeEditor가 주입.
        명시적으로 전달하지 않으면 PipeManager 자체 속성을 fallback으로 사용.
        """
        if not self._has_node(start) or not self._has_node(end):
            return False, "노드가 존재하지 않습니다."
        if start == end:
            return False, "시점과 종점이 같습니다."
        if self.get_pipe_id_by_nodes(start, end):
            return False, f"[{start}-{end}] 배관이 이미 존재합니다."

        is_valid, msg = self.validate_new_segment(start, end)
        if not is_valid:
            return False, f"[{start} -> {end}]\n{msg}"

        effective_ds = design_settings if design_settings is not None else self._design_settings
        effective_dn = current_dn if design_settings is not None else self._current_dn
        effective_getter = pipe_spec_getter if pipe_spec_getter is not None else self._pipe_spec_getter

        final_attrs = self._resolve_pipe_attributes(
            override_attrs=attributes,
            design_settings=effective_ds,
            current_dn=effective_dn,
            pipe_spec_getter=effective_getter,
        )

        self.pipe_id_counter += 1
        pid = f"P{self.pipe_id_counter}"

        new_pipe = Pipe(pid, start, end)
        new_pipe.type = final_attrs["type"]
        new_pipe.nominal_mm = final_attrs["nominal_mm"]
        new_pipe.diameter = final_attrs["diameter"]
        new_pipe.C = final_attrs["C"]
        new_pipe.roughness_mm = final_attrs["roughness_mm"]
        new_pipe.length_m = self.compute_length(start, end)
        new_pipe.equivalent_length = float((attributes or {}).get("equivalent_length", 0.0))

        self.graph.add_pipe(new_pipe)
        self._refresh_pipe_indices(pid)

        return True, "OK"

    def add_pipe(self, start: str, end: str, silent: bool = False, **kwargs) -> bool:
        """Legacy wrapper — create_pipe의 bool만 반환."""
        success, _msg = self.create_pipe(start, end, attributes=None, silent=silent, **kwargs)
        return success

    def delete_pipe(self, start: str, end: str) -> None:
        pid = self.get_pipe_id_by_nodes(start, end)
        if pid:
            self.graph.remove_pipe(pid)
            self._refresh_pipe_indices(pid)

    def delete_pipe_data(self, pid: str) -> None:
        """호환성 유지용 빈 함수."""
        pass

    def update_pipe_data(self, pid: str, **kwargs: Any) -> None:
        pipe = self.graph.get_pipe(pid)
        if pipe is None:
            return
        patch: dict[str, Any] = {}
        if "type" in kwargs:
            patch["type"] = kwargs["type"]
        if "diameter" in kwargs:
            patch["diameter"] = float(kwargs["diameter"])
        if "nominal_mm" in kwargs:
            patch["nominal_mm"] = int(kwargs["nominal_mm"])
        if "C" in kwargs:
            patch["C"] = float(kwargs["C"])
        if "length_m" in kwargs:
            patch["length_m"] = float(kwargs["length_m"])
        if "equivalent_length" in kwargs:
            patch["equivalent_length"] = float(kwargs["equivalent_length"])
        if "roughness_mm" in kwargs:
            patch["roughness_mm"] = float(kwargs["roughness_mm"])
        if patch:
            self.graph.update_pipe(pid, **patch)
            self._refresh_pipe_indices(pid)

    def apply_pipe_attributes(self, pipe_id: str, attrs: dict) -> None:
        pipe = self.graph.get_pipe(pipe_id)
        if pipe is None:
            return
        patch: dict[str, Any] = {}
        for k, v in attrs.items():
            if hasattr(pipe, k):
                patch[k] = v
        if patch:
            self.graph.update_pipe(pipe_id, **patch)
            self._refresh_pipe_indices(pipe_id)

    # ------------------------------------------------------------------
    # 속성 결정
    # ------------------------------------------------------------------

    def _resolve_pipe_attributes(
        self,
        override_attrs: dict | None = None,
        design_settings: dict | None = None,
        current_dn: str = "DN25",
        pipe_spec_getter=None,
    ) -> dict:
        ds = design_settings or {}
        std_default = ds.get("standard", "KSD3507 (SPP)")
        c_default = float(ds.get("C", 120.0))
        rough_default = float(ds.get("roughness_mm", 0.085))

        try:
            dn_default = int(str(current_dn).upper().replace("DN", ""))
        except (ValueError, TypeError):
            dn_default = 25

        attrs = override_attrs or {}
        target_std = attrs.get("type", std_default)
        target_dn = int(attrs.get("nominal_mm", dn_default))
        target_c = float(attrs.get("C", c_default))
        target_rough = float(attrs.get("roughness_mm", rough_default))

        if "diameter" in attrs and float(attrs["diameter"]) > 0:
            target_dia = float(attrs["diameter"])
        else:
            spec = pipe_spec_getter(target_std, target_dn) if pipe_spec_getter else None
            if spec:
                target_dia = spec["inner_d_mm"]
            else:
                target_dia = float(target_dn)

        return {
            "type": str(target_std),
            "nominal_mm": int(target_dn),
            "diameter": float(target_dia),
            "C": float(target_c),
            "roughness_mm": float(target_rough),
        }

    # ------------------------------------------------------------------
    # 조회
    # ------------------------------------------------------------------

    def get_pipe_props(self, start: str, end: str) -> tuple[str, float]:
        pid = self.get_pipe_id_by_nodes(start, end)
        p = self.graph.get_pipe(pid) if pid else None
        if p is not None:
            return p.type, p.diameter
        return "", 0.0

    def get_pipe_length(self, start: str, end: str) -> float | None:
        p_start = self._get_node_coord(start)
        p_end = self._get_node_coord(end)
        if p_start is None or p_end is None:
            return None
        return math.dist(p_start, p_end)

    def compute_length(self, n1: str, n2: str) -> float:
        p1 = self._get_node_coord(n1)
        p2 = self._get_node_coord(n2)
        if p1 is None or p2 is None:
            return 0.0
        return dist_3d(p1, p2)

    def update_all_pipe_lengths(self) -> None:
        for p in list(self.graph.pipes.values()):
            if self._has_node(p.start) and self._has_node(p.end):
                length = self.compute_length(p.start, p.end)
                self.graph.update_pipe(p.id, length_m=length)
                self._refresh_pipe_indices(p.id)

    def _find_pipe_index(self, start: str, end: str) -> int | None:
        key = self._pipe_key(start, end)
        for i, p in enumerate(self.graph.pipes.values()):
            if self._pipe_key(p.start, p.end) == key:
                return i
        return None

    def rebuild_pipe_data_from_pipes(self) -> None:
        """호환성 유지용 빈 함수."""
        pass

    # ------------------------------------------------------------------
    # 검증 / 충돌 감지
    # ------------------------------------------------------------------

    def validate_new_segment(
        self, start: str, end: str, ignore_pipes: list | None = None
    ) -> tuple[bool, str]:
        p1 = self._get_node_coord(start)
        p2 = self._get_node_coord(end)
        if p1 is None or p2 is None:
            return False, "노드를 찾을 수 없습니다."

        if is_same_point(p1, p2):
            return False, "배관 길이가 0입니다."

        axis = self._infer_pipe_axis(start, end)
        if axis:
            if self._check_collision(start, end, axis, ignore_pipes):
                return False, "기존 배관과 겹치거나 교차합니다."
        elif self._is_xy_diagonal(start, end):
            if self._check_diagonal_overlap(start, end, ignore_pipes):
                return False, "기존 배관과 겹치거나 교차합니다."
        else:
            return False, "배관이 X, Y, Z 축과 평행하지 않습니다."

        return True, "OK"

    def _infer_pipe_axis(
        self,
        start: str,
        end: str,
        temp_nodes: dict[str, tuple[float, float, float]] | None = None,
    ) -> str | None:
        p1 = self._get_node_coord(start, temp_nodes=temp_nodes)
        p2 = self._get_node_coord(end, temp_nodes=temp_nodes)
        if p1 is None or p2 is None:
            return None

        dx = abs(p2[0] - p1[0])
        dy = abs(p2[1] - p1[1])
        dz = abs(p2[2] - p1[2])

        tol = TOLERANCE
        non_zero = []
        if dx > tol:
            non_zero.append("X")
        if dy > tol:
            non_zero.append("Y")
        if dz > tol:
            non_zero.append("Z")

        if len(non_zero) != 1:
            return None
        return non_zero[0]

    def _is_xy_diagonal(
        self,
        start: str,
        end: str,
        temp_nodes: dict[str, tuple[float, float, float]] | None = None,
    ) -> bool:
        """수평(XY) 정확 45° 대각 세그먼트 판정.

        dz=0이고 |dx|=|dy|(부동소수 허용오차 수준)일 때만 True.
        수직 45°·임의각은 대상이 아니다.
        """
        p1 = self._get_node_coord(start, temp_nodes=temp_nodes)
        p2 = self._get_node_coord(end, temp_nodes=temp_nodes)
        if p1 is None or p2 is None:
            return False

        dx = p2[0] - p1[0]
        dy = p2[1] - p1[1]
        dz = p2[2] - p1[2]
        return (
            abs(dz) <= TOLERANCE
            and abs(dx) > TOLERANCE
            and abs(abs(dx) - abs(dy)) <= TOLERANCE
        )

    def _check_diagonal_overlap(
        self,
        start: str,
        end: str,
        ignore_pipes: list | None = None,
    ) -> bool:
        """대각(XY 45°) 세그먼트의 평행 겹침 검사.

        같은 대각 직선(방향 계열·z·직선 오프셋 동일) 위에서 구간이 겹치면 True.
        직교↔대각 교차는 검사하지 않는다(첫판 명시 한계).
        """
        p1 = self._get_node_coord(start)
        p2 = self._get_node_coord(end)
        if p1 is None or p2 is None:
            return True

        dx = p2[0] - p1[0]
        dy = p2[1] - p1[1]
        # 계열: +x+y/−x−y(D1)는 x−y가 상수, +x−y/−x+y(D2)는 x+y가 상수
        same_sign = (dx > 0) == (dy > 0)
        line_c = (p1[0] - p1[1]) if same_sign else (p1[0] + p1[1])
        z_ref = p1[2]
        seg_min, seg_max = sorted((p1[0], p2[0]))

        # 색인이 있으면 후보만 — 판정식은 아래 그대로다(색인은 superset 필터).
        if self._geom_index is not None:
            pool = (self.graph.pipes[pid]
                    for pid in self._geom_candidates("D", same_sign,
                                                     z_ref, line_c)
                    if pid in self.graph.pipes)
        else:
            pool = self.graph.pipes.values()

        for p in pool:
            s_old, e_old = p.start, p.end

            if ignore_pipes:
                k_curr = self._pipe_key(s_old, e_old)
                is_ignored = False
                for ig in (ignore_pipes if isinstance(ignore_pipes, list) else [ignore_pipes]):
                    if k_curr == self._pipe_key(ig[0], ig[1]):
                        is_ignored = True
                        break
                if is_ignored:
                    continue

            if not self._is_xy_diagonal(s_old, e_old):
                continue
            po1 = self._get_node_coord(s_old)
            po2 = self._get_node_coord(e_old)
            if po1 is None or po2 is None:
                continue

            odx = po2[0] - po1[0]
            ody = po2[1] - po1[1]
            if ((odx > 0) == (ody > 0)) != same_sign:
                continue
            if not is_close(po1[2], z_ref):
                continue
            o_c = (po1[0] - po1[1]) if same_sign else (po1[0] + po1[1])
            if not is_close(o_c, line_c):
                continue

            old_min, old_max = sorted((po1[0], po2[0]))
            # 같은 직선 위이므로 x 투영 구간 교집합으로 겹침을 판정
            if min(seg_max, old_max) - max(seg_min, old_min) > TOLERANCE:
                return True

        return False

    def _check_collision(
        self,
        start: str,
        end: str,
        axis: str,
        ignore_pipes: list | None = None,
    ) -> bool:
        p1 = self._get_node_coord(start)
        p2 = self._get_node_coord(end)
        if p1 is None or p2 is None:
            return True

        if axis == "X":
            seg_min, seg_max = sorted((p1[0], p2[0]))
            fixed_1, fixed_2 = p1[1], p1[2]
        elif axis == "Y":
            seg_min, seg_max = sorted((p1[1], p2[1]))
            fixed_1, fixed_2 = p1[0], p1[2]
        else:
            seg_min, seg_max = sorted((p1[2], p2[2]))
            fixed_1, fixed_2 = p1[0], p1[1]

        # 색인이 있으면 후보만 — 판정식은 아래 그대로다(색인은 superset 필터).
        if self._geom_index is not None:
            pool = (self.graph.pipes[pid]
                    for pid in self._geom_candidates("A", axis,
                                                     fixed_1, fixed_2)
                    if pid in self.graph.pipes)
        else:
            pool = self.graph.pipes.values()

        for p in pool:
            s_old, e_old = p.start, p.end

            if ignore_pipes:
                k_curr = self._pipe_key(s_old, e_old)
                is_ignored = False
                for ig in (ignore_pipes if isinstance(ignore_pipes, list) else [ignore_pipes]):
                    if k_curr == self._pipe_key(ig[0], ig[1]):
                        is_ignored = True
                        break
                if is_ignored:
                    continue

            axis_old = self._infer_pipe_axis(s_old, e_old)
            if axis_old == axis:
                po1 = self._get_node_coord(s_old)
                po2 = self._get_node_coord(e_old)
                if po1 is None or po2 is None:
                    continue

                if axis == "X":
                    if not (is_close(fixed_1, po1[1]) and is_close(fixed_2, po1[2])):
                        continue
                    old_min, old_max = sorted((po1[0], po2[0]))
                elif axis == "Y":
                    if not (is_close(fixed_1, po1[0]) and is_close(fixed_2, po1[2])):
                        continue
                    old_min, old_max = sorted((po1[1], po2[1]))
                else:
                    if not (is_close(fixed_1, po1[0]) and is_close(fixed_2, po1[1])):
                        continue
                    old_min, old_max = sorted((po1[2], po2[2]))

                overlap_start = max(seg_min, old_min)
                overlap_end = min(seg_max, old_max)

                if overlap_end - overlap_start > TOLERANCE:
                    return True
        return False

    def _has_axis_overlap(
        self,
        start: str,
        end: str,
        ignore_pipe: Any = None,
        temp_nodes: dict[str, tuple[float, float, float]] | None = None,
    ) -> bool:
        p_start = self._get_node_coord(start, temp_nodes=temp_nodes)
        p_end = self._get_node_coord(end, temp_nodes=temp_nodes)
        if p_start is None or p_end is None:
            return False

        axis_new = self._infer_pipe_axis(start, end, temp_nodes=temp_nodes)
        if axis_new is None:
            return False

        x1, y1, z1 = p_start
        x2, y2, z2 = p_end

        if axis_new == "X":
            seg_min, seg_max = sorted((x1, x2))
            fixed_1, fixed_2 = y1, z1
        elif axis_new == "Y":
            seg_min, seg_max = sorted((y1, y2))
            fixed_1, fixed_2 = x1, z1
        else:
            seg_min, seg_max = sorted((z1, z2))
            fixed_1, fixed_2 = x1, y1

        for p in self.graph.pipes.values():
            s_old, e_old = p.start, p.end

            if ignore_pipe:
                k_curr = self._pipe_key(s_old, e_old)
                is_ignored = False
                to_check = ignore_pipe if isinstance(ignore_pipe, list) else [ignore_pipe]
                for ig in to_check:
                    if k_curr == self._pipe_key(ig[0], ig[1]):
                        is_ignored = True
                        break
                if is_ignored:
                    continue

            axis_old = self._infer_pipe_axis(s_old, e_old)
            if axis_old != axis_new:
                continue

            po1 = self._get_node_coord(s_old)
            po2 = self._get_node_coord(e_old)
            if po1 is None or po2 is None:
                continue

            if axis_new == "X":
                if not (is_close(fixed_1, po1[1]) and is_close(fixed_2, po1[2])):
                    continue
                old_min, old_max = sorted((po1[0], po2[0]))
            elif axis_new == "Y":
                if not (is_close(fixed_1, po1[0]) and is_close(fixed_2, po1[2])):
                    continue
                old_min, old_max = sorted((po1[1], po2[1]))
            else:
                if not (is_close(fixed_1, po1[0]) and is_close(fixed_2, po1[1])):
                    continue
                old_min, old_max = sorted((po1[2], po2[2]))

            overlap_start = max(seg_min, old_min)
            overlap_end = min(seg_max, old_max)

            if overlap_end - overlap_start > TOLERANCE:
                return True

        return False

    def _get_node_coord(
        self,
        node_name: str,
        *,
        temp_nodes: dict[str, tuple[float, float, float]] | None = None,
    ) -> tuple[float, float, float] | None:
        if isinstance(temp_nodes, dict) and node_name in temp_nodes:
            return temp_nodes[node_name]
        node = self.graph.get_node(node_name)
        if node is not None:
            x, y, z = node.coords
            return (float(x), float(y), float(z))
        return None

    def _has_node(self, node_name: str) -> bool:
        return self._get_node_coord(node_name) is not None

    # ------------------------------------------------------------------
    # 교차 감지 / 분할
    # ------------------------------------------------------------------

    def _find_segment_intersections(
        self,
        p1: tuple[float, float, float],
        p2: tuple[float, float, float],
        axis_new: str,
    ) -> list[dict]:
        x1, y1, z1 = p1
        x2, y2, z2 = p2
        eps = self.intersection_tol
        intersections: list[dict] = []
        if abs(x1 - x2) < eps and abs(y1 - y2) < eps and abs(z1 - z2) < eps:
            return intersections

        if axis_new == "X":
            u1, u2 = x1, x2
        elif axis_new == "Y":
            u1, u2 = y1, y2
        elif axis_new == "Z":
            u1, u2 = z1, z2
        else:
            return intersections
        u_min, u_max = (u1, u2) if u1 <= u2 else (u2, u1)

        for p in self.graph.pipes.values():
            s, e = p.start, p.end
            axis_old = self._infer_pipe_axis(s, e)
            if axis_old is None or axis_old == axis_new:
                continue
            p_start = self._get_node_coord(s)
            p_end = self._get_node_coord(e)
            if p_start is None or p_end is None:
                continue
            xa, ya, za = p_start
            xb, yb, zb = p_end
            ok = False
            xi = yi = zi = None

            if axis_new == "X" and axis_old == "Y":
                if abs(z1 - za) > eps:
                    continue
                xi, yi, zi = xa, y1, z1
                if not (u_min - eps <= xi <= u_max + eps):
                    continue
                y_min, y_max = (ya, yb) if ya <= yb else (yb, ya)
                if not (y_min - eps <= yi <= y_max + eps):
                    continue
                ok = True
            elif axis_new == "X" and axis_old == "Z":
                if abs(y1 - ya) > eps:
                    continue
                xi, yi, zi = xa, y1, z1
                if not (u_min - eps <= xi <= u_max + eps):
                    continue
                z_min, z_max = (za, zb) if za <= zb else (zb, za)
                if not (z_min - eps <= zi <= z_max + eps):
                    continue
                ok = True
            elif axis_new == "Y" and axis_old == "X":
                if abs(z1 - za) > eps:
                    continue
                xi, yi, zi = x1, ya, z1
                if not (u_min - eps <= yi <= u_max + eps):
                    continue
                x_min, x_max = (xa, xb) if xa <= xb else (xb, xa)
                if not (x_min - eps <= xi <= x_max + eps):
                    continue
                ok = True
            elif axis_new == "Y" and axis_old == "Z":
                if abs(x1 - xa) > eps:
                    continue
                xi, yi, zi = x1, ya, z1
                if not (u_min - eps <= yi <= u_max + eps):
                    continue
                z_min, z_max = (za, zb) if za <= zb else (zb, za)
                if not (z_min - eps <= zi <= z_max + eps):
                    continue
                ok = True
            elif axis_new == "Z" and axis_old == "X":
                if abs(y1 - ya) > eps:
                    continue
                xi, yi, zi = x1, y1, za
                if not (u_min - eps <= zi <= u_max + eps):
                    continue
                x_min, x_max = (xa, xb) if xa <= xb else (xb, xa)
                if not (x_min - eps <= xi <= x_max + eps):
                    continue
                ok = True
            elif axis_new == "Z" and axis_old == "Y":
                if abs(x1 - xa) > eps:
                    continue
                xi, yi, zi = x1, y1, za
                if not (u_min - eps <= zi <= u_max + eps):
                    continue
                y_min, y_max = (ya, yb) if ya <= yb else (yb, ya)
                if not (y_min - eps <= yi <= y_max + eps):
                    continue
                ok = True

            if not ok:
                continue

            if axis_new == "X":
                denom = x2 - x1
            elif axis_new == "Y":
                denom = y2 - y1
            else:
                denom = z2 - z1
            if abs(denom) < eps:
                continue

            if axis_new == "X":
                t_new = (xi - x1) / denom
            elif axis_new == "Y":
                t_new = (yi - y1) / denom
            else:
                t_new = (zi - z1) / denom

            if not (eps < t_new < 1.0 - eps):
                continue
            L_old = math.dist((xa, ya, za), (xb, yb, zb))
            if L_old < eps:
                continue

            if axis_old == "X":
                denom_old = xb - xa
                t_old = (xi - xa) / denom_old if abs(denom_old) > eps else 0.0
            elif axis_old == "Y":
                denom_old = yb - ya
                t_old = (yi - ya) / denom_old if abs(denom_old) > eps else 0.0
            else:
                denom_old = zb - za
                t_old = (zi - za) / denom_old if abs(denom_old) > eps else 0.0

            if not (eps < t_old < 1.0 - eps):
                continue
            intersections.append({
                "t_new": t_new,
                "start": s,
                "end": e,
                "point": (xi, yi, zi),
                "dist_old": L_old * t_old,
            })

        intersections.sort(key=lambda info: info["t_new"])
        return intersections

    def split_pipe(self, start: str, end: str, dist_from_start: float, new_name: str) -> None:
        old_pid = self.get_pipe_id_by_nodes(start, end)
        if not old_pid:
            _log.warning("[오류] %s - %s 배관을 찾을 수 없습니다.", start, end)
            return

        p_start = self._get_node_coord(start)
        p_end = self._get_node_coord(end)
        if p_start is None or p_end is None:
            return

        x1, y1, z1 = p_start
        x2, y2, z2 = p_end
        L = math.dist((x1, y1, z1), (x2, y2, z2))
        d = float(dist_from_start)

        if not (0.0 < d < L):
            _log.warning("[오류] 분할 거리 d=%s 는 (0, %.3f) 사이여야 합니다.", d, L)
            return

        ratio = d / L
        nx = x1 + (x2 - x1) * ratio
        ny = y1 + (y2 - y1) * ratio
        nz = z1 + (z2 - z1) * ratio
        nx, ny, nz = snap_point((nx, ny, nz), self.grid_step)

        src = self.graph.get_pipe(old_pid)
        if src is None:
            return
        inheritance_attrs = {
            "type": src.type,
            "nominal_mm": src.nominal_mm,
            "diameter": src.diameter,
            "C": src.C,
            "roughness_mm": getattr(src, "roughness_mm", 0.045),
            "equivalent_length": 0.0,
        }

        self.node_mgr.add_node(new_name, nx, ny, nz, "기본")

        self.delete_pipe(start, end)
        self.create_pipe(start, new_name, attributes=inheritance_attrs)
        self.create_pipe(new_name, end, attributes=inheritance_attrs)

        _log.debug("[분할] %s-%s -> %s (속성 유지됨)", start, end, new_name)

    def _split_new_pipe_on_existing_nodes(self, start: str, end: str) -> None:
        p_start = self._get_node_coord(start)
        p_end = self._get_node_coord(end)
        if p_start is None or p_end is None:
            return
        axis = self._infer_pipe_axis(start, end)
        if axis is None:
            return
        x1, y1, z1 = p_start
        x2, y2, z2 = p_end
        if axis == "X":
            v1, v2 = x1, x2
        elif axis == "Y":
            v1, v2 = y1, y2
        else:
            v1, v2 = z1, z2
        if abs(v2 - v1) < self.intersection_tol:
            return
        v_min, v_max = (v1, v2) if v1 <= v2 else (v2, v1)
        candidates = []
        for name, node in self.graph.nodes.items():
            xn, yn, zn = node.coords
            if name in (start, end):
                continue
            if axis == "X":
                if abs(yn - y1) > self.intersection_tol or abs(zn - z1) > self.intersection_tol:
                    continue
                vn = xn
            elif axis == "Y":
                if abs(xn - x1) > self.intersection_tol or abs(zn - z1) > self.intersection_tol:
                    continue
                vn = yn
            else:
                if abs(xn - x1) > self.intersection_tol or abs(yn - y1) > self.intersection_tol:
                    continue
                vn = zn
            if not (v_min < vn < v_max):
                continue
            t = (vn - v1) / (v2 - v1)
            if not (0.0 < t < 1.0):
                continue
            candidates.append((t, name))
        if not candidates:
            return
        candidates.sort(key=lambda item: item[0])
        idx = self._find_pipe_index(start, end)
        if idx is None:
            return

        old_pid = self.get_pipe_id_by_nodes(start, end)
        backup_attrs: dict = {}
        src = self.graph.get_pipe(old_pid) if old_pid else None
        if src is not None:
            backup_attrs = {
                "type": src.type,
                "diameter": src.diameter,
                "nominal_mm": src.nominal_mm,
                "C": src.C,
            }

        self.delete_pipe(start, end)

        prev = start
        for _, n in candidates:
            self.add_pipe(prev, n)
            if backup_attrs:
                pid = self.get_pipe_id_by_nodes(prev, n)
                if pid:
                    self.update_pipe_data(pid, **backup_attrs)
            prev = n
        self.add_pipe(prev, end)
        if backup_attrs:
            pid = self.get_pipe_id_by_nodes(prev, end)
            if pid:
                self.update_pipe_data(pid, **backup_attrs)

    def _split_existing_pipes_at_node(self, node_name: str) -> None:
        p_node = self._get_node_coord(node_name)
        if p_node is None:
            return
        xn, yn, zn = p_node
        eps = self.intersection_tol
        for p in list(self.graph.pipes.values()):
            s, e = p.start, p.end
            if node_name in (s, e):
                continue
            axis = self._infer_pipe_axis(s, e)
            if axis is None:
                continue
            p1 = self._get_node_coord(s)
            p2 = self._get_node_coord(e)
            if p1 is None or p2 is None:
                continue
            x1, y1, z1 = p1
            x2, y2, z2 = p2
            if axis == "X":
                if abs(yn - y1) > eps or abs(zn - z1) > eps:
                    continue
                v1, v2, vn = x1, x2, xn
            elif axis == "Y":
                if abs(xn - x1) > eps or abs(zn - z1) > eps:
                    continue
                v1, v2, vn = y1, y2, yn
            else:
                if abs(xn - x1) > eps or abs(yn - y1) > eps:
                    continue
                v1, v2, vn = z1, z2, zn
            v_min, v_max = (v1, v2) if v1 <= v2 else (v2, v1)
            if not (v_min + eps < vn < v_max - eps):
                continue

            backup_attrs = {
                "type": p.type,
                "diameter": p.diameter,
                "nominal_mm": p.nominal_mm,
                "C": p.C,
            }

            self.delete_pipe(s, e)
            self.add_pipe(s, node_name)
            self.add_pipe(node_name, e)

            if backup_attrs:
                p1 = self.get_pipe_id_by_nodes(s, node_name)
                if p1:
                    self.update_pipe_data(p1, **backup_attrs)
                p2 = self.get_pipe_id_by_nodes(node_name, e)
                if p2:
                    self.update_pipe_data(p2, **backup_attrs)

    # ------------------------------------------------------------------
    # 루프 / 부속 감지
    # ------------------------------------------------------------------

    def is_pipe_in_loop(self, start_node: str, end_node: str) -> bool:
        if not self._has_node(start_node) or not self._has_node(end_node):
            return False

        visited: set[str] = {start_node}
        queue = [start_node]

        while queue:
            current = queue.pop(0)
            if current == end_node:
                return True

            for p in self.graph.pipes.values():
                if (p.start == start_node and p.end == end_node) or (
                    p.start == end_node and p.end == start_node
                ):
                    continue

                neighbor = None
                if p.start == current:
                    neighbor = p.end
                elif p.end == current:
                    neighbor = p.start

                if neighbor and neighbor not in visited:
                    visited.add(neighbor)
                    queue.append(neighbor)

        return False

    def _detect_fitting_at_node(self, node_name: str) -> tuple | None:
        if not self._has_node(node_name):
            return None

        connected_pipes = [
            p for p in self.graph.pipes.values()
            if p.start == node_name or p.end == node_name
        ]
        degree = len(connected_pipes)
        if degree < 2:
            return None

        vectors = []
        for p in connected_pipes:
            other = p.end if p.start == node_name else p.start
            p1 = self._get_node_coord(node_name)
            p2 = self._get_node_coord(other)
            if p1 is None or p2 is None:
                continue
            vec = get_vector(p1, p2)
            vectors.append((p, vec))

        if degree == 2:
            p_A, v_A = vectors[0]
            p_B, v_B = vectors[1]
            angle = angle_between_vectors(v_A, v_B)

            if abs(angle - 180.0) < 5.0:
                return "Straight", 180.0, [p_A, p_B]
            if abs(angle - 90.0) < 5.0:
                return "Elbow", 90.0, [p_A, p_B]
            if abs(angle - 135.0) < 5.0:
                return "Elbow45", 45.0, [p_A, p_B]
            return "Unknown_Bend", angle, [p_A, p_B]

        elif degree == 3:
            pairs = [(0, 1), (0, 2), (1, 2)]
            run_pair = None
            branch_idx = None
            for i, j in pairs:
                v_i = vectors[i][1]
                v_j = vectors[j][1]
                ang = angle_between_vectors(v_i, v_j)
                if abs(ang - 180.0) < 5.0:
                    run_pair = (i, j)
                    branch_idx = ({0, 1, 2} - {i, j}).pop()
                    break

            if run_pair:
                return "Tee", 90.0, {
                    "run": [vectors[run_pair[0]][0], vectors[run_pair[1]][0]],
                    "branch": vectors[branch_idx][0],
                }
            else:
                return "Wye", 0.0, [p[0] for p in vectors]

        return "Cross", 90.0, [p[0] for p in vectors]

    # ------------------------------------------------------------------
    # 복합 작업 (가지관, 자동라우팅)
    # ------------------------------------------------------------------

    def add_branch_from_node_axis(
        self,
        node: str,
        axis: str,
        length: float,
        name: str,
        node_type: str = "기본",
        pipe_attrs: dict | None = None,
    ) -> tuple[bool, dict]:
        p_base = self._get_node_coord(node)
        if p_base is None:
            return False, {"msg": "시작 노드가 없습니다."}

        length = snap_value(float(length), self.grid_step)
        if abs(length) < self.grid_step:
            return False, {"msg": "길이가 너무 짧습니다."}

        x1, y1, z1 = p_base

        # "-x", "+x", "x", "X" 등 부호 포함 형식을 모두 지원
        axis_str = axis.strip()
        sign = -1 if axis_str.startswith("-") else 1
        base_axis = axis_str.lstrip("+-").upper()
        if base_axis == "X":
            x2, y2, z2 = x1 + sign * length, y1, z1
        elif base_axis == "Y":
            x2, y2, z2 = x1, y1 + sign * length, z1
        elif base_axis == "Z":
            x2, y2, z2 = x1, y1, z1 + sign * length
        else:
            return False, {"msg": "축 정보가 잘못되었습니다."}

        x2, y2, z2 = snap_point((x2, y2, z2), self.grid_step)

        intersections = self._find_segment_intersections((x1, y1, z1), (x2, y2, z2), base_axis)

        target_node = None
        is_shortened = False

        valid_hits = [hit for hit in intersections if hit["t_new"] > 0.001]

        if valid_hits:
            closest_hit = valid_hits[0]
            hx, hy, hz = closest_hit["point"]
            hit_pipe_start = closest_hit["start"]
            hit_pipe_end = closest_hit["end"]
            dist_on_old_pipe = closest_hit["dist_old"]

            existing_node_at_hit = self.node_mgr.find_node_at_position(hx, hy, hz)

            if existing_node_at_hit:
                target_node = existing_node_at_hit
            else:
                split_node_name = self.node_mgr.generate_node_name("N")
                self.split_pipe(hit_pipe_start, hit_pipe_end, dist_on_old_pipe, split_node_name)
                target_node = split_node_name

            is_shortened = True
        else:
            existing = self.node_mgr.find_node_at_position(x2, y2, z2)
            if existing:
                target_node = existing
                if self._has_axis_overlap(node, target_node):
                    return False, {"msg": "해당 경로에 이미 배관이 존재합니다."}
            else:
                temp_name = "_TEMP_CHECK_"
                temp_nodes = {temp_name: (x2, y2, z2)}
                if self._has_axis_overlap(node, temp_name, temp_nodes=temp_nodes):
                    return False, {"msg": "해당 경로에 이미 배관이 존재합니다."}

                self.node_mgr.add_node(name, x2, y2, z2, node_type=node_type)
                self._split_existing_pipes_at_node(name)
                target_node = name

        success, create_msg = self.create_pipe(node, target_node, attributes=pipe_attrs)

        result_info: dict = {"msg": create_msg, "shortened": False}

        if success:
            if is_shortened:
                result_info["shortened"] = True
                result_info["real_target"] = target_node
                result_info["msg"] = (
                    f"경로상에 배관이 감지되었습니다.\n"
                    f"교차점({target_node})까지만 연결하고 멈췄습니다."
                )
            self._split_new_pipe_on_existing_nodes(node, target_node)
            return True, result_info
        else:
            if not is_shortened and target_node == name and self._has_node(name):
                self.node_mgr.delete_node(name, self.graph.pipes)
            return False, result_info

    # 수평(XY) 대각 4방향: direction 문자열 → (x부호, y부호)
    _DIAGONAL_DIRECTIONS: dict[str, tuple[float, float]] = {
        "x+y": (1.0, 1.0),
        "x-y": (1.0, -1.0),
        "-x+y": (-1.0, 1.0),
        "-x-y": (-1.0, -1.0),
    }

    def add_branch_from_node_diagonal(
        self,
        node: str,
        direction: str,
        run: float,
        name: str,
        node_type: str = "기본",
        pipe_attrs: dict | None = None,
    ) -> tuple[bool, dict]:
        """수평(XY) 45° 대각 가지관 생성 — 전용 경로.

        직교 add_branch_from_node_axis와 달리 경로상 교차 탐지·단축·자동분할을
        하지 않는다(첫판 명시 한계). 끝점이 기존 노드와 일치하면 접속한다.
        run = 축 투영거리(가로=세로 동일), 실길이는 ×√2.
        """
        p_base = self._get_node_coord(node)
        if p_base is None:
            return False, {"msg": "시작 노드가 없습니다."}

        sign_pair = self._DIAGONAL_DIRECTIONS.get(str(direction).strip())
        if sign_pair is None:
            return False, {"msg": "축 정보가 잘못되었습니다."}

        run = snap_value(float(run), self.grid_step)
        if abs(run) < self.grid_step:
            return False, {"msg": "길이가 너무 짧습니다."}

        x1, y1, z1 = p_base
        sx, sy = sign_pair
        # 끝점은 성분별 snap_point를 하지 않는다 — 베이스가 격자 밖일 때
        # 성분별 스냅이 |dx|=|dy|를 깨면 관문에서 거절되기 때문.
        # run 스냅만으로 격자 유지 목적은 달성된다(베이스가 격자 위면 끝점도 격자 위).
        x2, y2, z2 = x1 + sx * run, y1 + sy * run, z1

        existing = self.node_mgr.find_node_at_position(x2, y2, z2)
        created_node = False
        if existing:
            target_node = existing
        else:
            self.node_mgr.add_node(name, x2, y2, z2, node_type=node_type)
            target_node = name
            created_node = True

        success, create_msg = self.create_pipe(node, target_node, attributes=pipe_attrs)

        if not success and created_node and self._has_node(name):
            self.node_mgr.delete_node(name, self.graph.pipes)

        return success, {"msg": create_msg, "shortened": False}

    def auto_route_v1(self, start: str, end: str) -> bool:
        p_start = self._get_node_coord(start)
        p_end = self._get_node_coord(end)
        if p_start is None or p_end is None:
            return False
        if start == end:
            return False
        x1, y1, z1 = p_start
        x2, y2, z2 = p_end
        axis_orders = [
            ("X", "Y", "Z"), ("X", "Z", "Y"), ("Y", "X", "Z"),
            ("Y", "Z", "X"), ("Z", "X", "Y"), ("Z", "Y", "X"),
        ]

        def set_coord(pt, ax, value):
            x, y, z = pt
            if ax == "X":
                return (value, y, z)
            elif ax == "Y":
                return (x, value, z)
            else:
                return (x, y, value)

        tol = self.intersection_tol
        for order in axis_orders:
            try_points = []
            p = (x1, y1, z1)
            try_points.append(p)
            for ax in order:
                if ax == "X":
                    p = set_coord(p, "X", x2)
                elif ax == "Y":
                    p = set_coord(p, "Y", y2)
                else:
                    p = set_coord(p, "Z", z2)
                p = snap_point(p, self.grid_step)
                try_points.append(p)
            compact = [try_points[0]]
            for pt in try_points[1:]:
                if any(abs(a - b) > tol for a, b in zip(pt, compact[-1])):
                    compact.append(pt)

            invalid = False
            for pA, pB in zip(compact, compact[1:]):
                dx, dy, dz = pB[0] - pA[0], pB[1] - pA[1], pB[2] - pA[2]
                axes = []
                if abs(dx) > tol:
                    axes.append("X")
                if abs(dy) > tol:
                    axes.append("Y")
                if abs(dz) > tol:
                    axes.append("Z")
                if len(axes) != 1:
                    continue
                inters = self._find_segment_intersections(pA, pB, axes[0])
                if inters:
                    invalid = True
                    break
            if invalid:
                continue

            nodes_temp = []
            for i, (x, y, z) in enumerate(compact):
                if i == 0:
                    nodes_temp.append(start)
                    continue
                if i == len(compact) - 1:
                    nodes_temp.append(end)
                    continue
                existing = self.node_mgr.find_node_at_position(x, y, z)
                if existing:
                    nodes_temp.append(existing)
                else:
                    nodes_temp.append(f"_AUTO_TMP_{i}")

            overlap_found = False
            for a, b in zip(nodes_temp, nodes_temp[1:]):
                temp_nodes: dict[str, tuple[float, float, float]] = {}
                if isinstance(a, str) and a.startswith("_AUTO_TMP_"):
                    temp_nodes[a] = compact[nodes_temp.index(a)]
                if isinstance(b, str) and b.startswith("_AUTO_TMP_"):
                    temp_nodes[b] = compact[nodes_temp.index(b)]
                if self._has_axis_overlap(a, b, temp_nodes=temp_nodes):
                    overlap_found = True
                if overlap_found:
                    break
            if overlap_found:
                continue

            real_nodes = []
            for i, (x, y, z) in enumerate(compact):
                if i == 0:
                    real_nodes.append(start)
                    continue
                if i == len(compact) - 1:
                    real_nodes.append(end)
                    continue
                existing = self.node_mgr.find_node_at_position(x, y, z)
                if existing:
                    real_nodes.append(existing)
                else:
                    new_name = self.node_mgr.generate_node_name("N")
                    self.node_mgr.add_node(new_name, x, y, z)
                    self._split_existing_pipes_at_node(new_name)
                    real_nodes.append(new_name)
            for a, b in zip(real_nodes, real_nodes[1:]):
                self.add_pipe(a, b)
                self._split_new_pipe_on_existing_nodes(a, b)
            return True
        return False
