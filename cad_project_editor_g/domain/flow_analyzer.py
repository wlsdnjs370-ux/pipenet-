# 파일명 domain/flow_analyzer.py
"""
FlowAnalyzer — 유량 방향 기반 등가길이(피팅) 분석 엔진.

PySide6 의존성 없음 — 순수 도메인 계층.
경고는 logging 모듈로 출력한다.
호환 레이어: logic_flow_analyzer.py → from domain.flow_analyzer import FlowAnalyzer
"""
import math
import logging
from typing import Any, Protocol, runtime_checkable

from domain.constants import AUTO_DETECTED_FITTING_IDS, DN_STANDARD, is_tee_run

logger = logging.getLogger(__name__)


@runtime_checkable
class FlowAnalyzerGraphLike(Protocol):
    nodes: dict[str, Any]
    pipes: dict[str, Any]

    def get_node(self, name: str) -> Any | None:
        """Return a node object by id, or None when it is absent."""


@runtime_checkable
class FlowAnalyzerEditorLike(Protocol):
    graph: FlowAnalyzerGraphLike


@runtime_checkable
class FlowAnalyzerFittingLookupLike(Protocol):
    def get_fitting_data_v3(self, fit_id: str) -> Any | None:
        """Return fitting library data for a fitting id."""

    def get_fittings_v3_flat(self) -> list[Any]:
        """Return flat fitting library entries used by name fallback."""

# === NFPA 13 §28.2.3.1.3 Equivalent Length Modifier — 기준 내경 표 ===
# 출처: ASTM A135 Schedule 40 (배관 라이브러리 기준값과 일치)
# NFPA 기준 변경 시 수정할 곳: 본 dict + _apply_fitting() 의 4.87 지수
# (단위: mm)
DN_TO_NFPA_REF_ID_MM: dict[int, float] = {
    15:  15.8,  20:  21.0,  25:  26.6,
    32:  35.1,  40:  40.9,  50:  52.5,
    65:  62.7,  80:  77.9, 100: 102.3,
    125: 128.2, 150: 154.1,
    200: 202.7,
    250: 254.5,
    300: 303.2,
    350: 333.3,
    400: 381.0,
}


def _resolve_nfpa_ref_id_mm(pipe) -> float:
    """pipe.nominal_mm 기반으로 NFPA 13 §28.2.3.1.3 기준 내경(mm)을 반환.
    nominal_mm이 0이거나 표에 없는 DN이면 0.0 반환 → 호출부가 보정 스킵 처리."""
    try:
        dn = int(round(float(pipe.nominal_mm)))
    except (TypeError, ValueError):
        return 0.0
    return DN_TO_NFPA_REF_ID_MM.get(dn, 0.0)


class FlowAnalyzer:
    """
    [Logic Engine V4.1 - Valve ID Priority Patch]
    - 밸브: "추측하지 말고, 저장된 ID(Key)를 먼저 확인하라."
    - 엘보/티: 기존 기하학적 로직 유지 (기본값 적용)
    """
    def __init__(
        self,
        editor: FlowAnalyzerEditorLike | None = None,
        fitting_lookup: FlowAnalyzerFittingLookupLike | None = None,
        graph: FlowAnalyzerGraphLike | None = None,
    ):
        if graph is None and editor is None:
            raise TypeError("FlowAnalyzer requires either editor or graph")
        if fitting_lookup is None and editor is None:
            raise TypeError("FlowAnalyzer requires either editor or fitting_lookup")

        self.editor = editor
        self.graph: Any = graph if graph is not None else editor.graph
        self.fitting_lookup: Any = fitting_lookup if fitting_lookup is not None else editor

        # M3.2: 경고 변수 — session-level (라이브러리 변경 시에만 리셋, alert fatigue 회피)
        self._session_warned_log: set = set()

    def run(self, real_flow_map):
        """
        [Step 2] 등가길이 분석 실행
        - V4.2 Update: 노드 속성(ID) 최우선 반영
        """
        # M3.2 / M4.2: 보고서용 경고 — 매 run마다 초기화 (시나리오 변경 시 박제 방지)
        self._run_warned_for_report: list = []   # report 경고 (0 처리 항목 섹션)
        self._run_critical_warnings: list = []   # fatal 경고 (결과 신뢰 불가 배너)

        # 1. 초기화
        for p in self.graph.pipes.values():
            p.equivalent_length = 0.0
            p.fittings = []

        count_tee = 0
        count_elbow = 0
        count_elbow45 = 0
        count_valve = 0

        # [A] 데이터 구조 변환 (Graph Building)
        node_graph = {n: {'in': [], 'out': []} for n in self.graph.nodes}

        for pid, flow_info in real_flow_map.items():
            if pid not in self.graph.pipes:
                continue
            p_obj = self.graph.pipes[pid]
            u_start, u_end = flow_info['start'], flow_info['end']

            if u_start in node_graph:
                node_graph[u_start]['out'].append(p_obj)
            if u_end in node_graph:
                node_graph[u_end]['in'].append(p_obj)

        # [A-2] 물리적 연결 차수(degree) — 흐름과 무관하게 노드에 붙은 실제 배관 수.
        # 캡 등으로 유량 0인 다리는 real_flow_map에서 제외되지만 graph.pipes에는 남는다.
        # 이 값으로 "ㅗ(티, 3다리↑)"와 "ㄴ(엘보, 2다리)"를 구분한다.
        phys_degree: dict = {n: 0 for n in self.graph.nodes}
        for p in self.graph.pipes.values():
            s = getattr(p, "start", None)
            e = getattr(p, "end", None)
            if s in phys_degree:
                phys_degree[s] += 1
            if e in phys_degree:
                phys_degree[e] += 1

        # [B] 노드 순회하며 형상 및 밸브 분석
        for node_name, conn in node_graph.items():
            in_pipes = conn['in']
            out_pipes = conn['out']
            degree = len(in_pipes) + len(out_pipes)

            # ---------------------------------------------------
            # [Logic 1] 밸브(Valve) 탐지 (ID 우선 참조)
            # ---------------------------------------------------
            valve_id = self._get_fitting_id_from_node(node_name)

            if valve_id:
                target_pipe = out_pipes[0] if out_pipes else (in_pipes[0] if in_pipes else None)
                if target_pipe:
                    disp_name = str(valve_id)
                    if hasattr(self.fitting_lookup, "get_fitting_data_v3"):
                        fdata_v3 = self.fitting_lookup.get_fitting_data_v3(valve_id)
                        if fdata_v3 is not None:
                            try:
                                from services.i18n_service import I18nService
                                _lang = I18nService.instance().lang
                            except Exception:
                                _lang = "ko"
                            if _lang == "ko":
                                disp_name = fdata_v3.name_ko or fdata_v3.name_en or str(valve_id)
                            else:
                                disp_name = fdata_v3.name_en or fdata_v3.name_ko or str(valve_id)

                    self._apply_fitting(target_pipe, valve_id, disp_name)
                    count_valve += 1
                else:
                    logger.warning("밸브 '%s'에 연결된 배관을 찾지 못함 (고립됨?)", node_name)

            # ---------------------------------------------------
            # [Logic 2] 티/엘보 탐지 (기하학 분석 - 기존 유지)
            # ---------------------------------------------------
            if degree < 2:
                continue

            # (1) 분기 (Split)
            if len(in_pipes) == 1 and len(out_pipes) >= 1:
                p_in = in_pipes[0]
                for p_out in out_pipes:
                    angle = self._calculate_angle(p_in, p_out, node_name)
                    # 흐름상 1 in / 1 out 이라도 노드에 물리적 배관이 3개↑(ㅗ)이면
                    # 캡·무흐름 다리가 있는 '티'다 — 꺾임을 엘보로 강등하지 않는다.
                    # 물리적으로 정확히 2다리(ㄴ)일 때만 진짜 엘보.
                    single_flow = len(out_pipes) == 1 and len(in_pipes) == 1
                    is_elbow_node = single_flow and phys_degree.get(node_name, 0) < 3
                    if is_elbow_node and self._is_45_degree(angle):
                        # 45° 엘보는 진짜 엘보(2다리) 노드에만 적용.
                        # 티(3다리↑) 분류는 현행 유지 — 45° 합류/분기 티는 별도 시드가
                        # 없어 기존과 동일하게 90° 밴드(_is_90_degree)로만 판정한다.
                        self._apply_fitting(p_out, AUTO_DETECTED_FITTING_IDS["elbow_45"], "Elbow45")
                        count_elbow45 += 1
                    elif self._is_90_degree(angle):
                        if is_elbow_node:
                            self._apply_fitting(p_out, AUTO_DETECTED_FITTING_IDS["elbow_90"], "Elbow")
                            count_elbow += 1
                        else:
                            self._apply_fitting(p_out, AUTO_DETECTED_FITTING_IDS["tee_branch"], "Tee(Branch)")
                            count_tee += 1

            # (2) 합류 (Merge)
            elif len(in_pipes) >= 2 and len(out_pipes) == 1:
                p_out = out_pipes[0]
                is_branch_flow = False
                for p_in in in_pipes:
                    angle = self._calculate_angle(p_in, p_out, node_name)
                    if self._is_90_degree(angle):
                        is_branch_flow = True
                        break

                if is_branch_flow:
                    self._apply_fitting(p_out, AUTO_DETECTED_FITTING_IDS["tee_branch"], "Tee(Merge)")
                    count_tee += 1

        # 45° 엘보 줄은 검출이 있을 때만 붙인다 — 직교 전용 도면의 로그를 불변으로 유지.
        elbow45_line = (
            f"        - 45° 엘보(Elbow45) : {count_elbow45} 개\n" if count_elbow45 else ""
        )
        return (f"        - 티(Tee) : {count_tee} 개\n"
                f"        - 엘보(Elbow) : {count_elbow} 개\n"
                + elbow45_line
                + f"        - 밸브(Valve) : {count_valve} 개\n ")

    # -----------------------------------------------------------
    # [Helper] 라이브러리 매칭 (ID 우선 로직)
    # -----------------------------------------------------------
    def _get_fitting_id_from_node(self, node_name):
        """
        [정석 매칭 로직]
        1순위: Node.fitting_id를 그대로 사용 (가장 정확)
        2순위: 없다면 이름을 보고 라이브러리에서 ID를 역추적 (호환성)

        수조·펌프는 밸브 부속 경로에 넣지 않는다.
        (예: type "WT" → name_en "FlowTurned" 부분일치로 TEE_BRANCH 오탐 방지)
        """
        from domain.node_predicates import is_pump_graph_node, is_tank_graph_node

        node_obj = self.graph.get_node(node_name)
        if node_obj is None:
            return None
        if is_pump_graph_node(node_obj) or is_tank_graph_node(node_obj):
            return None
        saved_id = getattr(node_obj, "fitting_id", None)
        if saved_id and hasattr(self.fitting_lookup, "get_fitting_data_v3"):
            if self.fitting_lookup.get_fitting_data_v3(saved_id) is not None:
                return saved_id
        node_type = str(getattr(node_obj, "type", "") or "").strip()
        return self._find_valve_id_by_name_fallback(node_type)

    def _find_valve_id_by_name_fallback(self, node_type_name):
        """[보조] 이름으로 ID 찾기 (속성창에서 선택 안 하고 텍스트만 있을 때)"""
        if not node_type_name or node_type_name == "기본":
            return None

        search_key = str(node_type_name).replace(" ", "").lower()
        candidates = self.fitting_lookup.get_fittings_v3_flat()

        # 1차: category 값 매칭 (언더스코어 제거 후 공백 제거 비교)
        for item in candidates:
            if item.category.value.replace("_", "") == search_key:
                return item.id

        # 2차: id 또는 name_en 정확 매칭
        for item in candidates:
            if item.id.lower() == search_key:
                return item.id
            if item.name_en.replace(" ", "").lower() == search_key:
                return item.id

        # 3차: name_en 부분 매칭
        for item in candidates:
            if search_key in item.name_en.replace(" ", "").lower():
                return item.id

        return None

    def _is_tee_run_fitting(self, fit_id: str) -> bool:
        """v3 라이브러리에서 카테고리 조회 → TEE_RUN 판정.

        라이브러리에 없는 ID는 False 반환 — 없는 항목을 TEE_RUN으로 취급하지 않음.
        """
        if not fit_id or not hasattr(self.fitting_lookup, "get_fitting_data_v3"):
            return False
        item = self.fitting_lookup.get_fitting_data_v3(fit_id)
        if item is None:
            return False
        return is_tee_run(item.category)

    def _apply_fitting(self, pipe, fit_id, display_name):
        # ① TEE_RUN 가드 — 직류 티는 등가길이 0, 경고 없음 (이른 return)
        if self._is_tee_run_fitting(fit_id):
            return

        d_actual_mm = float(pipe.diameter)
        d_ref_mm = _resolve_nfpa_ref_id_mm(pipe)

        # ③ DN 산출 — v3 _get_lib_value 인터페이스에 맞춰 int로 전달
        try:
            dn = int(round(float(pipe.nominal_mm)))
        except (TypeError, ValueError):
            dn = 0

        # ④ 조건 강화: d_ref_mm > 0 and dn > 0 (두 조건 모두 충족 시에만 보정)
        if d_ref_mm > 0 and dn > 0:
            # NFPA 13 §28.2.3.1.3: Sch.40/30 기준 등가길이에서 출발, 실제 구경비로 보정
            # 지수 4.87 = NFPA 13 명문값 (H-W 본식 정밀값 4.8708과 의도적으로 다름)
            le_base_m = self._get_lib_value(fit_id, dn, pipe.id)  # ② d_ref_mm → dn 전달
            le_d_adj = le_base_m * (d_actual_mm / d_ref_mm) ** 4.87
        else:
            # ⑤ nominal_mm=0(구버전) 또는 NFPA 표 외 DN → dn 전달 (dn=0이면 내부에서 0 반환)
            le_d_adj = self._get_lib_value(fit_id, dn, pipe.id)

        # NFPA 13 §28.2.3.1.3: C값 보정 — 지수 1.85 (NFPA 정석값)
        if pipe.C > 0:
            le_adj = le_d_adj * ((pipe.C / 120.0) ** 1.85)
        else:
            le_adj = le_d_adj

        if le_adj > 0:
            pipe.equivalent_length += le_adj
            pipe.fittings.append(display_name)

    def _get_lib_value(self, fit_id: str, dn: int, pipe_id: str = "") -> float:
        """fit_id와 호칭경 DN(int)으로 등가길이(m)를 반환. v3 NFPA 표 룩업.

        L/D 방식 완전 폐기. mm/m 변환 없음 (NFPA 표값이 이미 m 단위).
        반환: 등가길이 m. 0 처리 케이스는 모두 경고 후 0.0 반환.

        경고 정책 (M4.2 — 3분류):
          - dn=0 (nominal_mm 미설정)          → _run_critical_warnings + _session_warned_log (fatal)
          - dn>0, fit_id v3 미등록            → _run_warned_for_report + _session_warned_log
          - dn∉DN_STANDARD (비표준 DN)        → _run_warned_for_report + _session_warned_log
          - dn∈DN_STANDARD, 셀 null           → _run_warned_for_report + _session_warned_log
        """
        # ── 분류 ①: dn=0 → fatal (결과 신뢰 불가)
        if dn == 0:
            key = (fit_id, 0)
            if fit_id and str(fit_id).strip() and key not in self._session_warned_log:
                self._session_warned_log.add(key)
                logger.warning(
                    "부속 '%s': DN=0 — nominal_mm 미설정. 등가길이 0 처리 (결과 신뢰 불가).",
                    fit_id,
                )
            self._append_critical(pipe_id, fit_id, fit_id, "DN=0: nominal_mm 미설정")
            return 0.0

        if not hasattr(self.fitting_lookup, "get_fitting_data_v3"):
            return 0.0

        item = self.fitting_lookup.get_fitting_data_v3(fit_id)

        # ── 분류 ②: fit_id v3 미등록 → report 경고 (USER 부속 0이면 무조건 경고)
        if item is None:
            key = (fit_id, dn)
            if fit_id and str(fit_id).strip() and key not in self._session_warned_log:
                self._session_warned_log.add(key)
                logger.warning(
                    "부속 데이터 누락: '%s' — v3 라이브러리에서 찾을 수 없습니다. "
                    "등가길이 0.0 m 적용.",
                    fit_id,
                )
            self._append_report_warning(pipe_id, fit_id, dn, fit_id, "v3 라이브러리 미등록으로 0 처리됨")
            return 0.0

        # ── 분류 ③④: 셀 null (표준 DN / 비표준 DN 구분 메시지)
        cell = item.get_cell(dn)
        if cell is None or cell.value is None:
            if dn not in DN_STANDARD:
                reason = f"비표준 DN{dn} — NFPA 표 범위 외로 0 처리됨"
            else:
                reason = "NFPA 13 미정의로 0 처리됨"
            key = (fit_id, dn)
            if key not in self._session_warned_log:
                self._session_warned_log.add(key)
                logger.warning("부속 '%s' DN%d: %s", fit_id, dn, reason)
            self._append_report_warning(pipe_id, fit_id, dn, item.name_ko, reason)
            return 0.0

        return cell.value

    # -----------------------------------------------------------
    # [M4.2] 경고 수집 헬퍼 — 중복 방지 append
    # -----------------------------------------------------------

    def _append_report_warning(
        self, pipe_id: str, fit_id: str, dn: int, fitting_name: str, reason: str
    ) -> None:
        """_run_warned_for_report에 (pipe_id, fit_id, dn) 단일화 append."""
        if not any(
            w["pipe_id"] == pipe_id and w["fitting_id"] == fit_id and w["dn"] == dn
            for w in self._run_warned_for_report
        ):
            self._run_warned_for_report.append({
                "pipe_id": pipe_id,
                "fitting_id": fit_id,
                "dn": dn,
                "fitting_name": fitting_name,
                "reason": reason,
            })

    def _append_critical(
        self, pipe_id: str, fit_id: str, fitting_name: str, reason: str
    ) -> None:
        """_run_critical_warnings에 (pipe_id, fit_id) 단일화 append."""
        if not any(
            w["pipe_id"] == pipe_id and w["fitting_id"] == fit_id
            for w in self._run_critical_warnings
        ):
            self._run_critical_warnings.append({
                "pipe_id": pipe_id,
                "fitting_id": fit_id,
                "fitting_name": fitting_name,
                "reason": reason,
                "expected_dn_set": "/".join(f"DN{d}" for d in DN_STANDARD),
            })

    # -----------------------------------------------------------
    # [M3.6] 경고 수집 API
    # -----------------------------------------------------------

    def get_report_warnings(self) -> list:
        """보고서용 경고 목록 반환. _run_warned_for_report의 사본 (직접 노출 금지).

        M4 보고서 "0 처리 항목" 섹션에 사용.
        매 run() 호출 후 갱신됨 — 시나리오 변경 시 박제 방지.
        반환 형식: [{"pipe_id": str, "fitting_id": str, "dn": int,
                    "fitting_name": str, "reason": str}]
        """
        return list(self._run_warned_for_report)

    def get_critical_warnings(self) -> list:
        """fatal 경고 목록 반환. _run_critical_warnings의 사본 (직접 노출 금지).

        DN=0(nominal_mm 미설정) 케이스 — 결과 신뢰 불가 배너에 사용.
        매 run() 호출 후 갱신됨.
        반환 형식: [{"pipe_id": str, "fitting_id": str, "fitting_name": str,
                    "reason": str, "expected_dn_set": str}]
        """
        return list(self._run_critical_warnings)

    def emit_session_log(self) -> None:
        """라이브러리 버전 변경 감지 시 호출 — 세션 누적 경고 일괄 출력 후 리셋.

        M4 보고서는 get_report_warnings()만 사용. 이 메서드는 alert fatigue 방지용 API.
        외부 호출자가 라이브러리 업데이트 시점에 호출하여 _session_warned_log를 비운다.
        """
        if self._session_warned_log:
            logger.warning(
                "[FlowAnalyzer] 라이브러리 업데이트 감지 — 세션 누적 경고 %d건:",
                len(self._session_warned_log),
            )
            for key in self._session_warned_log:
                logger.warning("  경고 항목: %s", key)
            self._session_warned_log = set()

    # -----------------------------------------------------------
    # [Helper] 기하학 계산 (각도) - 기존 유지
    # -----------------------------------------------------------
    def _calculate_angle(self, p_in, p_out, node):
        p1_s = self._get_node_coord(p_in.start)
        p1_e = self._get_node_coord(p_in.end)
        if p1_s is None or p1_e is None:
            return 0
        if p_in.end == node:
            v_in = (p1_e[0] - p1_s[0], p1_e[1] - p1_s[1], p1_e[2] - p1_s[2])
        else:
            v_in = (p1_s[0] - p1_e[0], p1_s[1] - p1_e[1], p1_s[2] - p1_e[2])

        p2_s = self._get_node_coord(p_out.start)
        p2_e = self._get_node_coord(p_out.end)
        if p2_s is None or p2_e is None:
            return 0
        if p_out.start == node:
            v_out = (p2_e[0] - p2_s[0], p2_e[1] - p2_s[1], p2_e[2] - p2_s[2])
        else:
            v_out = (p2_s[0] - p2_e[0], p2_s[1] - p2_e[1], p2_s[2] - p2_e[2])

        mag1 = math.sqrt(v_in[0] ** 2 + v_in[1] ** 2 + v_in[2] ** 2)
        mag2 = math.sqrt(v_out[0] ** 2 + v_out[1] ** 2 + v_out[2] ** 2)
        if mag1 * mag2 == 0:
            return 0

        dot = v_in[0] * v_out[0] + v_in[1] * v_out[1] + v_in[2] * v_out[2]
        cos_val = max(-1.0, min(1.0, dot / (mag1 * mag2)))
        return math.degrees(math.acos(cos_val))

    def _is_90_degree(self, angle):
        return 45 < angle < 135

    def _is_45_degree(self, angle):
        # 45° 엘보 밴드 [22.5°, 67.5°) — 흐름 꺾임각(직진=0°) 기준.
        # 직교 도면의 꺾임각은 0/90/180뿐이라 이 밴드는 직교 결과에 영향이 없다.
        # 하한 22.5°는 임포트 부정확 대각(43° 등)까지 흡수, 상한 67.5°부터 90° 밴드.
        return 22.5 <= angle < 67.5

    def _get_node_coord(self, node_name):
        node_obj = self.graph.get_node(node_name)
        if node_obj is None:
            return None
        x, y, z = node_obj.coords
        return (float(x), float(y), float(z))
