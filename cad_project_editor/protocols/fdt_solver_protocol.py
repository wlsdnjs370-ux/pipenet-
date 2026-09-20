"""
FdtSteadyStateSolverProtocol — FDT 트랜짓 정상상태 망해석 솔버 교체 가능 인터페이스(창구).

물리 코어(domain.fdt.front_tracker)는 매 스텝 "지금까지 물로 찬 배관 + 전진 중인 전선 +
헤드 상태(통제/젖음)"를 이 창구에 주고, 전선별 유입유량[m³/s]과 젖은 헤드 물압을 돌려받는다.
창구 뒤가 무엇인지는 코어가 모른다:
- 가짜 솔버(상수유량) — 걷는 해골·닫힌형 검산용(테스트).
- EpanetDirectSolver — EPANET 직접해(부속충수 전선+배압경계). 구현: infra.fdt_epanet_bridge.
- (후속) 자체식 트리 솔버 — EPANET 시간 실측 후 채택 판단.

계획서 §6 4계층 규약: 물리 코어는 domain 순수, 정상상태 망해석만 이 창구 뒤 infra(EPANET).

★정식 계약(U 슬라이스 §3) = `solve(full, fronts, controlled_heads, wet_head_ids)`. 트랜짓·안정화를
하나의 정상상태 해로 다룬다. **solve()가 다루는 건 헤드의 물 방수 경계(BC)뿐**이다 — 젖은
헤드엔 K√P 물 방수구, 마른 헤드엔 물 방수구를 제거(공통 microleak). **실제 공기 배기는 solve가
아니라 `front_tracker`가 처리**한다(연결 공기영역 blowdown). 옛 `front_inflows` /
`front_inflows_with_heads`는 **호환 확장**으로 내려놓는다 — Stage 1에선 코어가 아직 이들을
호출하지만(EpanetFrontSolver가 solve()에 위임), Stage 2에서 코어가 solve()를 직접 호출하면 걷어낼
예정이다.
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Protocol, Sequence, runtime_checkable

if TYPE_CHECKING:
    from domain.fdt.front_tracker import Front


@runtime_checkable
class FdtSteadyStateSolverProtocol(Protocol):
    """FDT 충수망 정상상태 솔버가 구현하는 창구. 정식 계약 = solve()."""

    def solve(
        self,
        full_pipe_ids: frozenset[str],
        fronts: "Sequence[Front]",
        controlled_heads: dict[str, float],
        wet_head_ids: "set[str] | frozenset[str]",
    ) -> tuple[list[float], dict[str, float]]:
        """현재 충수상태 + 헤드 상태에서 (전선별 유입[m³/s], {젖은 헤드: 물압[bar,게이지]})를 낸다.

        full_pipe_ids   : 지금까지 완전 충수된 건식 배관 ID 집합(습식측은 상시 충수).
        fronts          : 전진 중인 물선단 목록(배관·젖은 끝·채운 길이·앞 공기 배압).
        controlled_heads: {이 FDT 설비 개방 헤드 node_id → 물 방수 K[lpm/bar^0.5]}. 도면 전체
                          헤드가 아니라 이 설비가 통제하는 헤드만.
        wet_head_ids    : controlled_heads 중 현재 물이 닿아 방수 중인(emitter) 헤드(⊆ controlled).

        구현체는 펌프 공급곡선·습식측·완전충수 배관으로 정상상태를 풀되, ① controlled 헤드의
        빌더 방수구(가상 emitter)를 떼고 ② 그 노드 계수를 초기화한 뒤 ③ 젖은 헤드에만 K√P
        emitter를 달고 ④ 그 외 미세 누수(microleak)를 단일 값으로 정규화한다 — 검증·장치 경로가
        헤드 BC·microleak를 동일하게 처리(분기 제거). 각 전선 위치에 배압(`Front.
        back_pressure_gauge_pa`) 경계를 둬 그 전선으로 드는 물 유량을 낸다. 마른 헤드는 물
        방수구가 아니므로 반환 물압에 없다.

        반환: (fronts와 같은 순서의 유입유량[m³/s] 리스트, {젖은 헤드 node_id: 물압[bar,게이지]}).
        """
        ...

    # ── 호환 확장(Stage 1 — 코어가 아직 호출. Stage 2에서 solve()로 이관 후 제거) ──────────
    #   front_inflows(full_pipe_ids, fronts) -> list[float]
    #     헤드 BC 없는 트랜짓 솔브 = solve(full, fronts, {}, set())[0]. 현 트랜짓 거동(헤드 배기구)
    #     유지. 가짜 솔버는 이것만 구현해도 트랜짓 도달 테스트에 충분하다.
    #
    #   front_inflows_with_heads(full_pipe_ids, fronts, head_k_si)
    #     -> (전선별 유입[m³/s], {헤드 node_id: 물압[bar,게이지]})
    #     안정화 솔브 = solve(full, fronts, head_k_si, set(head_k_si)). head_k_si = {헤드 node_id:
    #     물 방수 K}. 개방 헤드를 모두 물 방수구로 두고 풀어 헤드 물압을 돌려준다. run_transit에
    #     required_pressures를 줄 때만 호출(B 작동시간).
    #
    # ── 선택 확장(H1 가시화 데이터 출구) — 시계열을 안 만드는 솔버는 미구현 가능(수압 nan) ─────
    #   node_pressures_bar(node_ids) -> {node_id: 물압[bar,게이지]}
    #     직전 solve 해에서 감시 지점(펌프·밸브·헤드) 수압을 돌려준다. 모델에 없는 노드(물 안
    #     닿은 헤드)는 키 생략. run_transit에 monitor를 줄 때만 호출.
