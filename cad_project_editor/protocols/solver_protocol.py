"""
HydraulicSolverProtocol — 수력학 솔버 교체 가능 인터페이스

이 Protocol을 구현하면 WNTR/EPANET 외에 다른 솔버로 교체할 수 있다.
현재 구현체: infra.epanet_adapter.EpanetAdapter
"""
from __future__ import annotations

from typing import TYPE_CHECKING, Protocol, runtime_checkable

if TYPE_CHECKING:
    from domain.models import HydraulicModel


@runtime_checkable
class HydraulicSolverProtocol(Protocol):
    """수력학 솔버가 반드시 구현해야 하는 메서드 집합."""

    def run_simulation(self, editor: object, method: str = "H-W") -> "HydraulicModel":
        """순산(Forward Simulation): 현재 네트워크 상태로 수력학 해석을 수행한다."""
        ...

    def run_unified_analysis(
        self,
        editor: object,
        pump_roles: dict[str, str],
        method: str = "H-W",
        *,
        demand_target_p_bar: float | None = None,
    ) -> "HydraulicModel":
        """통합 해석(역산 포함): 펌프별 역할(DEMAND/SUPPLY)에 따라 해석한다."""
        ...

    def calculate_demand(
        self,
        editor: object,
        target_p_bar: float = 1.0,
        method: str = "H-W",
    ) -> "HydraulicModel":
        """역산(Demand Calculation): 목표 압력에 맞는 펌프 사양을 역산한다."""
        ...

    def validate_topology(self, editor: object) -> dict:
        """토폴로지 검증: 펌프/PRV/오리피스의 연결 상태를 검사한다.

        Returns:
            {"valid": True} 또는 {"valid": False, "target": str, "title": str, "msg": str}
        """
        ...
