"""Pure domain types for explicitly configured antifreeze hydraulic analysis.

This module has no editor, UI, EPANET, or persistence dependencies.  Merely
importing it cannot activate antifreeze behavior; callers must create an
``AntifreezeAnalysisCase`` explicitly.
"""
from __future__ import annotations

import math
from dataclasses import dataclass
from enum import Enum
from typing import Any


GRAVITY_M_S2 = 9.80665
PA_PER_BAR = 100_000.0
M3S_PER_LPM = 1.0 / 60_000.0
M2S_PER_CST = 1.0e-6
PA_S_PER_MPA_S = 1.0e-3

# EPANET defines VISCOSITY relative to 1.0 cSt water at 20 degC and specific
# gravity relative to water at 4 degC.  The latter density is kept explicit so
# the calculation contract is not hidden in a rounded "SG = rho / 1000" rule.
EPANET_REFERENCE_KINEMATIC_VISCOSITY_M2_S = 1.0e-6
WATER_DENSITY_AT_4C_KG_M3 = 999.972


class ViscosityInputBasis(str, Enum):
    """The one viscosity quantity the user supplied."""

    DYNAMIC_MPA_S = "DYNAMIC_MPA_S"
    KINEMATIC_CST = "KINEMATIC_CST"



class FittingLossPolicy(str, Enum):
    """Supported first-release fitting policy.

    ``MINOR_LOSS_K`` is represented for forward-compatible persistence, but
    the first implementation must not select it implicitly.
    """

    NFPA_EQUIVALENT_LENGTH = "NFPA_EQUIVALENT_LENGTH"
    MINOR_LOSS_K = "MINOR_LOSS_K"


def _finite_float(value: Any, field_name: str) -> float:
    try:
        result = float(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(field_name) from exc
    if not math.isfinite(result):
        raise ValueError(field_name)
    return result


@dataclass(frozen=True)
class ResolvedFluidProperties:
    """Canonical, immutable calculation properties for one design point."""

    density_kg_m3: float
    dynamic_viscosity_pa_s: float
    kinematic_viscosity_m2_s: float
    epanet_relative_viscosity: float
    specific_gravity: float
    meter_head_per_bar: float
    bar_per_meter_head: float
    emitter_m3s_per_sqrt_m_per_k_si: float


@dataclass(frozen=True)
class AntifreezePropertyInput:
    """Project-owned user input for one antifreeze design condition."""

    product_name: str
    antifreeze_type: str
    design_temperature_c: float
    density_kg_m3: float
    viscosity_input_basis: ViscosityInputBasis
    viscosity_value: float
    source_reference: str
    concentration_percent: float | None = None

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "viscosity_input_basis",
            ViscosityInputBasis(self.viscosity_input_basis),
        )

        if not str(self.product_name).strip():
            raise ValueError("product_name")
        if not str(self.antifreeze_type).strip():
            raise ValueError("antifreeze_type")

        temperature = _finite_float(self.design_temperature_c, "design_temperature_c")
        density = _finite_float(self.density_kg_m3, "density_kg_m3")
        viscosity = _finite_float(self.viscosity_value, "viscosity_value")
        if density <= 0.0:
            raise ValueError("density_kg_m3")
        if viscosity <= 0.0:
            raise ValueError("viscosity_value")

        concentration = self.concentration_percent
        if concentration is not None:
            concentration = _finite_float(concentration, "concentration_percent")
            if not 0.0 <= concentration <= 100.0:
                raise ValueError("concentration_percent")

        object.__setattr__(self, "design_temperature_c", temperature)
        object.__setattr__(self, "density_kg_m3", density)
        object.__setattr__(self, "viscosity_value", viscosity)
        object.__setattr__(self, "concentration_percent", concentration)
        source_reference = str(self.source_reference).strip()
        if not source_reference:
            raise ValueError("source_reference")

        object.__setattr__(self, "product_name", str(self.product_name).strip())
        object.__setattr__(self, "antifreeze_type", str(self.antifreeze_type).strip())
        object.__setattr__(self, "source_reference", source_reference)

    def resolve(self) -> ResolvedFluidProperties:
        density = self.density_kg_m3

        if self.viscosity_input_basis is ViscosityInputBasis.DYNAMIC_MPA_S:
            dynamic_pa_s = self.viscosity_value * PA_S_PER_MPA_S
            kinematic_m2_s = dynamic_pa_s / density
        else:
            kinematic_m2_s = self.viscosity_value * M2S_PER_CST
            dynamic_pa_s = kinematic_m2_s * density

        bar_per_meter = density * GRAVITY_M_S2 / PA_PER_BAR
        meter_per_bar = 1.0 / bar_per_meter

        return ResolvedFluidProperties(
            density_kg_m3=density,
            dynamic_viscosity_pa_s=dynamic_pa_s,
            kinematic_viscosity_m2_s=kinematic_m2_s,
            epanet_relative_viscosity=(
                kinematic_m2_s / EPANET_REFERENCE_KINEMATIC_VISCOSITY_M2_S
            ),
            specific_gravity=density / WATER_DENSITY_AT_4C_KG_M3,
            meter_head_per_bar=meter_per_bar,
            bar_per_meter_head=bar_per_meter,
            emitter_m3s_per_sqrt_m_per_k_si=(
                M3S_PER_LPM * math.sqrt(bar_per_meter)
            ),
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "product_name": self.product_name,
            "antifreeze_type": self.antifreeze_type,
            "concentration_percent": self.concentration_percent,
            "design_temperature_c": self.design_temperature_c,
            "density_kg_m3": self.density_kg_m3,
            "viscosity_input_basis": self.viscosity_input_basis.value,
            "viscosity_value": self.viscosity_value,
            "source_reference": self.source_reference,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AntifreezePropertyInput":
        if not isinstance(data, dict):
            raise TypeError("antifreeze properties must be a dict")
        return cls(
            product_name=data.get("product_name", ""),
            antifreeze_type=data.get("antifreeze_type", ""),
            concentration_percent=data.get("concentration_percent"),
            design_temperature_c=data.get("design_temperature_c"),
            density_kg_m3=data.get("density_kg_m3"),
            viscosity_input_basis=data.get("viscosity_input_basis"),
            viscosity_value=data.get("viscosity_value"),
            source_reference=data.get("source_reference", ""),
        )


@dataclass(frozen=True)
class AntifreezeAnalysisCase:
    """An explicitly created antifreeze analysis case.

    The default project state is the absence of this object.  ``enabled`` is
    explicit and never derived from a Darcy formula selection or property data.
    """

    case_id: str
    name: str
    boundary_node_id: str
    properties: AntifreezePropertyInput
    fitting_loss_policy: FittingLossPolicy = FittingLossPolicy.NFPA_EQUIVALENT_LENGTH
    enabled: bool = True

    def __post_init__(self) -> None:
        object.__setattr__(
            self,
            "fitting_loss_policy",
            FittingLossPolicy(self.fitting_loss_policy),
        )
        if not str(self.case_id).strip():
            raise ValueError("case_id")
        if not str(self.name).strip():
            raise ValueError("name")
        if not str(self.boundary_node_id).strip():
            raise ValueError("boundary_node_id")
        if not isinstance(self.properties, AntifreezePropertyInput):
            raise TypeError("properties")
        object.__setattr__(self, "case_id", str(self.case_id).strip())
        object.__setattr__(self, "name", str(self.name).strip())
        object.__setattr__(self, "boundary_node_id", str(self.boundary_node_id).strip())
        object.__setattr__(self, "enabled", bool(self.enabled))
    def to_dict(self) -> dict[str, Any]:
        return {
            "case_id": self.case_id,
            "name": self.name,
            "boundary_node_id": self.boundary_node_id,
            "properties": self.properties.to_dict(),
            "fitting_loss_policy": self.fitting_loss_policy.value,
            "enabled": self.enabled,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "AntifreezeAnalysisCase":
        if not isinstance(data, dict):
            raise TypeError("antifreeze analysis case must be a dict")
        return cls(
            case_id=data.get("case_id", ""),
            name=data.get("name", ""),
            boundary_node_id=data.get("boundary_node_id", ""),
            properties=AntifreezePropertyInput.from_dict(data.get("properties", {})),
            fitting_loss_policy=data.get(
                "fitting_loss_policy",
                FittingLossPolicy.NFPA_EQUIVALENT_LENGTH.value,
            ),
            enabled=data.get("enabled", True),
        )


__all__ = [
    "AntifreezeAnalysisCase",
    "AntifreezePropertyInput",
    "EPANET_REFERENCE_KINEMATIC_VISCOSITY_M2_S",
    "FittingLossPolicy",
    "GRAVITY_M_S2",
    "ResolvedFluidProperties",
    "ViscosityInputBasis",
    "WATER_DENSITY_AT_4C_KG_M3",
]

