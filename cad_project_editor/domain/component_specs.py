"""Domain ComponentSpec models for node device-specific fields.

Step 2a keeps these specs detached from Node serialization. They provide a
named domain layer over the existing flat node meta without changing .kfp writes.
"""
from __future__ import annotations

import copy
from dataclasses import dataclass, field
from typing import Any, Protocol


class ComponentSpec(Protocol):
    type_id: str

    def to_flat_meta(self) -> dict[str, Any]:
        """Return the legacy flat meta representation for this spec."""


_SPEC_OWNED_FIELDS: dict[str, frozenset[str]] = {
    "base": frozenset(),
    "pump": frozenset(
        {
            "rated_q",
            "rated_p",
            "shutoff_p",
            "peak_q",
            "peak_p",
            "pump_library_id",
            "pump_curve_data",
        }
    ),
    "valve": frozenset({"fitting_id", "is_check_valve", "check_valve_direction"}),
    "prv": frozenset({"pressure_setting_bar", "loss_coefficient"}),
    "orifice": frozenset({"hole_diameter", "orifice_discharge_coeff", "pipe_dn"}),
    "head": frozenset({"k_factor_si", "required_pressure_bar", "head_spec_name"}),
    "nozzle": frozenset(
        {"k_factor_si", "required_pressure_bar", "nozzle_name", "base_demand_lps"}
    ),
    "wt": frozenset({"water_level", "has_check_valve"}),
    "reservoir": frozenset(),
    "hose_stream": frozenset({"hose_stream_flow_lps"}),
}


def component_spec_owned_field_names(type_id: str) -> frozenset[str]:
    """Return flat Node fields owned by the ComponentSpec for a type id."""
    return _SPEC_OWNED_FIELDS.get(str(type_id or "").strip().lower(), frozenset())


def component_spec_from_flat_meta(meta: dict[str, Any]) -> ComponentSpec:
    """Build a ComponentSpec from existing flat node meta.

    Unknown type ids are preserved as GenericSpec instead of raising. Loading
    future 3.x files must remain possible even before this version understands
    their semantics.
    """
    meta = meta if isinstance(meta, dict) else {}
    type_id = str(meta.get("type_id") or "base").strip() or "base"

    if type_id == "pump":
        return PumpSpec(
            rated_q=_float(meta.get("rated_q")),
            rated_p=_float(meta.get("rated_p")),
            shutoff_p=_float(meta.get("shutoff_p")),
            peak_q=_float(meta.get("peak_q")),
            peak_p=_float(meta.get("peak_p")),
            pump_library_id=_optional_str(meta.get("pump_library_id")),
            pump_curve_data=_list_copy(meta.get("pump_curve_data")),
        )
    if type_id == "valve":
        return ValveSpec(
            fitting_id=_optional_str(meta.get("fitting_id")),
            is_check_valve=_bool(meta.get("is_check_valve")),
            check_valve_direction=str(meta.get("check_valve_direction") or ""),
        )
    if type_id == "prv":
        return PrvSpec(
            pressure_setting_bar=_optional_float(meta.get("pressure_setting_bar")),
            loss_coefficient=_float(meta.get("loss_coefficient")),
        )
    if type_id == "orifice":
        return OrificeSpec(
            hole_diameter=_float(meta.get("hole_diameter")),
            orifice_discharge_coeff=_float(meta.get("orifice_discharge_coeff"), 0.65),
            pipe_dn=_float(meta.get("pipe_dn")),
        )
    if type_id == "head":
        return HeadSpec(
            k_factor_si=_optional_float(meta.get("k_factor_si", meta.get("k_factor"))),
            required_pressure_bar=_float(meta.get("required_pressure_bar")),
            head_spec_name=_optional_str(meta.get("head_spec_name")),
        )
    if type_id == "nozzle":
        return NozzleSpec(
            k_factor_si=_optional_float(meta.get("k_factor_si", meta.get("k_factor"))),
            required_pressure_bar=_float(meta.get("required_pressure_bar")),
            nozzle_name=_optional_str(meta.get("nozzle_name")),
            base_demand_lps=_float(meta.get("base_demand_lps")),
        )
    if type_id == "wt":
        return TankSpec(
            water_level=_float(meta.get("water_level")),
            has_check_valve=_bool(meta.get("has_check_valve"), True),
        )
    if type_id == "reservoir":
        return ReservoirSpec()
    if type_id == "hose_stream":
        return HoseStreamSpec(
            hose_stream_flow_lps=_float(meta.get("hose_stream_flow_lps"), 400.0 / 60.0),
        )
    if type_id == "base":
        return BaseSpec()

    return GenericSpec(
        type_id=type_id,
        payload={
            key: copy.deepcopy(value)
            for key, value in meta.items()
            if key not in {"type_id", "category_id", "type"}
        },
    )


def flat_meta_from_component_spec(spec: ComponentSpec) -> dict[str, Any]:
    """Return legacy flat meta for a ComponentSpec."""
    return spec.to_flat_meta()


def infer_component_type_id_from_flat_meta(meta: dict[str, Any] | None) -> str | None:
    """Infer a component type id from legacy flat meta values.

    This mirrors the field-value portion of normalize_node_type_id without
    applying display-name or explicit-base rules. Callers must keep those guards
    at the edge where stale metadata can be disambiguated.
    """
    if not isinstance(meta, dict):
        return None

    if meta.get("nozzle_name") or meta.get("nozzle_spec"):
        return "nozzle"
    if meta.get("head_spec_name"):
        return "head"
    if meta.get("pump_curve_data") or _positive(meta.get("rated_q")):
        return "pump"
    if _positive(meta.get("water_level")):
        return "wt"
    if meta.get("pressure_setting_bar") is not None:
        return "prv"
    if meta.get("fitting_id"):
        return "valve"
    if _positive(meta.get("hole_diameter")):
        return "orifice"

    k_factor = meta.get("k_factor_si") or meta.get("k_factor")
    if _positive(k_factor):
        return "head"

    return None


@dataclass(frozen=True)
class BaseSpec:
    type_id: str = "base"

    def to_flat_meta(self) -> dict[str, Any]:
        return {"type_id": self.type_id}


@dataclass(frozen=True)
class PumpSpec:
    type_id: str = "pump"
    rated_q: float = 0.0
    rated_p: float = 0.0
    shutoff_p: float = 0.0
    peak_q: float = 0.0
    peak_p: float = 0.0
    pump_library_id: str | None = None
    pump_curve_data: list[Any] = field(default_factory=list)

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "rated_q": self.rated_q,
            "rated_p": self.rated_p,
            "shutoff_p": self.shutoff_p,
            "peak_q": self.peak_q,
            "peak_p": self.peak_p,
            "pump_library_id": self.pump_library_id,
            "pump_curve_data": copy.deepcopy(self.pump_curve_data),
        }


@dataclass(frozen=True)
class ValveSpec:
    type_id: str = "valve"
    fitting_id: str | None = None
    is_check_valve: bool = False
    check_valve_direction: str = ""

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "fitting_id": self.fitting_id,
            "is_check_valve": self.is_check_valve,
            "check_valve_direction": self.check_valve_direction,
        }


@dataclass(frozen=True)
class PrvSpec:
    type_id: str = "prv"
    pressure_setting_bar: float | None = None
    loss_coefficient: float = 0.0

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "pressure_setting_bar": self.pressure_setting_bar,
            "loss_coefficient": self.loss_coefficient,
        }


@dataclass(frozen=True)
class OrificeSpec:
    type_id: str = "orifice"
    hole_diameter: float = 0.0
    orifice_discharge_coeff: float = 0.65
    pipe_dn: float = 0.0

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "hole_diameter": self.hole_diameter,
            "orifice_discharge_coeff": self.orifice_discharge_coeff,
            "pipe_dn": self.pipe_dn,
        }


@dataclass(frozen=True)
class HeadSpec:
    type_id: str = "head"
    k_factor_si: float | None = None
    required_pressure_bar: float = 0.0
    head_spec_name: str | None = None

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "k_factor_si": self.k_factor_si,
            "required_pressure_bar": self.required_pressure_bar,
            "head_spec_name": self.head_spec_name,
        }


@dataclass(frozen=True)
class NozzleSpec:
    type_id: str = "nozzle"
    k_factor_si: float | None = None
    required_pressure_bar: float = 0.0
    nozzle_name: str | None = None
    base_demand_lps: float = 0.0

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "k_factor_si": self.k_factor_si,
            "required_pressure_bar": self.required_pressure_bar,
            "nozzle_name": self.nozzle_name,
            "base_demand_lps": self.base_demand_lps,
        }


@dataclass(frozen=True)
class TankSpec:
    type_id: str = "wt"
    water_level: float = 0.0
    has_check_valve: bool = True

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "water_level": self.water_level,
            "has_check_valve": self.has_check_valve,
        }


@dataclass(frozen=True)
class ReservoirSpec:
    type_id: str = "reservoir"

    def to_flat_meta(self) -> dict[str, Any]:
        return {"type_id": self.type_id}


@dataclass(frozen=True)
class HoseStreamSpec:
    type_id: str = "hose_stream"
    hose_stream_flow_lps: float = 400.0 / 60.0

    def to_flat_meta(self) -> dict[str, Any]:
        return {
            "type_id": self.type_id,
            "hose_stream_flow_lps": self.hose_stream_flow_lps,
        }


@dataclass(frozen=True)
class GenericSpec:
    type_id: str
    payload: dict[str, Any] = field(default_factory=dict)

    def to_flat_meta(self) -> dict[str, Any]:
        return {"type_id": self.type_id, **copy.deepcopy(self.payload)}


def _float(value: Any, default: float = 0.0) -> float:
    if value is None or value == "":
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _optional_float(value: Any) -> float | None:
    if value is None or value == "":
        return None
    return _float(value)


def _positive(value: Any) -> bool:
    return _float(value) > 0.0


def _bool(value: Any, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return default
    return bool(value)


def _optional_str(value: Any) -> str | None:
    if value is None:
        return None
    return str(value)


def _list_copy(value: Any) -> list[Any]:
    if isinstance(value, list):
        return copy.deepcopy(value)
    return []
