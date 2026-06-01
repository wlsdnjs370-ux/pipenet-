"""
pipenet_validator_v4.py — 기존 PipenetGuideValidator의 v4 확장

This module wraps (does NOT replace) the existing pipenet_validator.py
PipenetGuideValidator class and adds the v4 capabilities:

  - 12-scenario aggregation across multiple PIPENET runs
  - 3 imbalance metrics (ΔP, CV, τ_water) on top of 6 hard rules
  - 3-source trace (NFTC + HB + PhD) per decision
  - cross-scenario comparison and worst-case identification

The existing PIPE.001~006 rule set is preserved unchanged. This module is
purely additive — original report generation continues to work.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from nftc_rules import (
    RuleDecision,
    TripleTrace,
    Verdict,
    head_pressure_min_mpa,
    head_pressure_max_mpa,
    head_flow_min_lpm,
    emergency_power_min_minutes,
)
from hb_rules import (
    BRANCH_PIPE_V_LIMIT,
    MAIN_PIPE_V_LIMIT,
    validate_velocity,
    validate_churn_pressure,
    decide_pipe_material,
)
from phd_rules import (
    CalculationScenario,
    ImbalanceMetrics,
    calc_flow_cv,
    calc_pressure_imbalance,
    calc_water_duration,
    evaluate_imbalance,
    evaluate_imbalance_tier,
)
from change_log import TripleTracer


# ---------------------------------------------------------------------------
# 1. Per-scenario validation result
# ---------------------------------------------------------------------------


@dataclass
class ScenarioValidationResult:
    """Validation outcome for a single PIPENET scenario."""

    scenario_id: str
    zone: str
    position: str
    purpose: str
    raw_validation: dict[str, Any]            # output of PipenetGuideValidator.validate()
    six_hard_summary: dict[str, Any]          # PIPE.001~006 summary
    head_flows_lpm: list[float]
    head_pressures_mpa: list[float]
    pipe_velocities_mps: list[dict[str, Any]]  # role + velocity
    delta_p_zone: dict[str, float]
    cv_flow: float
    verdict: str                              # PASS / REVIEW / FAIL


# ---------------------------------------------------------------------------
# 2. Aggregated v4 validation
# ---------------------------------------------------------------------------


@dataclass
class V4ValidationReport:
    """Complete v4 validation report across 12 scenarios."""

    project_id: str
    scenario_results: list[ScenarioValidationResult] = field(default_factory=list)
    overall_imbalance: ImbalanceMetrics | None = None
    six_hard_aggregate: dict[str, Any] = field(default_factory=dict)
    three_imbalance_aggregate: dict[str, Any] = field(default_factory=dict)
    worst_scenario: str | None = None
    overall_verdict: str = "IN_PROGRESS"
    trace_summary: dict[str, int] = field(default_factory=dict)
    diagnosis_messages: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# 3. The v4 validator wrapper
# ---------------------------------------------------------------------------


class PipenetValidatorV4:
    """v4 wrapper around the existing PipenetGuideValidator.

    Workflow:
      1. For each (scenario, sdf_path, pipenet_pdf_path) triple,
         run PipenetGuideValidator (unchanged) and capture result.
      2. Extract head flows, pressures, velocities for ⑥.5 imbalance.
      3. Compute 3 imbalance metrics across all scenarios.
      4. Identify worst-case scenario (drives ⑦ redesign trigger).
      5. Emit V4ValidationReport with 3-source trace.
    """

    def __init__(self, project_id: str, *, tank_total_volume_m3: float = 40.0,
                 legal_duration_minutes: float = 20.0) -> None:
        self.project_id = project_id
        self.tank_total_volume_m3 = tank_total_volume_m3
        self.legal_duration_minutes = legal_duration_minutes
        self.tracer = TripleTracer()

    # --------------------------- Entry point --------------------------------

    def validate_scenarios(
        self,
        *,
        scenarios: list[CalculationScenario],
        sdf_paths_by_scenario: dict[str, Path],
        pipenet_pdf_paths_by_scenario: dict[str, Path],
        cad_path: Path | None = None,
        project_meta_path: Path | None = None,
        design_policy_path: Path | None = None,
    ) -> V4ValidationReport:
        """Validate every scenario and aggregate results.

        Each scenario must have a matching SDF + PIPENET PDF. Missing inputs
        for a scenario produce a REVIEW verdict for that scenario only.
        """
        report = V4ValidationReport(project_id=self.project_id)
        for sc in scenarios:
            sdf = sdf_paths_by_scenario.get(sc.scenario_id)
            pdf = pipenet_pdf_paths_by_scenario.get(sc.scenario_id)
            if not sdf or not pdf or not sdf.exists() or not pdf.exists():
                report.scenario_results.append(
                    self._missing_input_result(sc, sdf, pdf)
                )
                continue
            result = self._validate_one_scenario(
                sc,
                sdf_path=sdf,
                pdf_path=pdf,
                cad_path=cad_path,
                project_meta_path=project_meta_path,
                design_policy_path=design_policy_path,
            )
            report.scenario_results.append(result)
        # Aggregate
        self._aggregate(report)
        return report

    # --------------------------- One scenario -------------------------------

    def _validate_one_scenario(
        self,
        scenario: CalculationScenario,
        *,
        sdf_path: Path,
        pdf_path: Path,
        cad_path: Path | None,
        project_meta_path: Path | None,
        design_policy_path: Path | None,
    ) -> ScenarioValidationResult:
        """Run PipenetGuideValidator on one scenario and post-process."""
        try:
            from pipenet_validator import PipenetGuideValidator  # type: ignore[import-not-found]
        except ImportError:
            return self._missing_input_result(scenario, sdf_path, pdf_path)
        try:
            guide = PipenetGuideValidator(
                report_path=pdf_path,
                sdf_path=sdf_path,
                cad_path=cad_path,
                project_meta_path=project_meta_path,
                design_policy_path=design_policy_path,
            )
            raw = guide.validate()
        except Exception as exc:
            return ScenarioValidationResult(
                scenario_id=scenario.scenario_id,
                zone=scenario.zone,
                position=scenario.position,
                purpose=scenario.purpose,
                raw_validation={"error": str(exc)},
                six_hard_summary={"error": str(exc)},
                head_flows_lpm=[],
                head_pressures_mpa=[],
                pipe_velocities_mps=[],
                delta_p_zone={},
                cv_flow=0.0,
                verdict="ERROR",
            )
        # Six hard rules summary (PIPE.001~006 + HW + velocity + nozzle)
        six_hard = self._summarize_six_hard(raw)
        # Extract head flows / pressures / velocities for imbalance
        head_flows = self._extract_head_flows(raw)
        head_pressures = self._extract_head_pressures(raw)
        velocities = self._extract_velocities(raw)
        # Delta-P per zone (single scenario — minimal grouping)
        zone_pressures = {scenario.zone: head_pressures} if head_pressures else {}
        delta_p = calc_pressure_imbalance(zone_pressures)
        # CV
        cv = calc_flow_cv(head_flows)
        # Per-scenario verdict
        if six_hard.get("FAIL", []) or any(not v.get("ok", True) for v in velocities):
            verdict = "FAIL"
        elif six_hard.get("WARNING", []):
            verdict = "REVIEW"
        else:
            verdict = "PASS"
        # Trace
        self.tracer.record(
            decision_key=f"scenario.{scenario.scenario_id}.verdict",
            trace=TripleTrace(
                nftc="NFTC 103 §2.2.1.11 (P/Q) + §2.2.1.10 (체절)",
                hb="HB §2.4.5 (유속) + §2.4.16 (체절 120%)",
                phd="박사논문 §3.x — 시나리오별 검증",
            ),
            value=verdict,
        )
        return ScenarioValidationResult(
            scenario_id=scenario.scenario_id,
            zone=scenario.zone,
            position=scenario.position,
            purpose=scenario.purpose,
            raw_validation=raw,
            six_hard_summary=six_hard,
            head_flows_lpm=head_flows,
            head_pressures_mpa=head_pressures,
            pipe_velocities_mps=velocities,
            delta_p_zone=delta_p,
            cv_flow=cv,
            verdict=verdict,
        )

    # --------------------------- Extractors ---------------------------------

    def _summarize_six_hard(self, raw: dict[str, Any]) -> dict[str, Any]:
        """Pick out PASS/FAIL/WARNING for the 6 hard rules."""
        results = raw.get("results") or {}
        return {
            "PASS": list(results.get("PASS") or []),
            "FAIL": list(results.get("FAIL") or []),
            "WARNING": list(results.get("WARNING") or []),
            "pipe_rule_results": raw.get("pipe_rule_results") or [],
        }

    def _extract_head_flows(self, raw: dict[str, Any]) -> list[float]:
        tables = raw.get("tables") or {}
        flows: list[float] = []
        for row in tables.get("nozzle_flows") or []:
            q = row.get("actual_flow_lpm")
            if isinstance(q, (int, float)):
                flows.append(float(q))
        return flows

    def _extract_head_pressures(self, raw: dict[str, Any]) -> list[float]:
        tables = raw.get("tables") or {}
        ps: list[float] = []
        for row in tables.get("nozzle_flows") or []:
            p_kgcm2 = row.get("inlet_pressure_kgf_cm2")
            if isinstance(p_kgcm2, (int, float)):
                ps.append(float(p_kgcm2) / 10.197)  # → MPa
        return ps

    def _extract_velocities(self, raw: dict[str, Any]) -> list[dict[str, Any]]:
        tables = raw.get("tables") or {}
        out: list[dict[str, Any]] = []
        for row in tables.get("velocity_checks") or []:
            v = row.get("velocity_mps")
            limit = row.get("velocity_limit_mps")
            ok = row.get("velocity_ok")
            role = row.get("pipe_role") or "other"
            out.append({
                "pipe_id": row.get("label"),
                "role": role,
                "velocity_mps": v,
                "limit_mps": limit,
                "ok": ok,
            })
        return out

    # --------------------------- Aggregation --------------------------------

    def _aggregate(self, report: V4ValidationReport) -> None:
        """Roll up scenario results into report-level metrics."""
        # 6-hard aggregate
        all_fails: list[str] = []
        all_warnings: list[str] = []
        for r in report.scenario_results:
            for f in r.six_hard_summary.get("FAIL", []) or []:
                all_fails.append(f"[{r.scenario_id}] {f}")
            for w in r.six_hard_summary.get("WARNING", []) or []:
                all_warnings.append(f"[{r.scenario_id}] {w}")
        report.six_hard_aggregate = {
            "fail_count": len(all_fails),
            "warning_count": len(all_warnings),
            "fail_messages": all_fails[:50],
            "warning_messages": all_warnings[:50],
        }
        # 3-imbalance aggregate (across all scenarios)
        all_flows: list[float] = []
        zone_pressures: dict[str, list[float]] = {}
        for r in report.scenario_results:
            all_flows.extend(r.head_flows_lpm)
            zone_pressures.setdefault(r.zone, []).extend(r.head_pressures_mpa)
        if all_flows or zone_pressures:
            imbalance = evaluate_imbalance(
                head_flows_lpm=all_flows or [80.0],
                zone_pressures=zone_pressures,
                tank_total_volume_m3=self.tank_total_volume_m3,
                legal_duration_minutes=self.legal_duration_minutes,
            )
            report.overall_imbalance = imbalance
            report.three_imbalance_aggregate = {
                "tier": imbalance.tier,
                "delta_p_max_per_zone": imbalance.delta_p_max_mpa_per_zone,
                "cv_flow": imbalance.cv_flow,
                "tau_water_minutes": imbalance.tau_water_minutes,
                "duration_reduction_pct": imbalance.duration_reduction_pct,
            }
            report.diagnosis_messages.extend(imbalance.diagnosis_messages)
        # Worst scenario
        worst = None
        for r in report.scenario_results:
            if r.verdict == "FAIL":
                worst = r.scenario_id
                break
        if not worst:
            for r in report.scenario_results:
                if r.verdict == "REVIEW":
                    worst = r.scenario_id
                    break
        report.worst_scenario = worst
        # Overall verdict
        if all_fails or (report.overall_imbalance and report.overall_imbalance.tier == "redesign_required"):
            report.overall_verdict = "FAIL"
        elif all_warnings or (report.overall_imbalance and report.overall_imbalance.tier == "human_review"):
            report.overall_verdict = "REVIEW"
        else:
            report.overall_verdict = "PASS"
        # Trace summary
        report.trace_summary = self.tracer.summary()

    # --------------------------- Fallback -----------------------------------

    def _missing_input_result(
        self,
        scenario: CalculationScenario,
        sdf_path: Path | None,
        pdf_path: Path | None,
    ) -> ScenarioValidationResult:
        return ScenarioValidationResult(
            scenario_id=scenario.scenario_id,
            zone=scenario.zone,
            position=scenario.position,
            purpose=scenario.purpose,
            raw_validation={"error": "missing_input", "sdf": str(sdf_path), "pdf": str(pdf_path)},
            six_hard_summary={"PASS": [], "FAIL": [], "WARNING": ["missing input"]},
            head_flows_lpm=[],
            head_pressures_mpa=[],
            pipe_velocities_mps=[],
            delta_p_zone={},
            cv_flow=0.0,
            verdict="REVIEW",
        )


__all__ = [
    "ScenarioValidationResult",
    "V4ValidationReport",
    "PipenetValidatorV4",
]
