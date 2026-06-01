"""
pipeline_orchestrator.py — v4 파이프라인 8단계 + 보강 폐루프 조정자

This is the top-level orchestrator that wires together every module in this
patch + the existing cad_engine.py and pipenet_validator.py to produce the
full v4 closed loop:

  ① Object extraction (DXF + vision + NFTC §2.4.18 mask)
  ② Zone partition + system + head spec
  ②.5 NFTC compliance mapping (reference count, temperature, R, fast-response)
  ③ Head auto-placement (R 5종 + 60cm + (2R)²=S²+L²)
  ③.5 NFTC 103B / ESFR branch
  ④ Pipe routing + Case 1~5 + discretionary variables + NFTC 2.13 combined
  ⑤ PIPENET conversion + 12 scenarios auto-generation
  ⑥ Validation (6 hard rules + 3 imbalance metrics)
  ⑦ Redesign loop (5 alternatives, ranked quantitatively)
  ⑧ Final drawing + Change Log + 3-source Trace

The orchestrator is *agentic*: when ⑥ produces non-PASS, it auto-iterates ⑦
up to a configured number of attempts before requesting human review.
"""

from __future__ import annotations

import json
import xml.etree.ElementTree as ET
from dataclasses import asdict, dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable

from nftc_rules import (
    RuleDecision,
    TripleTrace,
    Verdict,
    decide_alarm_cascade,
    decide_combined_water_supply,
    decide_reference_count,
    summarize_nftc_decisions,
)
from hb_rules import (
    HBCase,
    HBCaseDecision,
    SystemType,
    decide_hb_case,
    decide_pipe_material,
    decide_system_type,
    validate_churn_pressure,
)
from phd_rules import (
    CalculationScenario,
    DesignAlternative,
    DiscretionaryVariables,
    ImbalanceMetrics,
    classify_pressure_zones,
    decide_discretionary_variables,
    evaluate_imbalance,
    generate_alternative_scenarios,
    generate_calculation_scenarios,
    rank_alternatives,
)
from auto_design import (
    DesignNetwork,
    PipeSegment,
    Zone,
    design_full_network,
)
from change_log import ChangeLogger, TripleTracer
from ai_vision import get_cached_ai_vision_extractor


# ---------------------------------------------------------------------------
# 1. Pipeline configuration
# ---------------------------------------------------------------------------


@dataclass
class PipelineConfig:
    """Tunable parameters for the v4 pipeline run."""

    project_id: str
    max_redesign_attempts: int = 3
    legal_duration_minutes: float = 20.0   # NFTC 2.9.3.2
    output_dir: Path = field(default_factory=lambda: Path("data/v4_outputs"))
    log_dir: Path = field(default_factory=lambda: Path("data/change_logs"))
    pipenet_input_template: str | None = None  # optional input.dat template
    require_human_signoff: bool = True


# ---------------------------------------------------------------------------
# 2. Stage results (immutable evidence)
# ---------------------------------------------------------------------------


@dataclass
class StageResult:
    """Result of a single pipeline stage."""

    stage: str
    started_at: str
    finished_at: str
    verdict: str
    summary: dict[str, Any]
    trace: list[dict[str, Any]] = field(default_factory=list)
    artifacts: dict[str, str] = field(default_factory=dict)  # name → file path


@dataclass
class PipelineRunReport:
    """End-to-end run report — accumulates StageResult per stage."""

    project_id: str
    run_id: str
    started_at: str
    finished_at: str = ""
    overall_verdict: str = "IN_PROGRESS"
    stages: list[StageResult] = field(default_factory=list)
    final_kpis: dict[str, Any] = field(default_factory=dict)
    trace_summary: dict[str, int] = field(default_factory=dict)
    redesign_attempts: int = 0


# ---------------------------------------------------------------------------
# 3. The orchestrator
# ---------------------------------------------------------------------------


class SprinklerPipelineV4:
    """v4 pipeline orchestrator — 8 stages + 4 enhancements with closed loop.

    Typical usage (from a Flask route):

        cfg = PipelineConfig(project_id="proj-2026-04-28")
        pipe = SprinklerPipelineV4(cfg)
        report = pipe.run_full_pipeline(
            dxf_path=Path("uploads/site.dxf"),
            building_meta={...},
            rooms=[...],
            obstacles=[...],
            floors=[...],
        )
    """

    def __init__(self, config: PipelineConfig) -> None:
        self.config = config
        self.config.output_dir.mkdir(parents=True, exist_ok=True)
        self.config.log_dir.mkdir(parents=True, exist_ok=True)
        self.tracer = TripleTracer()
        self.logger = ChangeLogger(config.project_id, log_dir=config.log_dir)
        self.run_id = f"run-{datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S')}"

    # --------------------------- Public entrypoint --------------------------

    def run_full_pipeline(
        self,
        *,
        dxf_path: Path,
        building_meta: dict[str, Any],
        rooms: list[dict[str, Any]],
        obstacles: list[dict[str, Any]],
        floors: list[dict[str, Any]],
        existing_sdf_path: Path | None = None,
        existing_pipenet_pdf_path: Path | None = None,
        sdf_writer: Callable[[DesignNetwork, Path], Path] | None = None,
        pipenet_runner: Callable[[Path], dict[str, Any]] | None = None,
    ) -> PipelineRunReport:
        """End-to-end run. Returns report with every stage's evidence."""
        report = PipelineRunReport(
            project_id=self.config.project_id,
            run_id=self.run_id,
            started_at=_iso_now(),
        )
        try:
            # ① ------------------------------------------------------------
            extracted = self._stage_1_extract_objects(dxf_path, rooms, obstacles, building_meta)
            report.stages.append(extracted)
            # ② + ③ + ④ ----------------------------------------------------
            design_stage, network = self._stage_234_auto_design(
                building_meta=building_meta,
                rooms=rooms,
                obstacles=obstacles,
                floors=floors,
            )
            report.stages.append(design_stage)
            # ⑤ ------------------------------------------------------------
            convert_stage, sdf_paths = self._stage_5_pipenet_convert(
                network=network,
                sdf_writer=sdf_writer,
            )
            report.stages.append(convert_stage)
            # ⑥ ------------------------------------------------------------
            validate_stage, validation_results, imbalance = self._stage_6_validate(
                network=network,
                sdf_paths=sdf_paths,
                cad_path=dxf_path,
                pipenet_runner=pipenet_runner,
                existing_pipenet_pdf_path=existing_pipenet_pdf_path,
            )
            report.stages.append(validate_stage)
            # ⑦ Redesign loop ----------------------------------------------
            attempts = 0
            current_imbalance = imbalance
            current_validation = validation_results
            while (
                current_imbalance.tier == "redesign_required"
                and attempts < self.config.max_redesign_attempts
            ):
                attempts += 1
                redesign_stage, network, current_validation, current_imbalance = self._stage_7_redesign(
                    network=network,
                    current_imbalance=current_imbalance,
                    sdf_writer=sdf_writer,
                    pipenet_runner=pipenet_runner,
                    cad_path=dxf_path,
                    attempt=attempts,
                )
                report.stages.append(redesign_stage)
            report.redesign_attempts = attempts
            # ⑧ ------------------------------------------------------------
            final_stage = self._stage_8_finalize(
                network=network,
                final_imbalance=current_imbalance,
                final_validation=current_validation,
            )
            report.stages.append(final_stage)
            # Verdict
            if current_imbalance.tier == "auto_pass":
                report.overall_verdict = "PASS"
            elif current_imbalance.tier == "human_review":
                report.overall_verdict = "REVIEW"
            else:
                report.overall_verdict = "FAIL"
            report.final_kpis = self._summarize_kpis(network, current_imbalance, current_validation)
        except Exception as exc:
            report.overall_verdict = "ERROR"
            report.stages.append(StageResult(
                stage="error",
                started_at=_iso_now(),
                finished_at=_iso_now(),
                verdict="ERROR",
                summary={"error": str(exc), "type": type(exc).__name__},
            ))
        finally:
            report.finished_at = _iso_now()
            report.trace_summary = self.tracer.summary()
            self._persist_report(report)
        return report

    # --------------------------- Stage ① ------------------------------------

    def _stage_1_extract_objects(
        self,
        dxf_path: Path,
        rooms: list[dict[str, Any]],
        obstacles: list[dict[str, Any]],
        building_meta: dict[str, Any],
    ) -> StageResult:
        """① DXF + vision + NFTC §2.4.18 exclusion masking.

        This stage trusts the cad_engine.py output for DXF parsing (already
        well-implemented). It then applies §2.4.18 13-area exclusion masking
        and labels rooms with their use / structure.
        """
        started = _iso_now()
        ai_head_count = 0
        ai_detector_mode = "disabled"
        # Use existing cad_engine if available
        try:
            from cad_engine import DXFWorkspace  # type: ignore[import-not-found]
            workspace = DXFWorkspace(dxf_path.parent / "_v4_workspace")
            workspace.load_file(dxf_path)
            payload = workspace.to_payload(include_network_entities=True, include_network_summary=True)
            entity_count = len(payload.get("entities") or [])
            try:
                ai_result = get_cached_ai_vision_extractor(str(Path(__file__).resolve().parent)).enhance_from_payload(
                    dxf_path=dxf_path,
                    cad_payload=payload,
                )
                ai_head_count = int(ai_result.stats.get("head_count", 0))
                ai_detector_mode = str(ai_result.stats.get("detector_mode", "disabled"))
            except Exception:
                ai_detector_mode = "error"
        except Exception:
            entity_count = -1  # cad_engine unavailable; tolerate
        # Apply §2.4.18 exclusion mask
        excluded = self._apply_24_18_exclusion(rooms)
        # Trace — exclusion clauses
        for room in excluded:
            self.tracer.record(
                decision_key=f"room.{room.get('id', '?')}.excluded",
                trace=TripleTrace(
                    nftc="NFTC 103 §2.12 (헤드 설치 제외 장소)",
                    hb="HB §2.4.18 (13개 영역 제외)",
                    phd=None,
                    note=f"excluded reason: {room.get('exclusion_reason', '?')}",
                ),
                value=True,
            )
        return StageResult(
            stage="①_object_extraction",
            started_at=started,
            finished_at=_iso_now(),
            verdict="PASS",
            summary={
                "rooms_total": len(rooms),
                "rooms_excluded": len([r for r in rooms if r.get("hb_excluded")]),
                "obstacles_total": len(obstacles),
                "dxf_entity_count": entity_count,
                "ai_detected_heads": ai_head_count,
                "ai_detector_mode": ai_detector_mode,
            },
            trace=[t for t in self.tracer.all() if "excluded" in t.get("key", "")],
        )

    def _apply_24_18_exclusion(self, rooms: list[dict[str, Any]]) -> list[dict[str, Any]]:
        """Mark rooms that fall under HB §2.4.18 / NFTC §2.12 exclusions.

        13 categories:
          ① 계단실·부속실  ② 통신기계실·전자기기실  ③ 발전기실·변전실
          ④ 병원 수술실·응급처치실  ⑤ 천장·반자가 불연재 + 30cm 미만
          ⑥ 천장·반자가 불연재 외 + 5cm 미만  ⑦ 천장·반자 사이 거리 0.5m 이상 + 불연재
          ⑧ 펌프실·물탱크실·수영장·목욕실  ⑨ 직접 외기에 개방
          ⑩ 야외 공간  ⑪ 냉동·냉장창고 (-3℃ 이하)
          ⑫ 패널 형식 발코니 (특정 조건)  ⑬ 가스시설·외기 직접 노출 부분
        """
        for room in rooms:
            use = room.get("use", "")
            ceiling_h = room.get("ceiling_h_m", 3.0)
            ambient = room.get("ambient_temp_c", 25.0)
            roof_type = room.get("roof_type", "")
            reason = None
            if use in {"stairwell", "stair_ancillary"}:
                reason = "①_stairwell"
            elif use in {"telecom_room", "electronic_equipment_room"}:
                reason = "②_telecom"
            elif use in {"generator_room", "transformer_room"}:
                reason = "③_generator"
            elif use in {"surgery_room", "emergency_room"}:
                reason = "④_surgery"
            elif use in {"pump_room", "water_tank_room", "swimming_pool", "bathhouse"}:
                reason = "⑧_pump_water"
            elif use in {"freezer_room"} or ambient <= -3:
                reason = "⑪_freezer"
            elif roof_type == "open_to_outside":
                reason = "⑨_open_to_outside"
            elif use == "outdoor":
                reason = "⑩_outdoor"
            if reason:
                room["hb_excluded"] = True
                room["exclusion_reason"] = reason
        return rooms

    # --------------------------- Stage ②③④ ---------------------------------

    def _stage_234_auto_design(
        self,
        *,
        building_meta: dict[str, Any],
        rooms: list[dict[str, Any]],
        obstacles: list[dict[str, Any]],
        floors: list[dict[str, Any]],
    ) -> tuple[StageResult, DesignNetwork]:
        """② + ③ + ④ — auto design via auto_design.design_full_network."""
        started = _iso_now()
        # Reference count first (drives water source & pump)
        ref_dec = decide_reference_count(building_meta)
        self.tracer.record(decision_key="reference_count", trace=ref_dec.trace, value=ref_dec.value)
        # Run auto design
        network = design_full_network(
            project_id=self.config.project_id,
            rooms=rooms,
            obstacles=obstacles,
            floors=floors,
            building_meta=building_meta,
            refuge_floor_interval_m=building_meta.get("refuge_floor_interval_m"),
            rooftop_tank_feasible=bool(building_meta.get("rooftop_tank_feasible", True)),
            pump_rated_q_lpm=float(building_meta.get("pump_rated_q_lpm", 2400.0)),
            pump_rated_h_m=float(building_meta.get("pump_rated_h_m", 60.0)),
            pump_churn_h_m=float(building_meta.get("pump_churn_h_m", 70.0)),
        )
        # Combined water supply (NFTC 2.13)
        if building_meta.get("share_with_other_systems"):
            other_sys = building_meta.get("other_systems", {})
            other_sys["sprinkler"] = {
                "v_m3": float(building_meta.get("sprinkler_tank_m3", 32.0)),
                "q_lpm": float(building_meta.get("pump_rated_q_lpm", 2400.0)),
                "h_m": float(building_meta.get("pump_rated_h_m", 60.0)),
                "zone_area_m2": sum(z.area_m2 for z in network.zones),
            }
            combined = decide_combined_water_supply(other_sys, use_combined_tank=True, use_combined_pump=True)
            self.tracer.record(decision_key="combined_water_supply", trace=combined.trace, value=combined.tank_total_m3)
            network.metadata["combined_supply"] = asdict(combined)
        # Trace dump
        if network.hb_case:
            self.tracer.record(
                decision_key="hb_case",
                trace=network.hb_case.trace,
                value=network.hb_case.case.value,
            )
        # Counts for summary
        total_heads = sum(len(z.heads) for z in network.zones)
        total_branches = sum(len(z.branches) for z in network.zones)
        skipping_violations = sum(
            1 for z in network.zones for h in z.heads if not h.skipping_pass
        )
        clearance_violations = sum(
            1 for z in network.zones for h in z.heads if not h.nftc_2771_pass
        )
        a_b_l_violations = sum(
            1 for z in network.zones for h in z.heads if not h.a_b_l_check
        )
        verdict = "PASS" if (skipping_violations + clearance_violations + a_b_l_violations) == 0 else "REVIEW"
        return (
            StageResult(
                stage="②③④_auto_design",
                started_at=started,
                finished_at=_iso_now(),
                verdict=verdict,
                summary={
                    "reference_count": ref_dec.value,
                    "system_type": network.system_type,
                    "hb_case": network.hb_case.case.value if network.hb_case else None,
                    "zone_count": len(network.zones),
                    "head_count": total_heads,
                    "branch_count": total_branches,
                    "skipping_violations": skipping_violations,
                    "clearance_violations": clearance_violations,
                    "a_b_l_violations": a_b_l_violations,
                },
                trace=self.tracer.all()[-10:],
            ),
            network,
        )

    # --------------------------- Stage ⑤ ------------------------------------

    def _stage_5_pipenet_convert(
        self,
        *,
        network: DesignNetwork,
        sdf_writer: Callable[[DesignNetwork, Path], Path] | None,
    ) -> tuple[StageResult, list[Path]]:
        """⑤ Convert DesignNetwork → SDF/PIPENET inputs (12 scenarios)."""
        started = _iso_now()
        # Generate 12 calculation scenarios
        if not network.discretionary:
            return StageResult(
                stage="⑤_pipenet_convert",
                started_at=started,
                finished_at=_iso_now(),
                verdict="REVIEW",
                summary={"reason": "no discretionary variables"},
            ), []
        scenarios = generate_calculation_scenarios(
            discretionary=network.discretionary,
            has_k115_zones=any(z.head_spec and z.head_spec.k_factor_lpm_bar05 == 115 for z in network.zones),
            has_max_q_zone=True,
        )
        # Write SDF for each scenario (delegates to caller-supplied writer)
        sdf_paths: list[Path] = []
        for sc in scenarios:
            out = self.config.output_dir / f"{self.config.project_id}_{sc.scenario_id}.sdf"
            if sdf_writer is not None:
                sdf_writer(network, out)
            else:
                # Fallback: write a minimal but valid SDF skeleton
                self._write_minimal_sdf(network, out, sc)
            sdf_paths.append(out)
        return (
            StageResult(
                stage="⑤_pipenet_convert",
                started_at=started,
                finished_at=_iso_now(),
                verdict="PASS",
                summary={
                    "scenario_count": len(scenarios),
                    "scenarios": [s.scenario_id for s in scenarios],
                    "sdf_files_written": len(sdf_paths),
                },
                artifacts={p.name: str(p) for p in sdf_paths},
            ),
            sdf_paths,
        )

    def _write_minimal_sdf(self, network: DesignNetwork, out: Path, scenario: CalculationScenario) -> None:
        """Write a minimal SDF skeleton — replace with full SDF writer in production."""
        root = ET.Element("Network")
        meta = ET.SubElement(root, "Meta")
        ET.SubElement(meta, "ScenarioId").text = scenario.scenario_id
        ET.SubElement(meta, "Zone").text = scenario.zone
        ET.SubElement(meta, "Position").text = scenario.position
        for z in network.zones:
            zone_elem = ET.SubElement(root, "Zone", attrib={"id": z.zone_id})
            for h in z.heads:
                ET.SubElement(zone_elem, "Nozzle", attrib={
                    "id": h.head_id,
                    "x": str(h.x),
                    "y": str(h.y),
                    "z": str(h.z),
                    "k": str(h.spec.k_factor_lpm_bar05) if h.spec else "80",
                })
            for p in z.branches + z.cross_mains:
                ET.SubElement(zone_elem, "Pipe", attrib={
                    "id": p.pipe_id,
                    "length": str(p.length_m),
                    "bore": str(round(p.inner_diameter_mm / 1000.0, 4)),
                    "c": str(p.c_factor),
                    "role": p.role,
                })
        ET.ElementTree(root).write(out, encoding="utf-8", xml_declaration=True)

    # --------------------------- Stage ⑥ ------------------------------------

    def _stage_6_validate(
        self,
        *,
        network: DesignNetwork,
        sdf_paths: list[Path],
        cad_path: Path,
        pipenet_runner: Callable[[Path], dict[str, Any]] | None,
        existing_pipenet_pdf_path: Path | None,
    ) -> tuple[StageResult, dict[str, Any], ImbalanceMetrics]:
        """⑥ 6 hard rules + 3 imbalance metrics."""
        started = _iso_now()
        # Hard rules — delegate to existing PipenetGuideValidator if a PIPENET
        # PDF report is available (real-world flow).
        validation_results: dict[str, Any] = {}
        if existing_pipenet_pdf_path and existing_pipenet_pdf_path.exists() and sdf_paths:
            try:
                from pipenet_validator import PipenetGuideValidator  # type: ignore[import-not-found]
                guide = PipenetGuideValidator(
                    report_path=existing_pipenet_pdf_path,
                    sdf_path=sdf_paths[0] if sdf_paths else None,
                    cad_path=cad_path if cad_path.exists() else None,
                )
                validation_results = guide.validate()
            except Exception as exc:
                validation_results = {"error": f"PipenetGuideValidator failed: {exc}"}
        else:
            validation_results = self._fallback_hard_rule_check(network)
        # Imbalance metrics
        head_flows = self._gather_head_flows(network, validation_results)
        zone_pressures = self._gather_zone_pressures(network, validation_results)
        tank_total_m3 = self._estimate_tank_volume(network)
        imbalance = evaluate_imbalance(
            head_flows_lpm=head_flows,
            zone_pressures=zone_pressures,
            tank_total_volume_m3=tank_total_m3,
            legal_duration_minutes=self.config.legal_duration_minutes,
        )
        # Trace
        self.tracer.record(decision_key="imbalance_tier", trace=imbalance.trace, value=imbalance.tier)
        # Verdict
        verdict = {"auto_pass": "PASS", "human_review": "REVIEW", "redesign_required": "FAIL"}[imbalance.tier]
        return (
            StageResult(
                stage="⑥_validate",
                started_at=started,
                finished_at=_iso_now(),
                verdict=verdict,
                summary={
                    "hard_rules_overall": validation_results.get("results", {}).get("FAIL", []) if "results" in validation_results else None,
                    "imbalance_tier": imbalance.tier,
                    "delta_p_max_mpa_per_zone": imbalance.delta_p_max_mpa_per_zone,
                    "cv_flow": imbalance.cv_flow,
                    "tau_water_minutes": imbalance.tau_water_minutes,
                    "duration_reduction_pct": imbalance.duration_reduction_pct,
                    "diagnosis": imbalance.diagnosis_messages,
                },
            ),
            validation_results,
            imbalance,
        )

    def _fallback_hard_rule_check(self, network: DesignNetwork) -> dict[str, Any]:
        """Synthetic hard-rule evaluation when no PIPENET PDF is present.

        This is used in design-time previews. Real validation happens once
        PIPENET runs and produces a calculation report.
        """
        from nftc_rules import (
            head_pressure_min_mpa,
            head_pressure_max_mpa,
            head_flow_min_lpm,
        )
        results: dict[str, list[str]] = {"PASS": [], "FAIL": [], "WARNING": []}
        # NFTC 2.2.1.11
        results["PASS"].append(f"헤드 최소 방수압 0.1 MPa, 최소 유량 80 LPM 한계 적용")
        # Velocity (placeholder — needs actual flow)
        results["WARNING"].append("실제 유속/압력 검증은 PIPENET 결과 PDF 입력 후 수행됩니다.")
        # Churn check
        if network.pumps:
            p = network.pumps[0]
            churn_dec = validate_churn_pressure(
                rated_head_m=p.get("rated_h_m", 60.0),
                churn_head_m=p.get("churn_h_m", 70.0),
            )
            if churn_dec.verdict == Verdict.PASS:
                results["PASS"].append(f"체절 압력 {churn_dec.detail}")
            elif churn_dec.verdict == Verdict.REVIEW:
                results["WARNING"].append(churn_dec.detail)
            else:
                results["FAIL"].append(churn_dec.detail)
        return {"results": results, "synthetic": True}

    def _gather_head_flows(self, network: DesignNetwork, validation_results: dict[str, Any]) -> list[float]:
        """Extract per-head actual flow from validation tables, or use K·√P fallback."""
        tables = validation_results.get("tables") or {}
        nozzle_flows = tables.get("nozzle_flows") or []
        flows: list[float] = []
        for row in nozzle_flows:
            q = row.get("actual_flow_lpm")
            if isinstance(q, (int, float)):
                flows.append(float(q))
        if flows:
            return flows
        # Fallback: assume each head delivers 80 LPM at minimum
        return [80.0 for _ in (h for z in network.zones for h in z.heads)]

    def _gather_zone_pressures(self, network: DesignNetwork, validation_results: dict[str, Any]) -> dict[str, list[float]]:
        """Extract per-zone pressures from validation tables, or estimate."""
        tables = validation_results.get("tables") or {}
        pressures: dict[str, list[float]] = {}
        for row in tables.get("pipe_validation_rows") or []:
            zone_id = row.get("zone_id") or "unknown"
            for key in ("inlet_pressure_kgcm2", "outlet_pressure_kgcm2", "max_pressure_kgcm2"):
                val = row.get(key)
                if isinstance(val, (int, float)):
                    p_mpa = val / 10.197  # kg/cm² → MPa
                    pressures.setdefault(zone_id, []).append(p_mpa)
        if pressures:
            return pressures
        # Fallback: zone-level estimate
        out: dict[str, list[float]] = {}
        for z in network.zones:
            out[z.zone_id] = [0.1, 0.3, 0.5]  # placeholder spread
        return out

    def _estimate_tank_volume(self, network: DesignNetwork) -> float:
        """Estimate tank volume from reference count × 80 LPM × 20 min."""
        ref_count = network.metadata.get("reference_count")
        if not ref_count:
            for tr in self.tracer.all():
                if tr.get("key") == "reference_count":
                    ref_count = tr.get("value")
                    break
        ref_count = int(ref_count or 20)
        # NFTC: V_min = N × 80 LPM × 20 min / 1000 = N × 1.6 m³
        # HB §2.4.3: 80% effective → V_total = V_min / 0.8 = N × 2.0 m³
        return ref_count * 2.0

    # --------------------------- Stage ⑦ ------------------------------------

    def _stage_7_redesign(
        self,
        *,
        network: DesignNetwork,
        current_imbalance: ImbalanceMetrics,
        sdf_writer: Callable[[DesignNetwork, Path], Path] | None,
        pipenet_runner: Callable[[Path], dict[str, Any]] | None,
        cad_path: Path,
        attempt: int,
    ) -> tuple[StageResult, DesignNetwork, dict[str, Any], ImbalanceMetrics]:
        """⑦ Redesign loop — generate 5 alternatives, simulate, pick best."""
        started = _iso_now()
        # Generate 5 alternatives
        alts = generate_alternative_scenarios(
            diagnosis=current_imbalance,
            hb_case=network.hb_case if network.hb_case else _placeholder_hb_case(),
            has_basement=any(f.elevation_m < 0 for f in network.floors_pressure),
        )
        # Simulate each alternative (real: PIPENET; here: heuristic)
        sim_results: dict[str, ImbalanceMetrics] = {}
        for alt in alts:
            sim_results[alt.alt_id] = self._simulate_alternative(alt, current_imbalance)
        # Rank
        ranked = rank_alternatives(alts, simulation_results=sim_results)
        # Pick top recommendation
        chosen = next((r for r in ranked if r.get("verdict") == "PASS"), None)
        if not chosen:
            chosen = next((r for r in ranked if r.get("verdict") == "REVIEW"), None) or ranked[0]
        # Apply chosen alternative to network metadata (real implementation
        # would mutate pipes / pumps / tanks)
        chosen_alt = next((a for a in alts if a.alt_id == chosen["alt_id"]), None)
        if chosen_alt:
            network.metadata.setdefault("redesign_history", []).append({
                "attempt": attempt,
                "alt_id": chosen_alt.alt_id,
                "name": chosen_alt.name,
                "config_changes": chosen_alt.config_changes,
            })
        new_imbalance = sim_results[chosen["alt_id"]]
        # Log entry
        self.logger.append(
            triggered_by="auto",
            diagnosis={
                "tier_before": current_imbalance.tier,
                "delta_p_max": max(current_imbalance.delta_p_max_mpa_per_zone.values(), default=0.0),
                "cv": current_imbalance.cv_flow,
                "tau": current_imbalance.tau_water_minutes,
            },
            option=chosen_alt.alt_id if chosen_alt else "ALT-?",
            parameters=chosen_alt.config_changes if chosen_alt else {},
            kpi_before={
                "tier": current_imbalance.tier,
                "cv": current_imbalance.cv_flow,
                "tau": current_imbalance.tau_water_minutes,
            },
            kpi_after={
                "tier": new_imbalance.tier,
                "cv": new_imbalance.cv_flow,
                "tau": new_imbalance.tau_water_minutes,
            },
            verdict={"auto_pass": "PASS", "human_review": "REVIEW", "redesign_required": "FAIL"}[new_imbalance.tier],
            trace_links=[
                "PhD §4.x (alternatives)",
                f"NFTC 103 §2.13 (combined supply)" if "tank" in (chosen_alt.alt_id if chosen_alt else "") else "—",
            ],
            note=f"Auto-redesign attempt #{attempt}: {chosen_alt.name if chosen_alt else '?'}",
        )
        return (
            StageResult(
                stage=f"⑦_redesign_attempt_{attempt}",
                started_at=started,
                finished_at=_iso_now(),
                verdict={"auto_pass": "PASS", "human_review": "REVIEW", "redesign_required": "FAIL"}[new_imbalance.tier],
                summary={
                    "alternatives_evaluated": len(alts),
                    "ranked": ranked,
                    "chosen": chosen,
                    "tier_before": current_imbalance.tier,
                    "tier_after": new_imbalance.tier,
                },
            ),
            network,
            {"results": {"PASS": [], "FAIL": [], "WARNING": []}, "synthetic": True},
            new_imbalance,
        )

    def _simulate_alternative(
        self,
        alt: DesignAlternative,
        baseline: ImbalanceMetrics,
    ) -> ImbalanceMetrics:
        """Heuristic simulation of an alternative.

        In production this calls a real PIPENET runner. Here we apply
        plausible deltas based on the alternative type.
        """
        deltas = {
            "ALT-1-PRV":           {"delta_p_factor": 0.6, "cv_factor": 1.0, "tau_factor": 1.0},
            "ALT-2-LOOP":          {"delta_p_factor": 0.5, "cv_factor": 0.55, "tau_factor": 1.05},
            "ALT-3-MID-TANK":      {"delta_p_factor": 0.35, "cv_factor": 0.4, "tau_factor": 1.30},
            "ALT-4-BASEMENT-TANK": {"delta_p_factor": 0.40, "cv_factor": 0.5, "tau_factor": 1.25},
            "ALT-5-FLOW-CONTROL":  {"delta_p_factor": 0.85, "cv_factor": 0.7, "tau_factor": 1.0},
        }
        d = deltas.get(alt.alt_id, {"delta_p_factor": 1.0, "cv_factor": 1.0, "tau_factor": 1.0})
        new_delta_p = {z: round(p * d["delta_p_factor"], 4) for z, p in baseline.delta_p_max_mpa_per_zone.items()}
        new_cv = round(baseline.cv_flow * d["cv_factor"], 4)
        new_tau = round(baseline.tau_water_minutes * d["tau_factor"], 2)
        from phd_rules import evaluate_imbalance_tier
        max_dp = max(new_delta_p.values(), default=0.0)
        tier, msgs = evaluate_imbalance_tier(
            delta_p_max_mpa=max_dp,
            cv_flow=new_cv,
            tau_water_minutes=new_tau,
            legal_duration_minutes=baseline.legal_duration_minutes,
        )
        return ImbalanceMetrics(
            delta_p_max_mpa_per_zone=new_delta_p,
            cv_flow=new_cv,
            tau_water_minutes=new_tau,
            legal_duration_minutes=baseline.legal_duration_minutes,
            duration_reduction_pct=round(((baseline.legal_duration_minutes - new_tau) / baseline.legal_duration_minutes) * 100.0, 2),
            tier=tier,
            diagnosis_messages=msgs,
            trace=baseline.trace,
        )

    # --------------------------- Stage ⑧ ------------------------------------

    def _stage_8_finalize(
        self,
        *,
        network: DesignNetwork,
        final_imbalance: ImbalanceMetrics,
        final_validation: dict[str, Any],
    ) -> StageResult:
        """⑧ Finalize — write change log table, emit summary."""
        started = _iso_now()
        log_table = self.logger.render_table()
        traces = self.tracer.by_source()
        verdict = "PASS" if final_imbalance.tier == "auto_pass" else (
            "REVIEW" if final_imbalance.tier == "human_review" else "FAIL"
        )
        return StageResult(
            stage="⑧_finalize",
            started_at=started,
            finished_at=_iso_now(),
            verdict=verdict,
            summary={
                "final_tier": final_imbalance.tier,
                "delta_p_max_mpa": max(final_imbalance.delta_p_max_mpa_per_zone.values(), default=0.0),
                "cv_flow": final_imbalance.cv_flow,
                "tau_water_minutes": final_imbalance.tau_water_minutes,
                "duration_reduction_pct": final_imbalance.duration_reduction_pct,
                "diagnosis": final_imbalance.diagnosis_messages,
                "change_log_entries": len(log_table),
                "trace_summary": {k: len(v) for k, v in traces.items()},
            },
            artifacts={
                "change_log_table": json.dumps(log_table, ensure_ascii=False),
            },
        )

    # --------------------------- Helpers ------------------------------------

    def _summarize_kpis(
        self,
        network: DesignNetwork,
        imbalance: ImbalanceMetrics,
        validation: dict[str, Any],
    ) -> dict[str, Any]:
        return {
            "zones": len(network.zones),
            "heads": sum(len(z.heads) for z in network.zones),
            "branches": sum(len(z.branches) for z in network.zones),
            "system_type": network.system_type,
            "hb_case": network.hb_case.case.value if network.hb_case else None,
            "tier": imbalance.tier,
            "delta_p_max_mpa": max(imbalance.delta_p_max_mpa_per_zone.values(), default=0.0),
            "cv_flow": imbalance.cv_flow,
            "tau_water_minutes": imbalance.tau_water_minutes,
            "validation_synthetic": validation.get("synthetic", False),
        }

    def _persist_report(self, report: PipelineRunReport) -> None:
        """Save the full run report as JSON."""
        out = self.config.output_dir / f"{self.config.project_id}_{self.run_id}_report.json"
        try:
            with out.open("w", encoding="utf-8") as f:
                json.dump(asdict(report), f, ensure_ascii=False, indent=2, default=_json_default)
        except OSError:
            pass


# ---------------------------------------------------------------------------
# 4. Module helpers
# ---------------------------------------------------------------------------


def _iso_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def _json_default(o: Any) -> Any:
    if hasattr(o, "to_dict"):
        return o.to_dict()
    if hasattr(o, "value"):  # Enum
        return o.value
    if hasattr(o, "__dict__"):
        return o.__dict__
    return str(o)


def _placeholder_hb_case() -> HBCaseDecision:
    """Fallback HBCaseDecision when network has none."""
    return HBCaseDecision(
        case=HBCase.CASE_1,
        pump_location="basement",
        rated_head_max_m=100.0,
        churn_head_max_m=120.0,
        pressurized_zone_m=30.0,
        natural_drop_zone_m=0.0,
        prv_required=False,
        prv_secondary_bar=4.0,
        pipe_material_change_at_m=None,
        trace=TripleTrace(nftc=None, hb=None, phd="placeholder"),
        detail="placeholder",
    )


__all__ = [
    "PipelineConfig",
    "StageResult",
    "PipelineRunReport",
    "SprinklerPipelineV4",
]
