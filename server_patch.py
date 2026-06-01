"""
server_patch.py — Flask 서버에 v4 라우트만 추가하는 패치 모듈

This module adds v4 pipeline routes to the existing Flask app *without*
modifying any template, static asset, or existing route handler.

Application:
  At the bottom of `대조 서버.py`, add ONE line just before
  `app.run(...)`:

      from server_patch import register_v4_routes
      register_v4_routes(app)

That's it. No HTML/CSS changes. No existing-route changes.

What gets added:
  - GET  /api/v4/status                 — health + module versions
  - POST /api/v4/auto-design            — full pipeline run
  - POST /api/v4/zone-partition         — stage ② only
  - POST /api/v4/head-placement         — stage ③ only
  - POST /api/v4/pipe-routing           — stage ④ only
  - POST /api/v4/scenarios-generate     — stage ⑤ scenarios only
  - POST /api/v4/imbalance-evaluate     — stage ⑥ imbalance metrics only
  - POST /api/v4/alternatives-generate  — stage ⑦ 5 alternatives only
  - POST /api/v4/validate-scenarios     — multi-scenario PIPENET validation
  - GET  /api/v4/change-log/<project>   — change log table
  - GET  /api/v4/trace/<project>        — 3-source trace summary
"""

from __future__ import annotations

import json
import traceback
from dataclasses import asdict
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, request

from nftc_rules import (
    Verdict,
    decide_horizontal_distance,
    decide_temperature_rating,
    decide_reference_count,
    decide_combined_water_supply,
    decide_esfr_branch,
    is_fast_response_required,
    summarize_nftc_decisions,
)
from hb_rules import (
    decide_hb_case,
    decide_pipe_material,
    decide_system_type,
    decide_zone_partition,
    validate_velocity,
    validate_churn_pressure,
)
from phd_rules import (
    classify_pressure_zones,
    decide_discretionary_variables,
    evaluate_imbalance,
    generate_alternative_scenarios,
    generate_calculation_scenarios,
    rank_alternatives,
)
from auto_design import (
    AutoHeadPlacer,
    AutoPipeRouter,
    AutoZonePlanner,
    Zone,
    design_full_network,
)
from change_log import ChangeLogger, TripleTracer
from pipeline_orchestrator import PipelineConfig, SprinklerPipelineV4
from pipenet_validator_v4 import PipenetValidatorV4
from ai_vision import get_cached_ai_vision_extractor
from closed_vocab_ocr import get_cached_closed_vocab_ocr


V4_VERSION = "v4.0.0"
DEFAULT_OUTPUT_DIR = Path("data/v4_outputs")
DEFAULT_LOG_DIR = Path("data/change_logs")


# ---------------------------------------------------------------------------
# Helper — JSON serialization for dataclasses / enums
# ---------------------------------------------------------------------------


def _to_json_safe(obj: Any) -> Any:
    if obj is None or isinstance(obj, (str, int, float, bool)):
        return obj
    if isinstance(obj, dict):
        return {str(k): _to_json_safe(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [_to_json_safe(v) for v in obj]
    if hasattr(obj, "to_dict"):
        return _to_json_safe(obj.to_dict())
    if hasattr(obj, "value") and hasattr(obj, "name"):  # Enum
        return obj.value
    if hasattr(obj, "__dataclass_fields__"):
        return _to_json_safe(asdict(obj))
    if hasattr(obj, "__dict__"):
        return _to_json_safe(obj.__dict__)
    return str(obj)


def _err(message: str, status: int = 400) -> tuple:
    return (jsonify({"ok": False, "error": message}), status)


def _ok(payload: Any) -> Any:
    return jsonify({"ok": True, "data": _to_json_safe(payload)})


# ---------------------------------------------------------------------------
# Route registrar
# ---------------------------------------------------------------------------


def register_v4_routes(app: Flask) -> Flask:
    """Register every v4 route on the given Flask app and return it."""

    # -----------------------------------------------------------------------
    # 0. Status
    # -----------------------------------------------------------------------

    @app.get("/api/v4/status")
    def v4_status():
        return _ok({
            "version": V4_VERSION,
            "modules": {
                "nftc_rules": "loaded",
                "hb_rules": "loaded",
                "phd_rules": "loaded",
                "auto_design": "loaded",
                "change_log": "loaded",
                "pipeline_orchestrator": "loaded",
                "pipenet_validator_v4": "loaded",
                "ai_vision": "loaded",
                "closed_vocab_ocr": "loaded" if (Path(__file__).resolve().parent / "models" / "closed_vocab_ocr" / "model.pt").exists() else "missing",
            },
            "endpoints": [
                "/api/v4/status",
                "/api/v4/auto-design",
                "/api/v4/ai/extract-from-dxf",
                "/api/v4/ai/ocr-closed-vocab",
                "/api/v4/zone-partition",
                "/api/v4/head-placement",
                "/api/v4/pipe-routing",
                "/api/v4/scenarios-generate",
                "/api/v4/imbalance-evaluate",
                "/api/v4/alternatives-generate",
                "/api/v4/validate-scenarios",
                "/api/v4/change-log/<project>",
                "/api/v4/trace/<project>",
                "/api/v4/nftc/reference-count",
                "/api/v4/nftc/horizontal-distance",
                "/api/v4/nftc/temperature-rating",
                "/api/v4/nftc/esfr-branch",
                "/api/v4/nftc/combined-supply",
                "/api/v4/hb/case",
                "/api/v4/hb/system-type",
                "/api/v4/hb/pipe-material",
            ],
        })

    # -----------------------------------------------------------------------
    # 1. Full pipeline
    # -----------------------------------------------------------------------

    @app.post("/api/v4/auto-design")
    def v4_auto_design():
        """Run the full v4 pipeline.

        JSON body:
          {
            "project_id": "...",
            "dxf_path": "data/uploads/site.dxf",
            "building_meta": {...},
            "rooms": [...],
            "obstacles": [...],
            "floors": [{"label": "F1", "z_m": 0.0}, ...],
            "existing_pipenet_pdf_path": "...",   # optional
            "max_redesign_attempts": 3            # optional
          }
        """
        try:
            body = request.get_json(force=True) or {}
            project_id = body.get("project_id") or "v4-default"
            dxf_path = Path(body.get("dxf_path") or "")
            building_meta = body.get("building_meta") or {}
            rooms = body.get("rooms") or []
            obstacles = body.get("obstacles") or []
            floors = body.get("floors") or []
            existing_pdf = body.get("existing_pipenet_pdf_path")
            cfg = PipelineConfig(
                project_id=project_id,
                max_redesign_attempts=int(body.get("max_redesign_attempts", 3)),
                output_dir=DEFAULT_OUTPUT_DIR,
                log_dir=DEFAULT_LOG_DIR,
            )
            pipe = SprinklerPipelineV4(cfg)
            report = pipe.run_full_pipeline(
                dxf_path=dxf_path,
                building_meta=building_meta,
                rooms=rooms,
                obstacles=obstacles,
                floors=floors,
                existing_pipenet_pdf_path=Path(existing_pdf) if existing_pdf else None,
            )
            return _ok(report)
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/ai/extract-from-dxf")
    def v4_ai_extract_from_dxf():
        """Run the trained triangle-head detector on a DXF file."""
        try:
            body = request.get_json(force=True) or {}
            dxf_path = Path(body.get("dxf_path") or "")
            if not dxf_path.exists():
                return _err(f"DXF path not found: {dxf_path}", status=404)
            from cad_engine import DXFWorkspace  # type: ignore[import-not-found]

            workspace = DXFWorkspace(dxf_path.parent / "_v4_workspace")
            workspace.load_file(dxf_path)
            payload = workspace.to_payload(include_network_entities=True, include_network_summary=True)
            extractor = get_cached_ai_vision_extractor(str(Path(__file__).resolve().parent))
            result = extractor.enhance_from_payload(
                dxf_path=dxf_path,
                cad_payload=payload,
            )
            return _ok(result)
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/ai/ocr-closed-vocab")
    def v4_ai_ocr_closed_vocab():
        """Recognize a label image against a fixed vocabulary."""
        try:
            body = request.get_json(force=True) or {}
            image_path = Path(body.get("image_path") or "")
            candidates = body.get("candidates")
            if not image_path.exists():
                return _err(f"Image path not found: {image_path}", status=404)
            ocr = get_cached_closed_vocab_ocr(str(Path(__file__).resolve().parent))
            result = ocr.predict(image_path=image_path, candidates=candidates if isinstance(candidates, list) else None)
            return _ok(result)
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    # -----------------------------------------------------------------------
    # 2. Stage-by-stage routes (for incremental UI / debugging)
    # -----------------------------------------------------------------------

    @app.post("/api/v4/zone-partition")
    def v4_zone_partition():
        """Stage ② — partition floors into zones."""
        try:
            body = request.get_json(force=True) or {}
            zones = decide_zone_partition(
                floor_area_m2=float(body.get("floor_area_m2", 0)),
                estimated_head_count=int(body.get("estimated_head_count", 0)),
                floor_label=str(body.get("floor_label", "")),
                is_grid_layout=bool(body.get("is_grid_layout", False)),
                is_apartment_loft=bool(body.get("is_apartment_loft", False)),
                fire_compartment_id=body.get("fire_compartment_id"),
                system_type=str(body.get("system_type", "wet")),
            )
            return _ok({"zones": zones})
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/head-placement")
    def v4_head_placement():
        """Stage ③ — auto-place heads in a zone.

        JSON body:
          {
            "zone": {...Zone fields...},
            "obstacles": [...]
          }
        """
        try:
            body = request.get_json(force=True) or {}
            zone_data = body.get("zone") or {}
            obstacles = body.get("obstacles") or []
            # Build minimal Zone from data (caller is responsible for head_spec)
            from auto_design import HeadSpec, Zone as ZoneCls
            spec_data = zone_data.get("head_spec") or {}
            head_spec = HeadSpec(
                zone_id=zone_data.get("zone_id", "Z0"),
                horizontal_distance_m=float(spec_data.get("horizontal_distance_m", 2.3)),
                k_factor_lpm_bar05=int(spec_data.get("k_factor_lpm_bar05", 80)),
                temperature_rating_min_c=float(spec_data.get("temperature_rating_min_c", 79.0)),
                temperature_rating_max_c=spec_data.get("temperature_rating_max_c"),
                rti_class=str(spec_data.get("rti_class", "standard")),
                head_type=str(spec_data.get("head_type", "pendent")),
                corrosion_resistant=bool(spec_data.get("corrosion_resistant", False)),
                is_esfr=bool(spec_data.get("is_esfr", False)),
                trace=__import__("nftc_rules").TripleTrace(),
            )
            zone = ZoneCls(
                zone_id=zone_data.get("zone_id", "Z0"),
                floor_label=zone_data.get("floor_label", "F1"),
                polygon=[tuple(pt) for pt in (zone_data.get("polygon") or [])],
                area_m2=float(zone_data.get("area_m2", 0)),
                use=str(zone_data.get("use", "other_low")),
                structure=str(zone_data.get("structure", "non_fire_resistant")),
                ceiling_h_m=float(zone_data.get("ceiling_h_m", 3.0)),
                head_spec=head_spec,
            )
            placer = AutoHeadPlacer(zone=zone, obstacles=obstacles)
            heads = placer.place()
            return _ok({"heads": heads, "head_count": len(heads)})
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}", status=500)

    @app.post("/api/v4/pipe-routing")
    def v4_pipe_routing():
        """Stage ④ — route branches and cross-mains for a zone.

        JSON body:
          {
            "zone": {...},
            "heads": [...],
            "riser_xy": [x, y]
          }
        """
        try:
            body = request.get_json(force=True) or {}
            from auto_design import HeadInstance, HeadSpec, Zone as ZoneCls
            zone_data = body.get("zone") or {}
            spec_data = zone_data.get("head_spec") or {}
            head_spec = HeadSpec(
                zone_id=zone_data.get("zone_id", "Z0"),
                horizontal_distance_m=float(spec_data.get("horizontal_distance_m", 2.3)),
                k_factor_lpm_bar05=int(spec_data.get("k_factor_lpm_bar05", 80)),
                temperature_rating_min_c=float(spec_data.get("temperature_rating_min_c", 79.0)),
                temperature_rating_max_c=spec_data.get("temperature_rating_max_c"),
                rti_class=str(spec_data.get("rti_class", "standard")),
                head_type=str(spec_data.get("head_type", "pendent")),
                corrosion_resistant=bool(spec_data.get("corrosion_resistant", False)),
                is_esfr=bool(spec_data.get("is_esfr", False)),
                trace=__import__("nftc_rules").TripleTrace(),
            )
            zone = ZoneCls(
                zone_id=zone_data.get("zone_id", "Z0"),
                floor_label=zone_data.get("floor_label", "F1"),
                polygon=[tuple(pt) for pt in (zone_data.get("polygon") or [])],
                area_m2=float(zone_data.get("area_m2", 0)),
                use=str(zone_data.get("use", "other_low")),
                structure=str(zone_data.get("structure", "non_fire_resistant")),
                ceiling_h_m=float(zone_data.get("ceiling_h_m", 3.0)),
                head_spec=head_spec,
            )
            for h in body.get("heads") or []:
                zone.heads.append(HeadInstance(
                    head_id=h.get("head_id", "H?"),
                    zone_id=zone.zone_id,
                    x=float(h.get("x", 0)),
                    y=float(h.get("y", 0)),
                    z=float(h.get("z", 0)),
                    spec=head_spec,
                    branch_axis=str(h.get("branch_axis", "EW")),
                    cell_S=float(h.get("cell_S", 3.0)),
                    cell_L=float(h.get("cell_L", 3.0)),
                    nftc_2773_pass=True,
                    nftc_2771_pass=True,
                    skipping_pass=True,
                    a_b_l_check=True,
                ))
            riser = body.get("riser_xy")
            riser_xy = tuple(riser) if riser else None
            router = AutoPipeRouter(zone=zone, riser_xy=riser_xy)
            branches, cross_mains = router.route()
            return _ok({
                "branches": branches,
                "cross_mains": cross_mains,
                "branch_count": len(branches),
                "cross_main_count": len(cross_mains),
            })
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}", status=500)

    @app.post("/api/v4/scenarios-generate")
    def v4_scenarios_generate():
        """Stage ⑤ — generate up to 12 PIPENET calculation scenarios.

        JSON body:
          {
            "floors": [...FloorPressureZone-like...],
            "hb_case": {...},
            "elevated_tank_z_m": 30.0,
            "pump_rated_q_lpm": 2400, "pump_rated_h_m": 60, "pump_churn_h_m": 70,
            "has_k115_zones": false, "has_max_q_zone": true
          }
        """
        try:
            from phd_rules import FloorPressureZone, PressureZone
            body = request.get_json(force=True) or {}
            floors = []
            for f in body.get("floors") or []:
                zone_value = f.get("zone", "msp")
                floors.append(FloorPressureZone(
                    floor_label=f.get("floor_label", "?"),
                    elevation_m=float(f.get("elevation_m", 0)),
                    zone=PressureZone(zone_value),
                    natural_drop_pressure_bar=float(f.get("natural_drop_pressure_bar", 0)),
                    requires_prv=bool(f.get("requires_prv", False)),
                ))
            from hb_rules import HBCase, HBCaseDecision
            hb_data = body.get("hb_case") or {}
            from nftc_rules import TripleTrace
            hb_case = HBCaseDecision(
                case=HBCase(hb_data.get("case", "case_1")),
                pump_location=hb_data.get("pump_location", "basement"),
                rated_head_max_m=float(hb_data.get("rated_head_max_m", 100)),
                churn_head_max_m=float(hb_data.get("churn_head_max_m", 120)),
                pressurized_zone_m=float(hb_data.get("pressurized_zone_m", 0)),
                natural_drop_zone_m=float(hb_data.get("natural_drop_zone_m", 0)),
                prv_required=bool(hb_data.get("prv_required", False)),
                prv_secondary_bar=float(hb_data.get("prv_secondary_bar", 4.0)),
                pipe_material_change_at_m=hb_data.get("pipe_material_change_at_m"),
                trace=TripleTrace(hb=hb_data.get("trace_hb")),
                detail=hb_data.get("detail", ""),
            )
            discretionary = decide_discretionary_variables(
                floors=floors,
                hb_case=hb_case,
                elevated_tank_z_m=float(body.get("elevated_tank_z_m", 30)),
                pump_rated_q_lpm=float(body.get("pump_rated_q_lpm", 2400)),
                pump_rated_h_m=float(body.get("pump_rated_h_m", 60)),
                pump_churn_h_m=float(body.get("pump_churn_h_m", 70)),
            )
            scenarios = generate_calculation_scenarios(
                discretionary=discretionary,
                has_k115_zones=bool(body.get("has_k115_zones", False)),
                has_max_q_zone=bool(body.get("has_max_q_zone", True)),
            )
            return _ok({
                "scenarios": scenarios,
                "scenario_count": len(scenarios),
                "discretionary": discretionary,
            })
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}", status=500)

    @app.post("/api/v4/imbalance-evaluate")
    def v4_imbalance_evaluate():
        """Stage ⑥.5 — evaluate 3 imbalance metrics.

        JSON body:
          {
            "head_flows_lpm": [80, 82, ...],
            "zone_pressures": {"Z1": [0.1, 0.2, ...], ...},
            "tank_total_volume_m3": 40.0,
            "legal_duration_minutes": 20.0
          }
        """
        try:
            body = request.get_json(force=True) or {}
            metrics = evaluate_imbalance(
                head_flows_lpm=[float(x) for x in (body.get("head_flows_lpm") or [])],
                zone_pressures={
                    str(k): [float(p) for p in (v or [])]
                    for k, v in (body.get("zone_pressures") or {}).items()
                },
                tank_total_volume_m3=float(body.get("tank_total_volume_m3", 40.0)),
                legal_duration_minutes=float(body.get("legal_duration_minutes", 20.0)),
            )
            return _ok(metrics)
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/alternatives-generate")
    def v4_alternatives_generate():
        """Stage ⑦ — generate 5 design alternatives + ranking.

        JSON body:
          {
            "current_imbalance": {...ImbalanceMetrics-like...},
            "hb_case": {...},
            "simulation_results": { "ALT-1-PRV": {...}, ... }   # optional
          }
        """
        try:
            from phd_rules import ImbalanceMetrics
            from nftc_rules import TripleTrace
            from hb_rules import HBCase, HBCaseDecision
            body = request.get_json(force=True) or {}
            ci = body.get("current_imbalance") or {}
            current = ImbalanceMetrics(
                delta_p_max_mpa_per_zone=ci.get("delta_p_max_mpa_per_zone", {}),
                cv_flow=float(ci.get("cv_flow", 0)),
                tau_water_minutes=float(ci.get("tau_water_minutes", 0)),
                legal_duration_minutes=float(ci.get("legal_duration_minutes", 20.0)),
                duration_reduction_pct=float(ci.get("duration_reduction_pct", 0)),
                tier=str(ci.get("tier", "human_review")),
                diagnosis_messages=list(ci.get("diagnosis_messages") or []),
                trace=TripleTrace(),
            )
            hb_data = body.get("hb_case") or {}
            hb_case = HBCaseDecision(
                case=HBCase(hb_data.get("case", "case_1")),
                pump_location=hb_data.get("pump_location", "basement"),
                rated_head_max_m=float(hb_data.get("rated_head_max_m", 100)),
                churn_head_max_m=float(hb_data.get("churn_head_max_m", 120)),
                pressurized_zone_m=float(hb_data.get("pressurized_zone_m", 0)),
                natural_drop_zone_m=float(hb_data.get("natural_drop_zone_m", 0)),
                prv_required=bool(hb_data.get("prv_required", False)),
                prv_secondary_bar=float(hb_data.get("prv_secondary_bar", 4.0)),
                pipe_material_change_at_m=hb_data.get("pipe_material_change_at_m"),
                trace=TripleTrace(),
                detail=hb_data.get("detail", ""),
            )
            alts = generate_alternative_scenarios(
                diagnosis=current,
                hb_case=hb_case,
                has_basement=bool(body.get("has_basement", True)),
            )
            sim_in = body.get("simulation_results") or {}
            sim_results: dict[str, ImbalanceMetrics] = {}
            for aid, sim in sim_in.items():
                sim_results[aid] = ImbalanceMetrics(
                    delta_p_max_mpa_per_zone=sim.get("delta_p_max_mpa_per_zone", {}),
                    cv_flow=float(sim.get("cv_flow", 0)),
                    tau_water_minutes=float(sim.get("tau_water_minutes", 0)),
                    legal_duration_minutes=float(sim.get("legal_duration_minutes", 20.0)),
                    duration_reduction_pct=float(sim.get("duration_reduction_pct", 0)),
                    tier=str(sim.get("tier", "human_review")),
                    diagnosis_messages=list(sim.get("diagnosis_messages") or []),
                    trace=TripleTrace(),
                )
            ranked = rank_alternatives(alts, simulation_results=sim_results) if sim_results else []
            return _ok({
                "alternatives": alts,
                "ranked": ranked,
                "alternative_count": len(alts),
            })
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}", status=500)

    @app.post("/api/v4/validate-scenarios")
    def v4_validate_scenarios():
        """Validate 12 scenarios against 6 hard rules + 3 imbalance metrics."""
        try:
            from phd_rules import CalculationScenario
            body = request.get_json(force=True) or {}
            project_id = body.get("project_id") or "v4-validate"
            scenarios = []
            for s in body.get("scenarios") or []:
                scenarios.append(CalculationScenario(
                    scenario_id=s.get("scenario_id", "S?"),
                    zone=s.get("zone", "?"),
                    position=s.get("position", "?"),
                    floor=s.get("floor", "?"),
                    purpose=s.get("purpose", ""),
                    priority=s.get("priority", ""),
                    config_overrides=s.get("config_overrides") or {},
                ))
            sdf_paths = {k: Path(v) for k, v in (body.get("sdf_paths_by_scenario") or {}).items()}
            pdf_paths = {k: Path(v) for k, v in (body.get("pipenet_pdf_paths_by_scenario") or {}).items()}
            cad_path = Path(body["cad_path"]) if body.get("cad_path") else None
            v4 = PipenetValidatorV4(
                project_id=project_id,
                tank_total_volume_m3=float(body.get("tank_total_volume_m3", 40.0)),
                legal_duration_minutes=float(body.get("legal_duration_minutes", 20.0)),
            )
            report = v4.validate_scenarios(
                scenarios=scenarios,
                sdf_paths_by_scenario=sdf_paths,
                pipenet_pdf_paths_by_scenario=pdf_paths,
                cad_path=cad_path,
            )
            return _ok(report)
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}\n{traceback.format_exc()}", status=500)

    # -----------------------------------------------------------------------
    # 3. Audit endpoints — change log + trace
    # -----------------------------------------------------------------------

    @app.get("/api/v4/change-log/<project>")
    def v4_change_log(project: str):
        """Read the change log table for a project."""
        try:
            logger = ChangeLogger(project, log_dir=DEFAULT_LOG_DIR)
            return _ok({"rows": logger.render_table()})
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.get("/api/v4/trace/<project>")
    def v4_trace(project: str):
        """Read the most recent run report for a project (3-source trace)."""
        try:
            files = sorted(DEFAULT_OUTPUT_DIR.glob(f"{project}_*_report.json"))
            if not files:
                return _err("no run report found", status=404)
            latest = files[-1]
            with latest.open(encoding="utf-8") as f:
                report = json.load(f)
            return _ok({
                "report": report,
                "trace_summary": report.get("trace_summary"),
                "stages": [s.get("stage") for s in (report.get("stages") or [])],
            })
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    # -----------------------------------------------------------------------
    # 4. Atomic NFTC / HB rule queries (for UI quick-checks)
    # -----------------------------------------------------------------------

    @app.post("/api/v4/nftc/reference-count")
    def v4_nftc_reference_count():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_reference_count(body))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/nftc/horizontal-distance")
    def v4_nftc_horizontal_distance():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_horizontal_distance(
                room_use=str(body.get("room_use", "other_low")),
                structure=str(body.get("structure", "non_fire_resistant")),
                has_special_combustible=bool(body.get("has_special_combustible", False)),
                is_rack_storage=bool(body.get("is_rack_storage", False)),
            ))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/nftc/temperature-rating")
    def v4_nftc_temperature_rating():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_temperature_rating(
                ambient_temp_c=float(body.get("ambient_temp_c", 25.0)),
                is_factory_4m_high=bool(body.get("is_factory_4m_high", False)),
                is_warehouse_4m_high=bool(body.get("is_warehouse_4m_high", False)),
                is_rack_storage=bool(body.get("is_rack_storage", False)),
            ))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/nftc/esfr-branch")
    def v4_nftc_esfr_branch():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_esfr_branch(
                room_use=str(body.get("room_use", "other_low")),
                ceiling_h_m=float(body.get("ceiling_h_m", 3.0)),
            ))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/nftc/combined-supply")
    def v4_nftc_combined_supply():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_combined_water_supply(
                body.get("systems") or {},
                use_combined_tank=bool(body.get("use_combined_tank", True)),
                use_combined_pump=bool(body.get("use_combined_pump", True)),
            ))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/hb/case")
    def v4_hb_case():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_hb_case(
                building_height_m=float(body.get("building_height_m", 30.0)),
                refuge_floor_interval_m=body.get("refuge_floor_interval_m"),
                rooftop_tank_feasible=bool(body.get("rooftop_tank_feasible", True)),
                water_source_type=str(body.get("water_source_type", "fire_dedicated")),
            ))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/hb/system-type")
    def v4_hb_system_type():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_system_type(
                has_freezing_risk=bool(body.get("has_freezing_risk", False)),
                needs_open_heads=bool(body.get("needs_open_heads", False)),
                detector_priority=bool(body.get("detector_priority", False)),
                room_use=str(body.get("room_use", "")),
            ))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    @app.post("/api/v4/hb/pipe-material")
    def v4_hb_pipe_material():
        try:
            body = request.get_json(force=True) or {}
            return _ok(decide_pipe_material(float(body.get("operating_pressure_mpa", 1.0))))
        except Exception as exc:
            return _err(f"{type(exc).__name__}: {exc}", status=500)

    return app


__all__ = ["register_v4_routes", "V4_VERSION"]
