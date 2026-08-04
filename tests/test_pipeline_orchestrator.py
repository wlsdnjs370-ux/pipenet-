# -*- coding: utf-8 -*-
"""core/pipeline_orchestrator.py — v4 파이프라인 헬퍼 · 증거 수집 테스트.

run_full_pipeline 자체는 DXF·PIPENET 결과 PDF 를 요구해 단위 테스트 대상이
아니다. 대신 "계산한 적 없는 설계가 합격 도장을 다는" 경로를 막는 헬퍼
(_gather_*, _insufficient_evidence_metrics, _max_delta_p)와 보고서 직렬화를
고정한다.

실행::

    python -m pytest tests/test_pipeline_orchestrator.py -q
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

import pipeline_orchestrator as O  # noqa: E402
from auto_design import DesignNetwork  # noqa: E402
from nftc_rules import TripleTrace, Verdict  # noqa: E402


@pytest.fixture
def pipe(tmp_path):
    cfg = O.PipelineConfig(
        project_id="T-001",
        output_dir=tmp_path / "out",
        log_dir=tmp_path / "log",
    )
    return O.SprinklerPipelineV4(cfg)


# ────────────────────────────────────────────────────────────────────────────
# 설정 · 생성
# ────────────────────────────────────────────────────────────────────────────

def test_config_defaults_follow_nftc_and_hb():
    cfg = O.PipelineConfig(project_id="p")
    assert cfg.legal_duration_minutes == 20.0    # NFTC 2.9.3.2
    assert cfg.max_redesign_attempts == 3
    assert cfg.require_human_signoff is True


def test_init_creates_output_dirs(tmp_path, pipe):
    assert (tmp_path / "out").is_dir()
    assert (tmp_path / "log").is_dir()
    assert pipe.run_id.startswith("run-")


def test_run_ids_are_unique_within_the_same_second():
    """같은 초에 두 run 이 시작해도 보고서 파일이 겹치면 안 된다."""
    ids = {O._new_run_id() for _ in range(50)}
    assert len(ids) == 50


def test_iso_now_is_timezone_aware():
    dt = datetime.fromisoformat(O._iso_now())
    assert dt.tzinfo is not None


# ────────────────────────────────────────────────────────────────────────────
# ① NFTC 2.12 / HB §2.4.18 제외
# ────────────────────────────────────────────────────────────────────────────

@pytest.mark.parametrize("room, reason", [
    ({"use": "stairwell"}, "①_stairwell"),
    ({"use": "telecom_room"}, "②_telecom"),
    ({"use": "generator_room"}, "③_generator"),
    ({"use": "surgery_room"}, "④_surgery"),
    ({"use": "pump_room"}, "⑧_pump_water"),
    ({"use": "storage", "ambient_temp_c": -5.0}, "⑪_freezer"),
    ({"use": "storage", "roof_type": "open_to_outside"}, "⑨_open_to_outside"),
    ({"use": "outdoor"}, "⑩_outdoor"),
])
def test_exclusion_categories(pipe, room, reason):
    out = pipe._apply_24_18_exclusion([dict(room)])[0]
    assert out["hb_excluded"] is True
    assert out["exclusion_reason"] == reason


def test_ordinary_room_is_not_excluded(pipe):
    out = pipe._apply_24_18_exclusion([{"use": "office", "ambient_temp_c": 25.0}])[0]
    assert "hb_excluded" not in out


def test_freezer_threshold_is_minus_3c(pipe):
    warm = pipe._apply_24_18_exclusion([{"use": "storage", "ambient_temp_c": -2.9}])[0]
    cold = pipe._apply_24_18_exclusion([{"use": "storage", "ambient_temp_c": -3.0}])[0]
    assert "hb_excluded" not in warm
    assert cold["hb_excluded"] is True


# ────────────────────────────────────────────────────────────────────────────
# ⑥ 증거 수집 — 없으면 None
# ────────────────────────────────────────────────────────────────────────────

def test_head_flows_is_none_without_table(pipe):
    """표가 없을 때 80 LPM 을 채우면 CV=0 → auto_pass 라는 거짓 만점이 나온다."""
    assert pipe._gather_head_flows({}) is None
    assert pipe._gather_head_flows({"tables": {"nozzle_flows": []}}) is None


def test_head_flows_reads_actual_values(pipe):
    flows = pipe._gather_head_flows({"tables": {"nozzle_flows": [
        {"actual_flow_lpm": 81.2}, {"actual_flow_lpm": 95}, {"actual_flow_lpm": None}]}})
    assert flows == [81.2, 95.0]


def test_zone_pressures_is_none_without_table(pipe):
    assert pipe._gather_zone_pressures({}) is None
    assert pipe._gather_zone_pressures({"tables": {"pipe_validation_rows": []}}) is None


def test_zone_pressures_convert_kgcm2_to_mpa(pipe):
    p = pipe._gather_zone_pressures({"tables": {"pipe_validation_rows": [
        {"zone_id": "Z1", "inlet_pressure_kgcm2": 10.197,
         "outlet_pressure_kgcm2": 5.0985, "max_pressure_kgcm2": None}]}})
    assert p["Z1"][0] == pytest.approx(1.0)
    assert p["Z1"][1] == pytest.approx(0.5)
    assert len(p["Z1"]) == 2      # None 은 채우지 않는다


def test_zone_pressures_group_unlabeled_rows_as_unknown(pipe):
    p = pipe._gather_zone_pressures({"tables": {"pipe_validation_rows": [
        {"inlet_pressure_kgcm2": 4.0}]}})
    assert list(p) == ["unknown"]


def test_tank_volume_is_reference_count_times_two(pipe):
    net = DesignNetwork(project_id="p", building_height_m=30.0, hb_case=None,
                        system_type="wet", metadata={"reference_count": 30})
    assert pipe._estimate_tank_volume(net) == 60.0


def test_tank_volume_falls_back_to_20_references(pipe):
    net = DesignNetwork(project_id="p", building_height_m=30.0, hb_case=None,
                        system_type="wet")
    assert pipe._estimate_tank_volume(net) == 40.0


# ────────────────────────────────────────────────────────────────────────────
# 근거 미확보 판정
# ────────────────────────────────────────────────────────────────────────────

def test_insufficient_evidence_keeps_metrics_none():
    m = O._insufficient_evidence_metrics(
        legal_duration_minutes=20.0, missing=["헤드 유량표"])
    assert m.tier == O.TIER_INSUFFICIENT_EVIDENCE
    assert m.cv_flow is None
    assert m.tau_water_minutes is None
    assert m.duration_reduction_pct is None
    assert m.delta_p_max_mpa_per_zone == {}
    assert "헤드 유량표" in m.diagnosis_messages[0]


def test_insufficient_evidence_maps_to_review_not_pass():
    """근거가 없어 경고가 안 뜬 것을 '지표 모두 통과'로 보고하면 안 된다."""
    assert O._TIER_VERDICT[O.TIER_INSUFFICIENT_EVIDENCE] == "REVIEW"
    assert O._TIER_VERDICT["auto_pass"] == "PASS"
    assert O._TIER_VERDICT["redesign_required"] == "FAIL"


def test_max_delta_p_is_none_when_no_zone():
    m = O._insufficient_evidence_metrics(legal_duration_minutes=20.0, missing=["x"])
    assert O._max_delta_p(m) is None


def test_max_delta_p_picks_worst_zone():
    m = O._insufficient_evidence_metrics(legal_duration_minutes=20.0, missing=["x"])
    m.delta_p_max_mpa_per_zone = {"a": 0.2, "b": 0.7}
    assert O._max_delta_p(m) == 0.7


def test_placeholder_hb_case_is_marked_as_placeholder():
    c = O._placeholder_hb_case()
    assert c.detail == "placeholder"
    assert c.trace.phd == "placeholder"
    assert c.prv_required is False


# ────────────────────────────────────────────────────────────────────────────
# ⑥ 하드룰 폴백
# ────────────────────────────────────────────────────────────────────────────

def _net_with_pump(rated, churn):
    return DesignNetwork(
        project_id="p", building_height_m=30.0, hb_case=None, system_type="wet",
        pumps=[{"rated_h_m": rated, "churn_h_m": churn}])


def test_fallback_is_flagged_synthetic(pipe):
    out = pipe._fallback_hard_rule_check(_net_with_pump(100.0, 110.0))
    assert out["synthetic"] is True
    assert any("PIPENET" in w for w in out["results"]["WARNING"])


@pytest.mark.parametrize("churn, bucket, ratio", [
    (110.0, "PASS", "1.100"),
    (130.0, "WARNING", "1.300"),   # HB 초과 · NFTC 이내
    (150.0, "FAIL", "1.500"),
])
def test_fallback_routes_churn_verdict_to_bucket(pipe, churn, bucket, ratio):
    out = pipe._fallback_hard_rule_check(_net_with_pump(100.0, churn))
    assert any(ratio in msg for msg in out["results"][bucket])


def test_fallback_without_pump_skips_churn(pipe):
    out = pipe._fallback_hard_rule_check(
        DesignNetwork(project_id="p", building_height_m=30.0, hb_case=None,
                      system_type="wet"))
    assert out["results"]["FAIL"] == []


# ────────────────────────────────────────────────────────────────────────────
# KPI 요약 · 보고서 직렬화
# ────────────────────────────────────────────────────────────────────────────

def test_kpis_report_insufficient_evidence_as_null_not_zero(pipe):
    m = O._insufficient_evidence_metrics(legal_duration_minutes=20.0, missing=["x"])
    kpi = pipe._summarize_kpis(
        DesignNetwork(project_id="p", building_height_m=30.0, hb_case=None,
                      system_type="wet"),
        m, {"synthetic": True})
    assert kpi["tier"] == O.TIER_INSUFFICIENT_EVIDENCE
    assert kpi["delta_p_max_mpa"] is None
    assert kpi["cv_flow"] is None
    assert kpi["validation_synthetic"] is True
    assert kpi["zones"] == 0


def test_json_default_unwraps_enum_and_trace():
    assert O._json_default(Verdict.PASS) == "PASS"
    assert O._json_default(TripleTrace(nftc="x"))["NFTC"] == "x"
    assert isinstance(O._json_default(object()), str)


def test_persist_report_writes_readable_json(pipe, tmp_path):
    report = O.PipelineRunReport(
        project_id="T-001", run_id=pipe.run_id, started_at=O._iso_now(),
        overall_verdict="REVIEW",
        stages=[O.StageResult(stage="①", started_at=O._iso_now(),
                              finished_at=O._iso_now(), verdict="PASS",
                              summary={"rooms": 3})],
        final_kpis={"tier": O.TIER_INSUFFICIENT_EVIDENCE, "delta_p_max_mpa": None},
    )
    pipe._persist_report(report)
    files = list((tmp_path / "out").glob("*_report.json"))
    assert len(files) == 1
    data = json.loads(files[0].read_text(encoding="utf-8"))
    assert data["overall_verdict"] == "REVIEW"
    assert data["final_kpis"]["delta_p_max_mpa"] is None
    assert data["stages"][0]["summary"]["rooms"] == 3


def test_persist_report_survives_unwritable_dir(pipe):
    """보고서 저장 실패가 파이프라인 전체를 죽이면 안 된다."""
    pipe.config.output_dir = Path("Z:/does/not/exist")
    pipe._persist_report(O.PipelineRunReport(
        project_id="T-001", run_id="r", started_at=O._iso_now()))


def test_all_exports_exist():
    for name in O.__all__:
        assert hasattr(O, name), name
