from __future__ import annotations

import importlib.util
from pathlib import Path

import pytest


MODULE_PATH = (
    Path(__file__).resolve().parents[1]
    / "sprinkler_ai_agent_server_source_2026-04-27"
    / "extracted"
    / "nftc_rules.py"
)
SPEC = importlib.util.spec_from_file_location("design_automation_nftc_rules", MODULE_PATH)
assert SPEC and SPEC.loader
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


def test_decide_standard_count_special_factory_rule_a() -> None:
    count, rule = MODULE.decide_standard_count("factory_special", 10, 6.0, True)
    assert (count, rule) == (30, "RULE-NFTC-211-A")


def test_decide_standard_count_high_head_rule_e() -> None:
    count, rule = MODULE.decide_standard_count("etc", 8, 8.0, False)
    assert (count, rule) == (20, "RULE-NFTC-211-E")


def test_decide_standard_count_apartment_with_parking_rule_608() -> None:
    count, rule = MODULE.decide_standard_count("apartment_with_basement_parking", 25, 3.0, False)
    assert (count, rule) == (30, "RULE-NFTC608-7-1")


def test_decide_standard_count_rejects_negative_floor() -> None:
    with pytest.raises(ValueError):
        MODULE.decide_standard_count("commercial", -1, 3.0, False)


def test_decide_horizontal_distance_stage() -> None:
    distance, clause = MODULE.decide_horizontal_distance("stage")
    assert distance == 1.7
    assert clause == "NFTC 2.7.3.1"


def test_decide_horizontal_distance_fire_resistant() -> None:
    distance, clause = MODULE.decide_horizontal_distance("fire_resistant")
    assert distance == 2.3
    assert clause == "NFTC 2.7.3.4"


def test_decide_horizontal_distance_invalid() -> None:
    with pytest.raises(ValueError):
        MODULE.decide_horizontal_distance("unknown")


def test_is_quick_response_required_hospital() -> None:
    required, clause = MODULE.is_quick_response_required("hospital_ward")
    assert required is True
    assert "2024.4.1" in clause


def test_is_quick_response_required_false_for_office() -> None:
    required, clause = MODULE.is_quick_response_required("office")
    assert required is False
    assert clause == ""


def test_is_quick_response_required_apartment_living() -> None:
    required, clause = MODULE.is_quick_response_required("apartment_living")
    assert (required, clause) == (True, "NFTC 2.7.5(1)")


def test_decide_temperature_rating_low_band() -> None:
    rating, rule = MODULE.decide_temperature_rating(25.0, False)
    assert rating == 79
    assert "<39C" in rule


def test_decide_temperature_rating_factory_proviso() -> None:
    rating, rule = MODULE.decide_temperature_rating(20.0, True)
    assert rating == 121
    assert "proviso" in rule


def test_decide_temperature_rating_high_band() -> None:
    rating, rule = MODULE.decide_temperature_rating(120.0, False)
    assert rating == 191
    assert ">=106C" in rule


def test_decide_temperature_rating_rejects_negative() -> None:
    with pytest.raises(ValueError):
        MODULE.decide_temperature_rating(-1.0, False)


def test_check_head_clearance_60cm_passes_for_distant_obstacle() -> None:
    passed, reasons = MODULE.check_head_clearance_60cm(
        (0.0, 0.0),
        [{"type": "beam", "xy_polygon": [(1.0, 1.0), (1.4, 1.0), (1.4, 1.3), (1.0, 1.3)]}],
    )
    assert passed is True
    assert reasons == []


def test_check_head_clearance_60cm_fails_for_close_obstacle() -> None:
    passed, reasons = MODULE.check_head_clearance_60cm(
        (0.0, 0.0),
        [{"type": "duct", "xy_polygon": [(-0.2, -0.1), (0.2, -0.1), (0.2, 0.1), (-0.2, 0.1)]}],
    )
    assert passed is False
    assert any("duct" in reason for reason in reasons)


def test_check_head_clearance_60cm_reports_invalid_polygon() -> None:
    passed, reasons = MODULE.check_head_clearance_60cm((0.0, 0.0), [{"type": "tray", "xy_polygon": "bad"}])
    assert passed is False
    assert any("format invalid" in reason for reason in reasons)


def test_decide_esfr_k_factor_k200() -> None:
    k_factor, rule = MODULE.decide_esfr_k_factor(9.1, True)
    assert (k_factor, rule) == (200, "NFTC 103B ESFR <=9.1m")


def test_decide_esfr_k_factor_k360() -> None:
    k_factor, rule = MODULE.decide_esfr_k_factor(13.0, True)
    assert (k_factor, rule) == (360, "NFTC 103B ESFR 12m~13.7m")


def test_decide_esfr_k_factor_requires_review_above_limit() -> None:
    k_factor, rule = MODULE.decide_esfr_k_factor(14.0, True)
    assert k_factor == 0
    assert "human review" in rule


def test_decide_esfr_k_factor_not_applicable() -> None:
    k_factor, rule = MODULE.decide_esfr_k_factor(8.0, False)
    assert (k_factor, rule) == (0, "NFTC 103B not applicable")


def test_check_shutoff_pressure_pass() -> None:
    passed, message = MODULE.check_shutoff_pressure(1.0, 1.39)
    assert passed is True
    assert "<=" in message


def test_check_shutoff_pressure_fail() -> None:
    passed, message = MODULE.check_shutoff_pressure(1.0, 1.41)
    assert passed is False
    assert "exceeds" in message


def test_check_shutoff_pressure_invalid_rated() -> None:
    with pytest.raises(ValueError):
        MODULE.check_shutoff_pressure(0.0, 1.0)


def test_check_emergency_power_duration_pass() -> None:
    passed, message = MODULE.check_emergency_power_duration(20.0)
    assert passed is True
    assert ">= 20.0min" in message


def test_check_emergency_power_duration_fail() -> None:
    passed, message = MODULE.check_emergency_power_duration(19.9)
    assert passed is False
    assert "< 20.0min" in message


def test_check_emergency_power_duration_invalid() -> None:
    with pytest.raises(ValueError):
        MODULE.check_emergency_power_duration(-5.0)
