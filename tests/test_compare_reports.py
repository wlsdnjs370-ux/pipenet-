# -*- coding: utf-8 -*-
"""calibration/compare_reports.py (모듈 A / T0) 테스트."""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent
for _p in (_ROOT, _ROOT / "core"):
    if str(_p) not in sys.path:
        sys.path.insert(0, str(_p))

from calibration.compare_reports import (  # noqa: E402
    ComparisonReport,
    compare,
    profile_from_text,
)


def _section(title: str, *lines: str) -> str:
    return "\n".join([title, "-" * len(title), *lines])


CONFIG_HEADER = (
    "Pipe Label   Input Node   Output Node    Nom.Bore       Length      "
    "Elevation       C      Fitt.eq.lnth"
)


def test_fitting_equivalent_length_multiplies_by_count():
    text = _section(
        "PIPE FITTINGS",
        "Pipe         Number x Type  Equivalent Length",
        "1         2 x  2    4.267       1 x  4    9.144",
    )
    prof = profile_from_text(text, "t")
    assert prof.fitting_count == 3
    assert prof.fitting_eq_length_m == pytest.approx(2 * 4.267 + 9.144, abs=0.005)


def test_non_numeric_pipe_labels_are_counted():
    """자동본은 기계실/라이저를 'M1/0', 'R4/0' 라벨로 낸다. 숫자만 세면 21개가 증발한다."""
    text = "\n".join(
        [
            _section(
                "PIPE CONFIGURATION",
                CONFIG_HEADER,
                "1            1          101          150.0    1.540    0.000   120.0   10.67",
                "M1/0      M1/0         M2/0          150.0    1.540   -2.000   120.0    0.000",
            ),
            "",
            _section(
                "PIPE FITTINGS",
                "Pipe         Number x Type  Equivalent Length",
                "M1/0         1 x  4    9.144",
            ),
        ]
    )
    prof = profile_from_text(text, "t")
    assert prof.pipe_count == 2
    assert prof.elevation_nonzero_pipes == 1
    assert prof.fitting_count == 1


def test_missing_section_yields_none_not_zero():
    """근거가 없으면 0이 아니라 None + 미확정 사유."""
    prof = profile_from_text("(빈 리포트)", "t")
    assert prof.pipe_count is None
    assert prof.fitting_eq_length_m is None
    assert prof.total_flow_lpm is None
    assert any("PIPE FITTINGS" in m for m in prof.unresolved)


def test_local_drops_excluded_from_major_drop():
    text = _section(
        "PIPE CONFIGURATION",
        CONFIG_HEADER,
        "1            1          101          150.0    1.540   -66.700   120.0    0.000",
        "2            2          102          150.0    1.540   -0.100    120.0    0.000",
    )
    prof = profile_from_text(text, "t")
    assert prof.elevation_major_drop_m == pytest.approx(-66.7)
    assert prof.elevation_total_drop_m == pytest.approx(-66.8)


# --- 실물 대조 (지시서 §2 재현) -------------------------------------------------

_DEFAULT_REF_DIR = Path.home() / "Desktop" / "제출용[최종]"
MANUAL_DOCX = Path(
    os.environ.get("PIPENET_REF_MANUAL_DOCX", _DEFAULT_REF_DIR / "3. 수리계산서_Pipenet_수작업.docx")
)
AUTO_DOCX = Path(
    os.environ.get("PIPENET_REF_AUTO_DOCX", _DEFAULT_REF_DIR / "3. 수리계산서_Pipenet_자동화.docx")
)

requires_ref = pytest.mark.skipif(
    not (MANUAL_DOCX.exists() and AUTO_DOCX.exists()),
    reason="기준 DOCX 없음 — PIPENET_REF_MANUAL_DOCX / PIPENET_REF_AUTO_DOCX 로 경로 지정",
)


@pytest.fixture(scope="module")
def report() -> ComparisonReport:
    return compare(MANUAL_DOCX, AUTO_DOCX)


@requires_ref
def test_flow_match_pct_is_78_7(report):
    assert round(report.flow_match_pct, 1) == 78.7


@requires_ref
@pytest.mark.parametrize(
    "attr, manual, auto",
    [
        ("pipe_count", 103, 136),
        ("elevation_nonzero_pipes", 34, 2),
        ("fitting_count", 73, 62),
        ("nozzle_count", 30, 30),
    ],
)
def test_section_2_counts(report, attr, manual, auto):
    assert getattr(report.manual, attr) == manual
    assert getattr(report.auto, attr) == auto


@requires_ref
def test_section_2_scalars(report):
    assert report.manual.fitting_eq_length_m == pytest.approx(253.54, abs=0.01)
    assert report.auto.fitting_eq_length_m == pytest.approx(132.60, abs=0.01)
    # 지시서 §2 의 "총 하강 75.0 → 53.4" 는 수작업은 주요하강, 자동은 전체하강을
    # 집계한 값이다. 두 축을 분리해 둘 다 고정한다.
    assert report.manual.elevation_major_drop_m == pytest.approx(-75.0, abs=0.05)
    assert report.manual.elevation_total_drop_m == pytest.approx(-78.0, abs=0.05)
    assert report.auto.elevation_major_drop_m == pytest.approx(-53.19, abs=0.05)
    assert report.auto.elevation_total_drop_m == pytest.approx(-53.4, abs=0.05)
    assert report.manual.total_flow_lpm == pytest.approx(2903.1, abs=0.1)
    assert report.auto.total_flow_lpm == pytest.approx(2286.1, abs=0.1)
    assert (report.manual.deviation_min_pct, report.manual.deviation_max_pct) == (0.05, 49.97)
    assert (report.auto.deviation_min_pct, report.auto.deviation_max_pct) == (-26.83, 23.97)


@requires_ref
def test_material_libraries(report):
    assert set(report.manual.materials) == {"KSD 3507", "CPVC", "FX", "DP"}
    assert set(report.auto.materials) == {"KSD 3507"}
    assert sorted(report.manual.pipes_by_material.values(), reverse=True) == [49, 30, 20, 4]
    assert list(report.auto.pipes_by_material) == ["KSD 3507"]


@requires_ref
def test_markdown_leads_with_match_pct(report):
    assert report.to_markdown().splitlines()[0].startswith("# 총유량 일치율 78.7%")
