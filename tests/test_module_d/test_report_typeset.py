# -*- coding: utf-8 -*-
"""모듈 D — D5 수리계산서 PDF.

여기가 통과한다는 것은 "PIPENET 이 워드로 찍은 수치를, 워드를 거치지 않고 우리가
같은 글자로 찍는다"는 뜻이다. 마지막 자리가 갈리면 검토자가 우리 계산서와 원본
계산서를 나란히 놓고 다른 숫자를 보게 된다.
"""
from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

import pytest

_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core.d_report_typeset import (  # noqa: E402
    build_blocks,
    fixed,
    fortran_e,
    fortran_g,
    integer,
    paginate,
    render_report,
)
from core.d_result_binder import bind_results  # noqa: E402

SUB = _ROOT / "routes" / "제출용[최종]"
AUTO_SDF, AUTO_XML = SUB / "2. Pipenet_auto.sdf", SUB / "2. Pipenet_auto.xml"

pytestmark = pytest.mark.skipif(not AUTO_XML.exists(), reason="제출용 결과 XML 없음")


@pytest.fixture(scope="module")
def bound():
    return bind_results(AUTO_SDF, AUTO_XML)


@pytest.fixture(scope="module")
def blocks(bound):
    return build_blocks(bound, datetime(2026, 1, 1))[0]


# ── Fortran 수치 서식 ───────────────────────────────────────────────────────


@pytest.mark.parametrize("value,shown", [
    # 워드 원문에서 옮겨 적는다. 파이썬 % 서식(짝수 반올림)이면 마지막 자리가 갈린다.
    (81.8153586, "81.82"),
    (1.0245343720842488, "1.025"),
    (1.1604727404363364, "1.160"),
    (0.9142469650696211, "0.9142"),
])
def test_fortran_g_rounds_half_up(value, shown):
    assert fortran_g(value) == shown


def test_fortran_g_marks_the_integer_with_a_point():
    # 유효숫자를 다 쓴 자리는 점으로 끝난다 — 점이 없으면 정확한 정수로 오해한다.
    assert fortran_g(-1042.0) == "-1042."
    assert fortran_g(2286.123) == "2286."


def test_fortran_g_falls_to_exponent_below_a_tenth():
    assert fortran_g(0.098397) == "9.8397E-02"
    assert fortran_g(0.1) == "0.1000"


def test_fortran_e_keeps_the_mantissa_under_one():
    # NOZZLE CONFIGURATION 의 E 열은 가수를 0.1~1 로 둔다 (0.10000E+01).
    assert fortran_e(1.0) == "0.10000E+01"
    assert fortran_e(0.86329) == "0.86329E+00"


def test_carrying_does_not_lose_a_digit():
    assert fortran_g(9.99999) == "10.00"
    assert fortran_e(0.999999) == "0.10000E+01"


def test_missing_values_stay_blank():
    # 지시서 7-4 — 없는 값은 추정하지 않는다.
    for fmt in (fortran_g, fortran_e, fixed, integer):
        assert fmt(None) == ""
        assert fmt(float("nan")) == ""


def test_fixed_and_integer_round_half_up():
    assert fixed(0.00125, 4) == "0.0013"
    assert integer(2.5) == "3"


# ── 표시값 ──────────────────────────────────────────────────────────────────


def _rows(blocks, title):
    return next(b for b in blocks if b.title == title).body


def test_pressure_is_shown_in_the_declared_unit(blocks):
    # SI 로 결합해 두고 화면에는 파일이 선언한 kg/cm2 로 되돌린다 (지시서 7-5).
    # 기대값은 워드 계산서 FLOW IN PIPES 원문에서 옮겨 적는다 (관로 10).
    row = next(r for r in _rows(blocks, "FLOW IN PIPES") if r.split()[0] == "10")
    cells = row.split()
    assert cells[4:10] == ["5.015", "2.612", "2.403", "2.403", "2286.", "10.19"]
    assert cells[-1] == "E"


def test_nozzle_result_pressure_is_not_gauged_twice(blocks):
    # 결과 입구압은 이미 게이지압이다. 기준면을 또 빼면 대기압 한 겹만큼 낮아진다.
    # 기대값은 워드 계산서 FLOW THROUGH NOZZLES 원문 (노즐 2).
    row = next(r for r in _rows(blocks, "FLOW THROUGH NOZZLES") if r.split()[0] == "2")
    assert row.split()[2] == "0.86329E+00"


def test_nozzle_input_pressure_is_gauged(blocks):
    # 반대로 노즐 입력표의 최소·최대 작동압은 절대압이라 기준면을 빼야 한다.
    row = next(r for r in _rows(blocks, "NOZZLE CONFIGURATION") if r.split()[0] == "1")
    assert row.split()[-2:] == ["0.10000E+01", "0.12000E+02"]


def test_nominal_bore_keeps_the_catalogue_number(blocks):
    # 호칭경은 환산할 값이 아니라 규격 번호다.
    row = next(r for r in _rows(blocks, "PIPE CONFIGURATION") if r.split()[0] == "10")
    assert row.split()[3] == "65.00"


# ── 절 구성 ─────────────────────────────────────────────────────────────────


def test_sections_follow_the_pipenet_order(blocks):
    titles = [b.title for b in blocks]
    ordered = [t for t in titles if not t.startswith("AVAILABLE PIPE SIZES")]
    assert ordered == [
        "CONTROL INFORMATION", "FLUID SYSTEM", "DESIGN INFORMATION",
        "PIPE CONFIGURATION", "PIPE FITTINGS", "NOZZLE CONFIGURATION",
        "SPECIAL EQUIPMENT", "DESIGNED DIAMETERS & FLOWRATES", "FLOW IN PIPES",
        "FLOW THROUGH NOZZLES", "FLOW AT INLETS",
        "MATERIALS TAKE-OFF / PIPE LENGTHS", "MATERIALS TAKE-OFF / NOZZLES",
        "MATERIALS TAKE-OFF / FITTINGS", "MATERIALS TAKE-OFF / OTHER EQUIPMENT",
        "생성 출처와 고지", "WARNINGS",
    ]


def test_every_row_of_the_input_survives(blocks):
    for title, count in (("PIPE CONFIGURATION", 136), ("FLOW IN PIPES", 136),
                         ("NOZZLE CONFIGURATION", 30), ("FLOW THROUGH NOZZLES", 30),
                         ("DESIGNED DIAMETERS & FLOWRATES", 136)):
        assert len(_rows(blocks, title)) == count, title


def test_omitted_columns_say_why(blocks):
    # 지시서 7-4 — 없는 값은 조용히 채우지 않고 없다고 적는다.
    designed = next(b for b in blocks if b.title == "DESIGNED DIAMETERS & FLOWRATES")
    assert any("Pipe Group" in line for line in designed.tail)
    nozzles = next(b for b in blocks if b.title == "FLOW THROUGH NOZZLES")
    assert any("FlowDens" in line for line in nozzles.tail)


def test_warnings_match_pipenets_own(bound):
    """경고는 옮겨 적는 값이 아니라 우리가 내리는 판정이다.

    유속 초과는 재질별 규격표의 최대유속을, 최소압 미달은 절대압/게이지압 기준면 차를
    제대로 잡아야 같은 집합이 나온다. 기대값은 워드 계산서 WARNINGS 원문에서 옮겨 적는다.
    """
    _blocks, _sheet, over, under = build_blocks(bound, datetime(2026, 1, 1))
    assert over == ("10", "11")
    assert under == tuple(str(n) for n in list(range(1, 16)) + [18, 19, 22])


def test_provenance_names_this_program_not_pipenet(blocks):
    body = "\n".join(next(b for b in blocks if b.title == "생성 출처와 고지").body)
    assert "PIPENET 이 만든 문서가 아니다" in body
    assert "2. Pipenet_auto.xml" in body and "모듈 D5" in body


def test_continued_table_reprints_its_header():
    from core.d_report_typeset import Block

    block = Block("T", header=("h1", "h2"), body=tuple(str(n) for n in range(20)))
    pages = paginate([block], 10)
    assert len(pages) > 1
    assert pages[1][:3] == ["T (이어짐)", "=" * 18, "h1"]


# ── PDF ─────────────────────────────────────────────────────────────────────


def test_report_renders_without_pipenet(tmp_path):
    out = tmp_path / "report.pdf"
    result = render_report(AUTO_SDF, AUTO_XML, out, generated=datetime(2026, 1, 1))
    assert out.exists() and result.pages > 1
    # 표가 접히면 열이 어긋난다 — 폭에 맞춰 글자를 줄이되 최대치를 넘지 않는다.
    assert 0 < result.font_pt <= 8.5
