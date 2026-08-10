# -*- coding: utf-8 -*-
"""모듈 D — D6 일괄 처리.

여기가 통과한다는 것은 "폴더를 통째로 걸어 두고 자리를 떠도 된다"는 뜻이다.
한 세트가 터지면 그 세트만 실패로 남고, 나머지는 끝까지 간다 (지시서 D6 수용 기준).
"""
from __future__ import annotations

import shutil
import sys
from datetime import datetime
from pathlib import Path

import pytest
from pypdf import PdfReader

_ROOT = Path(__file__).resolve().parent.parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core.d_batch import (  # noqa: E402
    DEFAULT_PRESETS,
    REPORT_PART,
    discover,
    process_set,
    run_batch,
)

SUB = _ROOT / "routes" / "제출용[최종]"
AUTO_SDF, AUTO_XML = SUB / "2. Pipenet_auto.sdf", SUB / "2. Pipenet_auto.xml"
STAMP = datetime(2026, 1, 1)

pytestmark = pytest.mark.skipif(not AUTO_XML.exists(), reason="제출용 결과 XML 없음")


def _plant(root: Path, name: str, *, xml: bool = True, slf: bool = False) -> Path:
    """입력 한 벌을 심는다. 이름만 바꾼 사본이라 수치는 제출용과 같다."""
    root.mkdir(parents=True, exist_ok=True)
    sdf = root / f"{name}.sdf"
    shutil.copy(AUTO_SDF, sdf)
    if xml:
        shutil.copy(AUTO_XML, sdf.with_suffix(".xml"))
    if slf:
        sdf.with_suffix(".slf").write_bytes(b"")
    return sdf


# ── 세트 인식 ───────────────────────────────────────────────────────────────


def test_sets_are_grouped_by_stem_beside_the_sdf(tmp_path):
    _plant(tmp_path / "101동", "구간1", slf=True)
    _plant(tmp_path / "102동", "구간1", xml=False)

    found = discover(tmp_path)
    assert [f.stem for f in found] == ["구간1", "구간1"]
    by_dir = {f.sdf.parent.name: f for f in found}
    assert by_dir["101동"].xml is not None and by_dir["101동"].slf is not None
    # SLF 는 선택이다 — 없다고 세트가 깨지지 않는다.
    assert by_dir["102동"].xml is None and by_dir["102동"].slf is None
    assert by_dir["102동"].missing == (".xml", ".slf")


def test_office_temp_files_are_not_input(tmp_path):
    _plant(tmp_path, "구간1")
    shutil.copy(AUTO_SDF, tmp_path / "~$구간1.sdf")
    assert [f.stem for f in discover(tmp_path)] == ["구간1"]


def test_same_name_in_two_folders_does_not_overwrite(tmp_path):
    _plant(tmp_path / "101동", "구간1")
    _plant(tmp_path / "102동", "구간1")
    out = tmp_path / "out"

    report = run_batch(tmp_path, out, generated=STAMP)
    assert len(report.succeeded) == 2
    assert {r.combined.relative_to(out).as_posix() for r in report.results} == {
        "101동/구간1.pdf", "102동/구간1.pdf"}


# ── 한 세트 ─────────────────────────────────────────────────────────────────


def test_a_set_yields_the_report_and_both_drawings(tmp_path):
    sdf = _plant(tmp_path, "구간1")
    result = process_set(discover(sdf)[0], tmp_path / "out", generated=STAMP)

    assert result.ok
    assert [p.name for p in result.parts] == [REPORT_PART, *DEFAULT_PRESETS]
    assert result.combined.exists()
    # 합본은 조각의 합이고, 조각 이름이 그대로 책갈피가 된다.
    merged = PdfReader(str(result.combined))
    assert len(merged.pages) == sum(p.pages for p in result.parts)
    assert [o.title for o in merged.outline] == [REPORT_PART, *DEFAULT_PRESETS]


def test_without_results_the_drawings_still_come_out(tmp_path):
    # 결과 XML 이 없으면 계산서는 만들 수 없다. 조용히 빈 표를 찍는 대신 사유를 남긴다.
    sdf = _plant(tmp_path, "구간1", xml=False)
    result = process_set(discover(sdf)[0], tmp_path / "out", generated=STAMP)

    assert result.ok
    assert [p.name for p in result.parts] == list(DEFAULT_PRESETS)
    assert any("결과 XML 이 없다" in note for note in result.notes)


def test_an_unknown_preset_is_refused_before_anything_is_written(tmp_path):
    sdf = _plant(tmp_path, "구간1")
    out = tmp_path / "out"
    result = process_set(discover(sdf)[0], out, presets=("압력본", "무지개본"),
                         generated=STAMP)

    assert not result.ok and "무지개본" in result.error
    assert not out.exists()


# ── 격리 ────────────────────────────────────────────────────────────────────


def test_one_broken_set_does_not_stop_the_others(tmp_path):
    for n in range(1, 6):
        _plant(tmp_path, f"구간{n}", xml=n != 4)
    # 4 번은 SDF 자체가 열리지 않는다.
    (tmp_path / "구간4.sdf").write_text("<Project", encoding="utf-8")

    report = run_batch(tmp_path, tmp_path / "out", generated=STAMP)

    assert len(report.results) == 5
    assert len(report.failed) == 1
    assert report.failed[0].file_set.stem == "구간4"
    assert report.failed[0].error
    assert {r.file_set.stem for r in report.succeeded} == {"구간1", "구간2", "구간3", "구간5"}
    for result in report.succeeded:
        assert result.combined.exists()


def test_a_broken_piece_does_not_take_the_rest_of_the_set(tmp_path, monkeypatch):
    import core.d_batch as batch

    real = batch.render_iso

    def fail_on_flow(source, output, *, preset=None, **kw):
        if preset == "유량본":
            raise RuntimeError("일부러 낸 고장")
        return real(source, output, preset=preset, **kw)

    monkeypatch.setattr(batch, "render_iso", fail_on_flow)
    sdf = _plant(tmp_path, "구간1")
    result = process_set(discover(sdf)[0], tmp_path / "out", generated=STAMP)

    assert result.partial and not result.ok
    assert [p.name for p in result.done] == [REPORT_PART, "압력본"]
    # 나온 것만 묶는다 — 못 나온 조각 때문에 합본까지 잃지는 않는다.
    assert [o.title for o in PdfReader(str(result.combined)).outline] == [
        REPORT_PART, "압력본"]
    broken = next(p for p in result.parts if not p.ok)
    assert broken.name == "유량본" and "일부러 낸 고장" in broken.error


def test_the_report_names_every_file_and_counts_them(tmp_path):
    _plant(tmp_path, "구간1")
    (tmp_path / "구간2.sdf").write_text("<Project", encoding="utf-8")

    report = run_batch(tmp_path, tmp_path / "out", generated=STAMP)
    text = report.text()

    assert "구간1.sdf" in text and "구간2.sdf" in text
    assert "성공 1 · 부분 0 · 실패 1" in text


def test_progress_is_reported_as_it_goes(tmp_path):
    for n in range(1, 4):
        _plant(tmp_path, f"구간{n}", xml=False)
    seen = []
    run_batch(tmp_path, tmp_path / "out", generated=STAMP,
              progress=lambda i, total, r: seen.append((i, total, r.ok)))
    # 끝나고 몰아서가 아니라 세트마다 한 번씩 — 진행 표시가 실제 진행이어야 한다.
    assert seen == [(1, 3, True), (2, 3, True), (3, 3, True)]
