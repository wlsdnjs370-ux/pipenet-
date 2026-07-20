"""Regression tests for an optional real PipeNet SDF sample."""

from pathlib import Path

import pytest

from pipenet_converter.export_tables import write_network_tables
from pipenet_converter.sdf_parser import parse_sdf
from pipenet_converter.sdf_writer import write_sdf


REAL_SAMPLE_SDF = (
    Path(__file__).resolve().parents[1]
    / "data"
    / "sample"
    / "1-1. 다이소 세종허브센터 지상4층 창고.sdf"
)


def test_real_daiso_sdf_parse_export_write_round_trip(tmp_path: Path) -> None:
    if not REAL_SAMPLE_SDF.exists():
        pytest.skip(f"Optional real sample SDF is absent: {REAL_SAMPLE_SDF}")

    parsed = parse_sdf(REAL_SAMPLE_SDF)

    assert len(parsed.nodes) > 0
    assert len(parsed.pipes) > 0
    assert len(parsed.nozzles) > 0
    assert len(parsed.active_nozzles()) > 0

    export_dir = tmp_path / "parsed_tables"
    write_network_tables(parsed, export_dir)
    assert (export_dir / "network_3d_nodes.csv").exists()
    assert (export_dir / "network_3d_pipes.csv").exists()
    assert (export_dir / "network_3d_nozzles.csv").exists()

    generated_sdf = tmp_path / "round_trip.sdf"
    write_sdf(parsed, generated_sdf)
    reparsed = parse_sdf(generated_sdf)

    assert len(reparsed.nodes) == len(parsed.nodes)
    assert len(reparsed.pipes) == len(parsed.pipes)
    assert len(reparsed.nozzles) == len(parsed.nozzles)
