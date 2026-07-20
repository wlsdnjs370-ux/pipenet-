"""Tests for redline item generation."""

from pathlib import Path
import json

import pandas as pd

from pipenet_converter.comparator import CompareTolerance, compare_networks
from pipenet_converter.isometric import render_isometric_png
from pipenet_converter.models import Fitting, Node, Nozzle, Pipe, PipeNetwork
from pipenet_converter.redline import create_redline_items, write_redline_items


def test_pipe_midpoint_redline_works() -> None:
    reference = _network()
    candidate = _network()
    candidate.pipes["P001"].length_m = 2.0
    issues = compare_networks(reference, candidate, CompareTolerance(length_m=0.1))

    items = create_redline_items(issues, reference, candidate)

    pipe_item = next(item for item in items if item.issue_type == "PIPE_LENGTH_MISMATCH")
    assert pipe_item.object_type == "Pipe"
    assert pipe_item.object_id == "P001"
    assert pipe_item.x == 5.0
    assert pipe_item.y == 0.0
    assert pipe_item.z == 0.0
    assert pipe_item.suggested_action == "Check CAD centerline length and SDF pipe length"


def test_nozzle_redline_works() -> None:
    reference = _network()
    candidate = _network()
    candidate.nozzles["NZ1"].status = 0
    issues = compare_networks(reference, candidate, CompareTolerance())

    items = create_redline_items(issues, reference, candidate)

    nozzle_item = next(item for item in items if item.issue_type == "NOZZLE_STATUS_MISMATCH")
    assert nozzle_item.object_type == "Nozzle"
    assert nozzle_item.object_id == "NZ1"
    assert nozzle_item.x == 10.0
    assert nozzle_item.y == 5.0
    assert nozzle_item.z == 0.5


def test_write_redline_items_creates_csv_and_json(tmp_path: Path) -> None:
    reference = _network()
    candidate = _network()
    candidate.pipes["P001"].fittings = [Fitting("elbow", 1)]
    issues = compare_networks(reference, candidate, CompareTolerance())
    items = create_redline_items(issues, reference, candidate)

    write_redline_items(items, tmp_path)

    csv_path = tmp_path / "redline_items.csv"
    json_path = tmp_path / "redline_items.json"
    assert csv_path.exists()
    assert json_path.exists()
    csv_rows = pd.read_csv(csv_path)
    json_rows = json.loads(json_path.read_text(encoding="utf-8"))
    assert list(csv_rows.columns) == [
        "redline_id",
        "issue_type",
        "severity",
        "object_type",
        "object_id",
        "x",
        "y",
        "z",
        "message",
        "suggested_action",
    ]
    assert len(csv_rows) == len(items)
    assert len(json_rows) == len(items)


def test_render_isometric_png_accepts_redline_items(tmp_path: Path) -> None:
    reference = _network()
    candidate = _network()
    candidate.pipes["P001"].length_m = 2.0
    issues = compare_networks(reference, candidate, CompareTolerance(length_m=0.1))
    items = create_redline_items(issues, reference, candidate)
    output_path = tmp_path / "redline_iso.png"

    render_isometric_png(candidate, output_path, redline_items=items)

    assert output_path.exists()
    assert output_path.stat().st_size > 0


def _network() -> PipeNetwork:
    network = PipeNetwork(title="Redline")
    network.add_node(Node("INPUT", 0.0, 0.0, 0.0, "Input"))
    network.add_node(Node("N001", 10.0, 0.0, 0.0, "main_pipe"))
    network.add_node(Node("@/1", 10.0, 5.0, 0.5, "head_output"))
    network.add_pipe(
        Pipe(
            "P001",
            "INPUT",
            "N001",
            0.15,
            10.0,
            0.0,
            fittings=[Fitting("gate", 1)],
        )
    )
    network.add_nozzle(Nozzle("NZ1", "N001", "@/1", 0.00266666667, status=1))
    return network
