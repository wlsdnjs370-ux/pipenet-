"""Tests for applying elevation rules."""

from pathlib import Path

import pytest

from pipenet_converter.elevation import (
    apply_elevation_rules,
    load_elevation_rules,
    recompute_pipe_rise_and_length,
)
from pipenet_converter.models import Node, Nozzle, Pipe, PipeNetwork


RULES_CSV = """rule_id,priority,match_field,match_value,z_m,description
Z_MAIN_4F_UPPER,10,node_type,main_pipe,48.55,4F upper warehouse main pipe
Z_BRANCH_4F_UPPER,20,node_type,branch_pipe,48.85,4F upper warehouse branch pipe
Z_HEAD_4F_UPPER,30,node_type,head,49.15,4F upper warehouse sprinkler head
Z_RISER_4F,40,node_type,riser,34.95,4F riser connection
Z_DEFAULT,999,default,default,48.55,default 4F upper main
"""


def test_load_elevation_rules_sorts_by_priority(tmp_path: Path) -> None:
    rules_path = tmp_path / "elevation_rules.csv"
    rules_path.write_text(RULES_CSV, encoding="utf-8")

    rules = load_elevation_rules(rules_path)

    assert [rule.rule_id for rule in rules] == [
        "Z_MAIN_4F_UPPER",
        "Z_BRANCH_4F_UPPER",
        "Z_HEAD_4F_UPPER",
        "Z_RISER_4F",
        "Z_DEFAULT",
    ]
    assert rules[0].z_m == 48.55


def test_invalid_elevation_rule_raises_clear_error(tmp_path: Path) -> None:
    rules_path = tmp_path / "bad_elevation_rules.csv"
    rules_path.write_text(
        "rule_id,priority,match_field,match_value,z_m,description\n"
        "BAD,10,node_type,head,not-a-number,bad\n",
        encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Invalid elevation rule value"):
        load_elevation_rules(rules_path)


def test_apply_elevation_rules_sets_head_branch_and_default_z(tmp_path: Path) -> None:
    rules_path = tmp_path / "elevation_rules.csv"
    rules_path.write_text(RULES_CSV, encoding="utf-8")
    rules = load_elevation_rules(rules_path)
    network = PipeNetwork(title="Elevation")
    network.add_node(Node("N_MAIN", 0.0, 0.0, 0.0, "main_pipe"))
    network.add_node(Node("N_BRANCH", 1000.0, 0.0, 0.0, "branch_pipe"))
    network.add_node(Node("N_HEAD", 1000.0, 500.0, 0.0, "head"))
    network.add_node(Node("@/1", 1000.0, 500.0, 0.0, "head_output"))
    network.add_node(Node("N_UNKNOWN", 2000.0, 0.0, 0.0, "unknown"))
    network.add_nozzle(Nozzle("NZ1", "N_HEAD", "@/1", 0.00266666667))

    apply_elevation_rules(network, rules)

    assert network.nodes["N_MAIN"].z == 48.55
    assert network.nodes["N_BRANCH"].z == 48.85
    assert network.nodes["N_HEAD"].z == 49.15
    assert network.nodes["N_UNKNOWN"].z == 48.55
    assert network.nodes["N_UNKNOWN"].metadata["elevation_rule_id"] == "Z_DEFAULT"
    assert network.nodes["N_HEAD"].metadata["elevation_rule_id"] == "Z_HEAD_4F_UPPER"
    assert network.nodes["@/1"].z == 49.15


def test_metadata_match_rule_can_assign_elevation() -> None:
    network = PipeNetwork(title="Metadata elevation")
    network.add_node(Node("N1", 0.0, 0.0, 0.0, "pipe", metadata={"zone": "rack"}))
    rules = [
        load_rule("Z_RACK", 1, "metadata.zone", "rack", 50.25),
        load_rule("Z_DEFAULT", 999, "default", "default", 48.55),
    ]

    apply_elevation_rules(network, rules)

    assert network.nodes["N1"].z == 50.25


def test_recompute_pipe_rise_and_length_in_orthogonal_mode() -> None:
    network = PipeNetwork(title="Recompute")
    network.add_node(Node("A", 0.0, 0.0, 48.55, "main_pipe"))
    network.add_node(Node("B", 3.0, 4.0, 48.55, "main_pipe"))
    network.add_node(Node("C", 3.0, 4.0, 49.15, "head"))
    network.add_pipe(Pipe("P1", "A", "B", 0.15, 0.0, 0.0))
    network.add_pipe(Pipe("P2", "B", "C", 0.032, 0.0, 0.0))
    network.add_pipe(Pipe("P3", "A", "C", 0.032, 0.0, 0.0))

    recompute_pipe_rise_and_length(network)

    assert network.pipes["P1"].rise_m == 0.0
    assert network.pipes["P1"].length_m == 5.0
    assert round(network.pipes["P2"].rise_m, 6) == 0.6
    assert round(network.pipes["P2"].length_m, 6) == 0.6
    assert round(network.pipes["P3"].rise_m, 6) == 0.6
    assert round(network.pipes["P3"].length_m, 6) == 5.6


def load_rule(rule_id: str, priority: int, match_field: str, match_value: str, z_m: float):
    from pipenet_converter.elevation import ElevationRule

    return ElevationRule(
        rule_id=rule_id,
        priority=priority,
        match_field=match_field,
        match_value=match_value,
        z_m=z_m,
        description="test",
    )
