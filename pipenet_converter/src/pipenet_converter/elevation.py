"""Elevation rule handling.

Version 1 applies human-authored CSV rules from section drawings and system
diagrams. It does not attempt to infer elevations directly from CAD graphics.
"""

from __future__ import annotations

import csv
from dataclasses import dataclass
from math import dist
from pathlib import Path

from pipenet_converter.models import Node, PipeNetwork


@dataclass(slots=True)
class ElevationRule:
    """A rule assigning elevation to matching nodes."""

    rule_id: str
    priority: int
    match_field: str
    match_value: str
    z_m: float
    description: str


def load_elevation_rules(path: str | Path) -> list[ElevationRule]:
    """Load elevation rules from ``elevation_rules.csv``."""
    rules_path = Path(path)
    if not rules_path.exists():
        raise FileNotFoundError(f"Elevation rules CSV file not found: {rules_path}")
    with rules_path.open(newline="", encoding="utf-8-sig") as csv_file:
        reader = csv.DictReader(csv_file)
        _require_columns(
            reader.fieldnames,
            {"rule_id", "priority", "match_field", "match_value", "z_m", "description"},
            rules_path,
        )
        rules = []
        for row_number, row in enumerate(reader, start=2):
            try:
                rules.append(
                    ElevationRule(
                        rule_id=str(row["rule_id"]).strip(),
                        priority=int(row["priority"]),
                        match_field=str(row["match_field"]).strip(),
                        match_value=str(row["match_value"]).strip(),
                        z_m=float(row["z_m"]),
                        description=str(row.get("description", "")).strip(),
                    )
                )
            except ValueError as exc:
                raise ValueError(f"Invalid elevation rule value in {rules_path} at row {row_number}: {exc}") from exc
    if not rules:
        raise ValueError(f"Elevation rules CSV file {rules_path} does not contain any rules.")
    _validate_elevation_rules(rules, rules_path)
    return sorted(rules, key=lambda rule: rule.priority)


def apply_elevation_rules(network: PipeNetwork, rules: list[ElevationRule]) -> PipeNetwork:
    """Apply elevation rules to network nodes and return the same network."""
    sorted_rules = sorted(rules, key=lambda rule: rule.priority)
    default_rule = _default_rule(sorted_rules)

    for node in network.nodes.values():
        rule = _matching_rule(node, sorted_rules) or default_rule
        if rule is not None:
            node.z = rule.z_m
            node.metadata["elevation_rule_id"] = rule.rule_id

    head_rule = _head_rule(sorted_rules)
    if head_rule is not None:
        for nozzle in network.nozzles.values():
            output_node = network.nodes.get(nozzle.output_node)
            if output_node is not None:
                output_node.z = head_rule.z_m
                output_node.metadata["elevation_rule_id"] = head_rule.rule_id

    return network


def recompute_pipe_rise_and_length(network: PipeNetwork, mode: str = "orthogonal") -> PipeNetwork:
    """Recompute pipe ``rise_m`` and ``length_m`` after elevation assignment."""
    if mode != "orthogonal":
        raise ValueError(f"Unsupported length recompute mode: {mode!r}")

    for pipe in network.pipes.values():
        from_node = network.nodes[pipe.from_node]
        to_node = network.nodes[pipe.to_node]
        horizontal_distance = dist((from_node.x, from_node.y), (to_node.x, to_node.y))
        rise = to_node.z - from_node.z
        pipe.rise_m = rise

        if horizontal_distance == 0.0:
            pipe.length_m = abs(rise)
        elif rise == 0.0:
            pipe.length_m = horizontal_distance
        else:
            pipe.length_m = horizontal_distance + abs(rise)

    return network


def _matching_rule(node: Node, rules: list[ElevationRule]) -> ElevationRule | None:
    for rule in rules:
        if rule.match_field == "default":
            continue
        if _rule_matches(node, rule):
            return rule
    return None


def _rule_matches(node: Node, rule: ElevationRule) -> bool:
    if rule.match_field == "node_type":
        return node.node_type == rule.match_value
    if rule.match_field == "source":
        return node.source == rule.match_value
    if rule.match_field.startswith("metadata."):
        metadata_key = rule.match_field.removeprefix("metadata.")
        value = node.metadata.get(metadata_key)
        return str(value) == rule.match_value if value is not None else False
    return False


def _default_rule(rules: list[ElevationRule]) -> ElevationRule | None:
    return next((rule for rule in rules if rule.match_field == "default"), None)


def _head_rule(rules: list[ElevationRule]) -> ElevationRule | None:
    return next(
        (
            rule
            for rule in rules
            if rule.match_field == "node_type" and rule.match_value in {"head", "head_output"}
        ),
        None,
    )


def _require_columns(fieldnames: list[str] | None, required: set[str], path: Path) -> None:
    available = set(fieldnames or [])
    missing = sorted(required - available)
    if missing:
        raise ValueError(
            f"CSV file {path} is missing required columns: {', '.join(missing)}. "
            f"Available columns: {', '.join(sorted(available))}"
        )


def _validate_elevation_rules(rules: list[ElevationRule], path: Path) -> None:
    valid_fields = {"node_type", "source", "default"}
    for rule in rules:
        if rule.priority < 0:
            raise ValueError(f"Invalid elevation rule {rule.rule_id!r} in {path}: priority must be >= 0.")
        if rule.match_field not in valid_fields and not rule.match_field.startswith("metadata."):
            raise ValueError(
                f"Invalid elevation rule {rule.rule_id!r} in {path}: unsupported match_field "
                f"{rule.match_field!r}."
            )
