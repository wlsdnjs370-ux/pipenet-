# AGENTS.md

This package follows the root project instructions.

## Purpose

Build a Python package that converts fire protection sprinkler CAD-derived data into a PipeNet-compatible SDF file.

The goal is a calculation-ready 3D sprinkler pipe centerline graph, not a full architectural BIM model.

## Implementation Policy

- Use Python 3.11+.
- Use small testable functions.
- Use type hints and docstrings.
- Use pytest.
- Keep DXF extraction, graph construction, elevation rules, validation, rendering, and SDF writing separated.
- Do not parse DWG directly.
- Do not require AutoCAD COM automation or licensed CAD software.
- Do not use OCR in version 1.
