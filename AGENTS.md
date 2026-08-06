# AGENTS.md

## 1. Project Purpose

This project builds a Python package that converts fire protection sprinkler CAD-derived data into a PipeNet-compatible SDF file.

The primary output is not an architectural BIM model. The primary output is a calculation-ready 3D sprinkler pipe centerline graph that can be exported to PipeNet SDF XML and validated against drawings, rules, and existing SDF files.

The package must support this workflow:

1. Extract X/Y pipe network information from fire protection plan drawings exported as DXF.
2. Extract or apply Z/elevation information from section drawings and system diagrams through CSV rule files.
3. Merge X/Y/Z into a 3D pipe centerline network.
4. Represent the network as Nodes, Pipes, Nozzles, Fittings, Equipment, and Valves.
5. Render a check isometric PNG.
6. Export a PipeNet-compatible `.sdf` XML file.
7. Validate the generated network against design rules and existing SDF files.

Core comparison targets:

- head count and head positions
- pipe length
- pipe diameter
- pipe connectivity
- riser and elevation
- elbow count
- tee count
- reducer or diameter-change points
- valve and alarm valve location
- nozzle status and flow
- PipeNet SDF compatibility

## 2. Repo Layout

Use this intended layout for new project code:

```text
src/
  pipenet_converter/
    __init__.py
    models.py
    units.py
    dxf/
      __init__.py
      extractor.py
      mapping.py
    graph/
      __init__.py
      builder.py
      elevation.py
      fittings.py
      riser.py
    sdf/
      __init__.py
      parser.py
      writer.py
      template.py
    render/
      __init__.py
      isometric.py
    validate/
      __init__.py
      rules.py
      compare.py
    io/
      __init__.py
      csv_io.py
      json_io.py
    cli.py
tests/
  fixtures/
    dxf/
    sdf/
    csv/
  unit/
  integration/
configs/
  layer_map.csv
  block_map.csv
  elevation_rules.csv
  pipeline.yaml
outputs/
  .gitkeep
docs/
  mapping_review.md
  dxf_extraction_review.md
```

Existing server files may remain in place. New conversion logic should live in `src/pipenet_converter/` and be testable without starting the Flask server.

## 3. Domain Rules

- Treat the sprinkler network as a graph, not as a visual drawing.
- The authoritative intermediate representation is:
  - Node
  - Pipe
  - Nozzle
  - Fitting
  - Equipment
  - Valve
- Plan drawings provide X/Y positions and connectivity candidates.
- Section drawings and system diagrams provide elevation, riser, floor, alarm valve, and vertical connection context.
- System diagrams are usually schematic. Do not match them to plans by raw coordinates. Match by riser ID, floor, zone, diameter, alarm valve, grid location, and drawing labels.
- PipeNet SDF `Position x/y` is display-oriented. Calculation values must be explicit on pipes and links:
  - `length`
  - `rise`
  - `bore`
  - `roughness-or-c`
  - `input`
  - `output`
- Preserve the difference between:
  - real CAD coordinates: `x_real`, `y_real`, `z_real`
  - PipeNet display coordinates: `x_pn`, `y_pn`
- Split pipes at:
  - branch tees
  - diameter changes
  - valves
  - alarm valves
  - check valves
  - pressure reducing valves
  - vertical pipe start/end
  - head connections
  - calculation-relevant loss elements
- Orient pipes from supply/root toward downstream heads before SDF writing.
- Fit fitting counts to a consistent rule:
  - degree >= 3 node: tee candidate
  - degree == 2 with angle change: elbow candidate
  - connected pipes with different diameters: reducer or split point
  - valve symbols: gate/check/equipment depending on mapping

## 4. Units

- CAD DXF coordinates are assumed to be millimeters unless configured otherwise.
- Internal real coordinates:
  - `x_real`: millimeters
  - `y_real`: millimeters
  - `z_real`: meters, when sourced from PipeNet-like elevation values
- Pipe lengths written to SDF must be meters.
- Pipe rises written to SDF must be meters.
- Pipe bore written to SDF must be meters.
- Diameter text mapping:
  - `25A` -> `0.025`
  - `32A` -> `0.032`
  - `40A` -> `0.04`
  - `50A` -> `0.05`
  - `65A` -> `0.065`
  - `80A` -> `0.08`
  - `100A` -> `0.1`
  - `125A` -> `0.125`
  - `150A` -> `0.15`
  - `200A` -> `0.2`
- Default C-factor is `120` unless configured.
- Default nozzle flow may be configured. Do not hard-code project-specific design flow outside config or fixtures.

## 5. Required Python Version

Use Python 3.11 or newer.

Code style requirements:

- Use type hints for public functions and dataclasses.
- Use docstrings for public modules, classes, and functions.
- Prefer small, testable functions.
- Avoid hidden global state.
- Keep parsing, graph construction, validation, rendering, and SDF writing in separate modules.

## 6. Package Dependencies

Preferred baseline dependencies:

- `ezdxf` for DXF parsing
- `pydantic` or `dataclasses` for structured models
- `lxml` or `xml.etree.ElementTree` for SDF XML parsing/writing
- `pandas` for CSV review tables when useful
- `PyYAML` for pipeline config
- `networkx` for graph validation and traversal when useful
- `matplotlib` for isometric PNG rendering
- `pytest` for tests
- `ruff` for linting and formatting if added to the project

Do not add heavyweight dependencies unless a step explicitly requires them.

## 7. Testing Commands

Expected commands once the package skeleton exists:

```powershell
python -m pytest
python -m pytest tests/unit
python -m pytest tests/integration
python -m pip install -e .
python -m pipenet_converter.cli --help
```

If `ruff` is configured:

```powershell
python -m ruff check src tests
python -m ruff format src tests
```

Every implementation step must include at least one focused test or a stated reason why a test is deferred.

## 8. Forbidden Approaches

- Do not parse DWG directly.
- Do not require AutoCAD COM automation.
- Do not require a licensed CAD application.
- Do not use OCR in version 1.
- Do not attempt to understand arbitrary drawings without mapping files.
- Do not create a full architectural BIM model as the primary goal.
- Do not treat DXF visual geometry as automatically calculation-ready.
- Do not write SDF XML directly from raw DXF entities without an intermediate validated network model.
- Do not mix Flask/server UI code with the core conversion package.
- Do not hard-code a single project's layer names into core logic. Put project-specific mapping in CSV/JSON/YAML config.

## 9. Definition of Done

A feature is done when:

- It has a focused, typed implementation.
- It has unit tests or integration tests covering the intended behavior.
- It reads and writes structured models rather than ad hoc strings.
- It documents assumptions and unsupported cases.
- It preserves units explicitly.
- It produces inspectable intermediate outputs when relevant.
- It fails with clear validation errors instead of silently guessing.
- It does not require licensed CAD software.
- It does not regress existing SDF parsing/writing fixtures.

For the end-to-end converter, done means:

- DXF-derived plan data can produce nodes, pipes, nozzles, and fittings.
- Elevation rules can assign Z values.
- System/riser data can be merged into the graph.
- A check isometric PNG can be rendered.
- CSV review files can be exported.
- SDF XML can be written from the validated network model.
- The generated SDF can be compared against an existing reference SDF for counts, lengths, diameters, fittings, and connectivity.

## 10. Incremental Implementation Policy

Implement in small, reviewable increments:

1. Create project skeleton and package metadata.
2. Define data models before parsers.
3. Build SDF parser before SDF writer.
4. Build validators before trusting writer output.
5. Build DXF extraction with mapping files.
6. Build 2D graph construction.
7. Add elevation rules and riser/system integration.
8. Add fitting detection.
9. Add isometric rendering and CSV exports.
10. Add CLI pipeline.
11. Add regression tests against real SDF fixtures.
12. Add project-specific mapping reviews only after generic logic is stable.

Each step should leave the repository in a runnable state. Avoid large rewrites. Prefer explicit intermediate CSV/JSON outputs so domain experts can review extraction quality before SDF generation.
