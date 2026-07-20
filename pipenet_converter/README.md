# pipenet-converter

Python package for converting fire protection sprinkler CAD-derived data into a PipeNet-compatible SDF file.

The project focuses on calculation-ready 3D sprinkler pipe centerline graphs, not full architectural BIM models.

Initial scope:

- DXF-derived plan extraction
- CSV/JSON mapping inputs
- elevation/riser rule integration
- intermediate Node/Pipe/Nozzle/Fitting/Equipment/Valve models
- isometric check rendering
- PipeNet SDF XML export
- validation and regression comparison

This package is scaffolded first; implementation is added incrementally.

## Installation

Use Python 3.11 or newer.

```powershell
cd C:\Users\admin\PycharmProjects\JupyterProject\pipenet_converter
python -m pip install -e .[dev]
python -m pytest
```

The converter reads DXF, CSV, JSON, YAML, and SDF files. It does not read DWG
directly. Convert DWG drawings to DXF before running the pipeline.

## Folder Setup

Recommended project folders:

```text
data/input/
  plan drawings as DXF
  layer_map_daiso_4f.csv
  block_map_daiso_4f.csv
  elevation_rules_daiso_4f.csv
  system_edges_daiso_4f.csv
  connection_map_daiso_4f.json
  template/reference .sdf

configs/
  daiso_4f_warehouse.yaml

outputs/
  generated CSV, PNG, SDF, validation, comparison, and report files
```

An example YAML config is provided at:

```text
configs/daiso_4f_warehouse.yaml
```

## Filling Mapping Files

Start from the templates in `data/sample/`:

- `layer_map_template.csv`
- `block_map_template.csv`
- `elevation_rules_template.csv`
- `system_edges_template.csv`
- `connection_map_template.json`

For `layer_map.csv`, classify only true sprinkler pipe centerline layers as
`pipe`, `main_pipe`, `branch_pipe`, or `riser`. Put architectural, dimensions,
hatches, and title layers as `background` or `ignore`.

For `block_map.csv`, classify sprinkler heads, valves, risers, alarm valves, and
equipment blocks. Unknown blocks should stay visible in review outputs until a
human confirms them.

For `elevation_rules.csv`, enter Z values from section drawings. Do not guess
elevations from visual 3D rendering.

For `system_edges.csv`, enter riser/system diagram edges manually. Treat the
system diagram as schematic and use this CSV as the authority for vertical
lengths and elevations.

## Running From YAML Config

After filling the input files, run:

```powershell
python -m pipenet_converter.cli run-config --config configs/daiso_4f_warehouse.yaml
```

Use verbose logging if needed:

```powershell
python -m pipenet_converter.cli run-config --config configs/daiso_4f_warehouse.yaml --verbose
```

Use strict mode to fail when validation `ERROR` issues exist:

```powershell
python -m pipenet_converter.cli run-config --config configs/daiso_4f_warehouse.yaml --strict
```

## Reading Outputs

Typical output files:

- `raw_segments.csv`: extracted pipe-like CAD line segments
- `raw_blocks.csv`: extracted CAD block inserts
- `raw_texts.csv`: extracted diameter text
- `extraction_warnings.csv`: unknown layer/block and missing extraction warnings
- `network_3d_nodes.csv`: generated network nodes
- `network_3d_pipes.csv`: generated pipe links
- `network_3d_nozzles.csv`: generated sprinkler nozzles
- `network_3d_fittings.csv`: generated tee/elbow/gate/check fitting counts
- `network_3d_equipment.csv`: equivalent-length equipment such as alarm valves
- `validation_report.csv`: structural validation issues
- `isometric_check.png`: visual check of the generated centerline graph
- `generated_pipenet.sdf`: generated PipeNet-style SDF
- `compare_report.csv`: comparison against reference SDF if configured
- `redline_items.csv/json`: machine-readable redline points if comparison issues exist
- `run_summary.txt`: high-level run summary

Do not treat a generated SDF as hydraulically correct until it opens in PipeNet
and the hydraulic calculation runs and converges.

## Real SDF Sample Regression

To run the optional real-sample regression test, place the PipeNet SDF file at:

```text
data/sample/1-1. 다이소 세종허브센터 지상4층 창고.sdf
```

Then run:

```powershell
python -m pytest tests/test_real_sdf_regression.py
```

If the file is absent, the test is skipped. When present, the test parses the
SDF, exports intermediate CSV tables, writes a new SDF, parses it again, and
checks that node, pipe, and nozzle counts are preserved.

## Real Project Workflow

1. Convert DWG drawings to DXF manually. The converter does not read DWG.
2. Copy `data/sample/layer_map_template.csv` to your project folder as `layer_map.csv`.
3. Fill `layer_map.csv` using the real CAD layer names.
4. Copy `data/sample/block_map_template.csv` to `block_map.csv`.
5. Fill `block_map.csv` using real sprinkler head, valve, alarm valve, and riser block names.
6. Copy `data/sample/elevation_rules_template.csv` to `elevation_rules.csv`.
7. Fill `elevation_rules.csv` from section drawing elevations and known pipe/head heights.
8. Copy `data/sample/system_edges_template.csv` to `system_edges.csv`.
9. Fill `system_edges.csv` from the system diagram and riser path.
10. Copy `data/sample/connection_map_template.json` to `connection_map.json`.
11. Set `connection_map.json` so the system/riser outlet connects to the plan node.
12. Run DXF extraction:

```powershell
python -m pipenet_converter.cli extract-dxf --dxf plan.dxf --layer-map layer_map.csv --block-map block_map.csv --out-dir outputs/extracted
```

13. Build the plan-derived network:

```powershell
python -m pipenet_converter.cli build-network --dxf plan.dxf --layer-map layer_map.csv --block-map block_map.csv --elevation-rules elevation_rules.csv --out-dir outputs/network
```

14. Merge the system/riser path:

```powershell
python -m pipenet_converter.cli merge-system --network-dir outputs/network --system-edges system_edges.csv --connection-map connection_map.json --out-dir outputs/merged
```

15. Render the check isometric:

```powershell
python -m pipenet_converter.cli render-iso --network-dir outputs/merged --output outputs/isometric_check.png
```

16. Write the PipeNet SDF:

```powershell
python -m pipenet_converter.cli write-sdf --network-dir outputs/merged --template template.sdf --output outputs/generated.sdf
```

17. Open `outputs/generated.sdf` in PipeNet and run the calculation check.

## Troubleshooting

- Missing input file: check the exact path passed to the CLI. File loaders report the missing path.
- Missing CSV column: compare your CSV header with the matching template in `data/sample`.
- Empty extraction: inspect `extraction_warnings.csv`. Common causes are unmapped pipe layers, XREF geometry not bound into the DXF, or using DWG instead of DXF.
- Unknown layers or blocks: update `layer_map.csv` or `block_map.csv`, then rerun `extract-dxf`.
- Invalid diameter label: use supported labels only: `25A`, `32A`, `40A`, `50A`, `65A`, `80A`, `100A`, `125A`, `150A`, `200A`.
- Invalid elevation rule: verify `priority` is an integer, `z_m` is numeric, and `match_field` is `node_type`, `source`, `metadata.<key>`, or `default`.
- Duplicate IDs: node and pipe IDs must be unique before SDF export.
- Strict mode: add `--strict` to fail the CLI with exit code `1` when validation `ERROR` issues exist.
- Verbose logging: add `--verbose` after the command name to print progress diagnostics.
