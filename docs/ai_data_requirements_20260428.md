# AI Data Requirements

## Goal

This document defines which data each AI-assisted module needs, what is already available in the repository, and what remains missing.

## Current AI Modules

### 1. Triangle Head Detection

Purpose:
- detect triangular sprinkler-head symbols from DXF-rendered plan views

Data required:
- cropped symbol images
- bounding boxes in YOLO format
- variations in rotation, scale, line thickness, clutter, and layer noise

Current status:
- available

Current local data:
- `data/triangle_head_dataset`
  - `700` train images
  - `140` validation images
- `data/head_templates/*.png`
- multiple local DXF files under `data/uploads`

What was done:
- trained `YOLO11n`
- deployed weights to `models/triangle_head_yolo_ai/weights/best.pt`

### 2. Valve Symbol Detection

Purpose:
- detect alarm valves, check valves, gate valves, PRVs, and related control symbols

Data required:
- symbol crops or full-sheet annotations
- class-separated labels:
  - `valve_alarm`
  - `valve_check`
  - `valve_gate`
  - `valve_prv`
  - `valve_misc`
- multiple office drawing styles

Current status:
- missing

Minimum dataset target:
- `1,500+` labeled valve instances
- at least `200+` instances per major class

### 3. Pipe / Riser / Equipment Detection

Purpose:
- improve CAD-derived extraction for noisy sheets and raster-heavy drawings

Data required:
- line- or mask-level annotations for:
  - branch pipe
  - cross-main
  - riser
  - equipment boundary
- difficult negatives such as dimensions, grids, hatch, and unrelated MEP lines

Current status:
- missing

Minimum dataset target:
- `300+` annotated sheets for detection
- `100+` sheets with segmentation masks if SAM-style supervision is desired

### 4. OCR for Design Labels

Purpose:
- read diameter, floor, elevation, room-name, and system tags

Data required:
- text-box annotations or a closed-vocabulary synthetic dataset
- vocabulary groups:
  - diameters: `25A~200A`
  - floors: `B3~40F`
  - system tags: `HSP/MSP/LSP/LLSP`, `AV`, `PV`, `PRV`, `FX`
  - elevations and pressure labels

Current status:
- partially available
- synthetic closed-vocabulary dataset generated locally
- lightweight classifier trained and deployed
- `paddleocr` is still not installed, so open-vocabulary OCR remains unavailable

Practical next dataset:
- synthetic text render dataset for closed vocabulary
- `20k~50k` generated label images with font, noise, and skew variation

Implemented closed vocabulary:
- diameters: `25A~200A`
- floors: `B1~B6`, `1F~40F`
- tags: `HSP`, `MSP`, `LSP`, `LLSP`, `AV`, `PV`, `PRV`, `FX`, `ESFR`, `RTI`, `QR`

### 5. Room / Zone Segmentation

Purpose:
- split plan images into rooms, corridors, shafts, and excluded spaces

Data required:
- polygon masks per region
- category labels aligned with NFTC / Hanbaek exclusion logic

Current status:
- missing

Minimum dataset target:
- `200+` floor plans with polygon masks

### 6. Head Placement Optimization

Purpose:
- optimize head layout beyond fixed grid sweep

Data required:
- zone polygons
- obstacle polygons
- accepted human layouts or accepted simulation outputs
- cost and violation outcomes

Current status:
- partially available as structured geometry inputs
- no supervised target dataset yet

Practical data source:
- historical accepted layouts exported from the server
- synthetic optimization traces generated from rule-constrained simulation

### 7. Pipe Routing / Diameter Recommendation

Purpose:
- recommend branch grouping, routing, and pipe-size choices

Data required:
- graph inputs
- accepted routed networks
- PIPENET pass/fail outcomes
- hydraulic metrics by route

Current status:
- missing as a training dataset

Practical data source:
- archived SDF + CAD pairs
- validated scenario outputs from `pipenet_validator.py`

### 8. Anomaly / Imbalance Detection

Purpose:
- flag suspicious hydraulic patterns before manual review

Data required:
- per-scenario metrics:
  - pressure
  - flow
  - velocity
  - imbalance
  - duration
- labels for pass / warning / fail

Current status:
- feature source exists
- training corpus not yet accumulated

Practical data source:
- batch export from existing validation runs

## What I Can Build Directly Now

Without outside annotation labor or new licensed data, I can directly work on:

1. triangle head detection
2. synthetic OCR closed-vocabulary data generation
3. anomaly-detection training from accumulated validator outputs
4. optimization trace collection for future head-placement learning

## What Still Requires User-Supplied or Team-Supplied Data

These cannot be made reliable from the current repository alone:

1. valve multi-class detection
2. room polygon segmentation
3. storage / ESFR-specific symbol detection
4. pipe routing oracle supervision
5. OCR on real raster scans across office-specific title/text styles

## Recommended Data Build Order

1. valve detection dataset
2. OCR closed-vocabulary synthetic dataset
3. validator-run feature lake for anomaly detection
4. room/shaft exclusion segmentation set
5. routed-network supervision set
