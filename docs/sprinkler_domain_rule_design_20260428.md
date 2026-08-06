# Sprinkler Domain Rule Design

## Scope

This note defines the missing domain-rule layer for the sprinkler conversion and validation package.

It is based on:

- the current repository state
- the provided dissertation PDF
- the provided Hanbaek sprinkler design HWP
- current official Korean fire-safety standards on `law.go.kr`

This is a design document, not the final implementation.

## Current Gap

The current codebase has almost no explicit implementation for:

- `HSP/MSP/LSP/LLSP` vertical pressure-zone logic
- Hanbaek `Case 1~5` topology logic
- `ESFR / 103B` system classification
- `RTI / 조기반응형` head-response classification

The existing package has graph, elevation, SDF parsing/writing, and generic validation modules, but it lacks a rule-classification layer that can decide:

1. what sprinkler system profile applies
2. what head family applies
3. what water-supply / pressure-zone topology applies
4. what legal and office-specific checks must be enforced

## Source Notes

### Dissertation PDF

Observed design facts from the provided PDF:

- domestic sprinkler design references `NFPC 103` / `NFTC 103`
- hydraulic design becomes dominant in taller buildings
- end-head pressure is discussed in the range `0.1 MPa` to `1.2 MPa`
- horizontal spacing reference values include `1.7 m`, `2.1 m`, `2.3 m`, `3.2 m`
- branch-pipe velocity limit `6 m/s`
- other-pipe velocity limit `10 m/s`
- high-rise zoning distinguishes `HSP`, `MSP`, `LSP`, and sometimes `LLSP`
- remote-zone and near-zone checks are both important
- `HSP/MSP/LSP/LLSP` are pressure / supply zones, not NFPA hazard classes
- NFPA occupancy classes remain useful as an analytical crosswalk:
  - `Light Hazard`
  - `Ordinary Hazard Group 1`
  - `Ordinary Hazard Group 2`
  - `Extra Hazard Group 1`
  - `Extra Hazard Group 2`

### Hanbaek HWP

Text extraction from the HWP succeeded partially. The extract clearly contains:

- sprinkler installation purpose
- water demand / water source discussion
- piping type and piping installation sections
- head spacing and head installation sections
- pre-action, deluge, and dry-pipe sections
- notes on NFPA-based hanger and valve handling concepts

Useful extracted points:

- dry systems can justify a `130%` design-area style increase for the domestic `10-head` case
- hanger design and valve operation constraints are operationally important
- head spacing logic and installation location rules are central

Important limitation:

- the extracted text did **not** expose literal `Case 1~5`, `RTI`, `ESFR`, or `103B` tokens
- that likely means those parts are image-based, table-based, or not text-extracted cleanly

So the Hanbaek case taxonomy should be treated as **partially confirmed** until the office guide is available as PDF or manually transcribed table data.

### Official Standards

Current official sources confirmed from `law.go.kr`:

- `스프링클러설비의 화재안전성능기준(NFPC 103)` `[시행 2026. 3. 1.]`
- `스프링클러설비의 화재안전기술기준(NFTC 103)` `[시행 2026. 3. 1.]`
- `화재조기진압용 스프링클러설비의 화재안전성능기준(NFPC 103B)` `[시행 2026. 3. 1.]`
- `화재조기진압용 스프링클러설비의 화재안전기술기준(NFTC 103B)` `[공고 2024. 7. 1.]`
- the law-book listing also shows the `NFTC 101~107` family as current standards metadata

Additional legal facts confirmed from official search snippets:

- `조기반응형 스프링클러헤드` is a named head category in sprinkler standards
- certain occupancies require quick-response heads
- `RTI` belongs to the sprinkler-head technical approval standard, which means RTI should be treated as **catalog / product metadata** rather than something inferred from geometry
- `NFPC/NFTC 103B` applies to early-suppression sprinkler installations, especially rack-storage / high-storage contexts

## Domain Model To Add

Add these concepts explicitly to the core model layer.

### 1. Building / Space Classification

- `building_use`
- `space_use`
- `occupancy_profile`
- `is_residential_room`
- `is_hospital_room`
- `is_lodging_room`
- `is_stage_area`
- `is_rack_storage`
- `storage_height_m`
- `ceiling_height_m`
- `freeze_risk`

### 2. System Classification

- `system_type`
  - `wet`
  - `dry`
  - `preaction_single_interlock`
  - `preaction_non_interlock`
  - `preaction_double_interlock`
  - `deluge`
  - `esfr`
- `legal_standard_family`
  - `nfpc_103`
  - `nftc_103`
  - `nfpc_103b`
  - `nftc_103b`

### 3. Pressure / Supply Zone Classification

- `pressure_zone_type`
  - `HSP`
  - `MSP`
  - `LSP`
  - `LLSP`
- `supply_mode`
  - `pump_direct`
  - `gravity_tank`
  - `pressure_tank`
  - `combined`
- `remote_zone_candidate`
- `near_zone_candidate`
- `requires_pressure_reducing_valve`

### 4. Head Classification

- `head_family`
  - `standard_closed`
  - `open`
  - `quick_response`
  - `esfr`
  - `residential`
  - `sidewall`
  - `dry_head`
- `response_class`
  - `standard_response`
  - `special_response`
  - `quick_response`
- `rti_value`
- `k_factor`
- `temperature_rating_c`
- `orientation`
  - `pendent`
  - `upright`
  - `sidewall`

### 5. Hanbaek Office Topology

- `hanbaek_case_id`
  - `case_1`
  - `case_2`
  - `case_3`
  - `case_4`
  - `case_5`
- `hanbaek_case_confidence`
  - `confirmed`
  - `inferred`
  - `unknown`

## Rule Pipeline

The rule engine should run in this order.

### Step 1. Applicability Gate

Decide whether the space belongs to:

- standard sprinkler (`103`)
- ESFR / early-suppression (`103B`)
- non-sprinkler fire system family

Decision inputs:

- rack-storage flag
- storage height
- ceiling height
- commodity / storage risk class if available
- office override

If `103B` applies, the engine must stop treating the space as a normal `103` branch for head selection.

### Step 2. System-Type Selection

Choose:

- wet
- dry
- pre-action
- deluge

Decision inputs:

- freeze risk
- stage / opening / exposure conditions
- equipment-room constraints
- special asset protection use
- office override

This decision changes:

- demand sizing adjustments
- valve/equipment requirements
- acceptable head families
- delay and reliability checks

### Step 3. Vertical Pressure-Zone Selection

Classify each served floor / riser branch into:

- `HSP`
- `MSP`
- `LSP`
- `LLSP`

This is not an NFPA hazard class. It is a pressure and supply topology class.

Decision inputs:

- floor elevation
- riser source
- pump discharge pressure
- gravity tank availability
- pressure-tank availability
- overpressure risk
- pressure-reducing valve placement

Derived checks:

- farthest zone minimum pressure and flow
- nearest zone maximum pressure
- branch velocity
- other-pipe velocity

### Step 4. Hanbaek Case Selection

Select office topology case from `Case 1~5`.

This must **not** be hard-coded in Python first. It belongs in a structured config table.

Suggested decision table inputs:

- system type
- pressure zone type
- water source type
- pump existence
- gravity tank existence
- pressure tank existence
- PRV existence
- single-zone vs multi-zone
- riser count
- building height band

The engine returns:

- `hanbaek_case_id`
- resolved topology diagram key
- required components list
- missing input fields if the case cannot be selected cleanly

### Step 5. Head Profile Selection

Choose the head profile.

Decision inputs:

- legal standard family
- room use
- rack-storage flag
- ceiling height
- freeze risk
- sidewall requirement
- product catalog data

Rules:

- `ESFR` head selection belongs to `103B`
- `조기반응형` is a quick-response requirement for specific occupancies
- `RTI` is not inferred from layout; it is loaded from approved head catalog data
- if quick-response is legally required and the selected product catalog entry is not quick-response by RTI/response class, validation must fail

### Step 6. Demand / Area / Count Rules

Determine the governing design-demand rule.

Possible modes:

- domestic standard-head-count mode
- hydraulic remote-area mode
- ESFR table-driven mode
- office-adjusted dry/preaction mode

Examples of constraints to encode:

- end-head minimum pressure checks
- end-head maximum pressure checks
- minimum discharge per head where applicable
- dry-system demand adjustment
- rack-storage ESFR pressure table lookup

### Step 7. Geometry and Equipment Validation

Apply geometry and network checks after system selection.

Checks include:

- head spacing
- wall clearance
- obstruction clearance
- branch / cross / main pipe role consistency
- pipe split at valves / reducers / tees / risers
- PRV location
- alarm-valve location
- dry / preaction valve placement
- supply-root orientation

### Step 8. Hydraulic Validation

Apply pressure and flow checks on the resolved graph.

Checks include:

- remote node pressure
- nearest node overpressure
- branch velocity `<= 6 m/s`
- other-pipe velocity `<= 10 m/s`
- diameter continuity and reducer correctness
- rise and elevation head correctness

## Data-Driven Design

Do not bury these rules in `if/else` blocks scattered across the codebase.

Use structured config files.

### Required Config Files

- `configs/head_catalog.csv`
  - manufacturer
  - model
  - legal_family
  - head_family
  - response_class
  - rti_value
  - k_factor
  - temperature_rating_c
  - orientation

- `configs/space_rule_map.csv`
  - space_use
  - quick_response_required
  - open_head_required
  - standard_family_default

- `configs/zone_rule_map.yaml`
  - floor bands
  - supply source logic
  - PRV thresholds
  - remote/near review policy

- `configs/hanbaek_case_map.yaml`
  - case id
  - predicate fields
  - required components
  - excluded combinations
  - notes

- `configs/esfr_pressure_table.csv`
  - standard version
  - max ceiling height
  - max storage height
  - head K-factor
  - orientation
  - min pressure

## Proposed Code Placement

Keep the current package testable without Flask.

### Extend Existing Files

- `pipenet_converter/src/pipenet_converter/models.py`
  - add enums and dataclasses for system, zone, head, case, and legal-profile metadata

- `pipenet_converter/src/pipenet_converter/config.py`
  - load YAML/CSV rule tables

- `pipenet_converter/src/pipenet_converter/pipeline.py`
  - insert a rule-resolution stage before graph validation/export

- `pipenet_converter/src/pipenet_converter/validator.py`
  - split generic graph checks from domain-rule checks

### Add New Files

- `pipenet_converter/src/pipenet_converter/system_classifier.py`
  - classify `103` vs `103B`, wet/dry/preaction/deluge, and pressure zone

- `pipenet_converter/src/pipenet_converter/head_selector.py`
  - resolve head family, response class, and catalog compatibility

- `pipenet_converter/src/pipenet_converter/domain_rules.py`
  - apply Hanbaek case logic and legal requirement checks

- `pipenet_converter/src/pipenet_converter/hydraulic_policy.py`
  - centralize pressure, flow, and velocity thresholds

- `pipenet_converter/src/pipenet_converter/rule_result.py`
  - structured result object with:
    - selected profiles
    - warnings
    - missing inputs
    - evidence trail

## Minimal Implementation Order

Implement in this order.

1. Add enums/dataclasses for:
   - legal family
   - system type
   - pressure zone
   - head family
   - response class
   - Hanbaek case
2. Add config loaders for catalog and rule tables.
3. Implement `103` vs `103B` applicability classifier.
4. Implement quick-response requirement classifier.
5. Implement RTI-based head validation against catalog data.
6. Implement `HSP/MSP/LSP/LLSP` zone classifier.
7. Implement Hanbaek `Case 1~5` selection from config table.
8. Add hydraulic-policy checks for pressure and velocity.
9. Export the resolved rule profile into reports and SDF review outputs.

## Validation Strategy

Tests must be data-driven and traceable.

### Unit Tests

- quick-response-required room classification
- `103` vs `103B` routing
- RTI and response-class validation
- pressure-zone classification
- Hanbaek case selection

### Integration Tests

- high-rise mixed-zone building with `HSP/MSP/LSP`
- rack-storage `103B` ESFR case
- residential quick-response case
- dry-system case with area/demand adjustment
- pre-action case with required valve metadata

## Important Design Constraint

`RTI`, quick-response, and ESFR are not purely geometric concepts.

That means the converter cannot derive them from DXF alone.

These must come from:

- head block metadata
- symbol mapping
- project design table input
- manufacturer/model catalog lookup
- office override

If those inputs are missing, the system must return:

- `unknown`
- `requires_user_mapping`

and must not silently assume a standard head.

## Next Practical Step

Before implementing the logic, the missing structured input tables should be prepared:

1. Hanbaek `Case 1~5` matrix
2. approved head catalog with `RTI`, `K`, orientation, and `ESFR` flags
3. space-use to quick-response requirement map
4. `103B` applicability table for rack-storage conditions

Without those tables, code can be written, but it will still be guessing.
