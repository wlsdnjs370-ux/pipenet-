# Claude v4 Integration Notes

## What Was Adopted

The imported Claude bundle was **not** used as a full replacement.

Adopted modules:

- `nftc_rules.py`
- `hb_rules.py`
- `phd_rules.py`
- `auto_design.py`
- `change_log.py`
- `pipeline_orchestrator.py`
- `pipenet_validator_v4.py`
- `server_patch.py`

## What Was Kept From The Existing Server

The existing server remains the system of record for:

- Flask app boot and routing surface
- current HTML / CSS / JS UI
- existing `pipenet_validator.py`
- existing CAD-SDF comparison workflow
- existing upload / report / feedback flows

## Integration Strategy

The integration is additive.

- `server_patch.py` registers `/api/v4/*` routes
- the main Flask app is not restructured
- the existing validator is wrapped, not replaced
- no template or static bundle was replaced by the Claude package

## Why This Split Was Chosen

The Claude bundle is stronger in:

- rule classification
- pipeline orchestration
- v4 API surface

The current repository is stronger in:

- deployed server behavior
- existing user workflow
- current validation and reporting integration

So the repository now uses:

- Claude logic where it improves domain-rule coverage
- existing server logic where it already works and is already wired to the UI

## Known Follow-up Work

1. Normalize human-facing strings in the added v4 modules.
2. Convert Hanbaek `Case 1~5` into structured config instead of code-heavy branching.
3. Add tests around:
   - `103` vs `103B`
   - `HSP/MSP/LSP/LLSP`
   - quick-response / RTI validation
   - Hanbaek case selection
4. Expose selected `/api/v4/*` features in the existing UI only after response shapes are finalized.
