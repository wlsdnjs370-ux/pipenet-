# Module H segment reference v2 — 2026-09-29

This is source interpretation, not a hydraulic adequacy certification. The
public policy remains `drawing_first_v1`; the internal algorithm stamp is
`segment_reference_v2`. F does not opt into native anchors or partial repetition.

## Changed behavior

- Match native drawing text at its bounds centre, or an explicitly associated
  LEADER endpoint. Keep original coordinates/rotation/text/layer/entity identity
  for review. UI upright labels do not change analysis rotation.
- Divide interpretation at head landmarks. A missing slot stays missing; known
  intervals on the same repeated branch can still be transferred independently.
- Verify component, endpoint roles, ordered head kinds, lengths and bend shape.
  Freeze original donors before transfer. Never recursively manufacture evidence.
- Direct target text wins. Conflicting donors, blocked fragments and missing
  native identity stay unresolved and block ordinary export.
- Keep every source link in a compacted pipe's evidence chain. Existing manual
  source changes, library inner-diameter conversion and undo/redo remain intact.
- No coordinates, connectivity, lengths, head positions, original DXF files or
  saved graph caches are changed by interpretation/audit.

## Read-only saved-drawing measurements

Reports are in `outputs/h-diameter-tracing-20260929/`; audit files are created
exclusively and cannot overwrite an earlier baseline.

| B1F full saved graph | Before | After |
|---|---:|---:|
| Physical XY edges | 22,474 | 22,474 |
| Native candidate texts | 527 | 530 |
| Owned annotations | 472 | 443 |
| Repeated reference edges | 67 | 362 |
| Resolved physical edges | 6,803 | 3,398 |
| Unresolved physical edges | 15,671 | 19,076 |

The physical XY edge sets are identical despite a different integer-index graph
revision. New source-backed matches resolve 1,256 previously unknown edges.
Stricter head-interval boundaries and ownership checks leave 4,661 formerly
filled edges unresolved. 252 still-resolved values differ; these are retained
with both evidence records in `comparison-verified.json` for drawing review, **not**
claimed as 252 verified corrections. These are full drawing graph counts, not
selected-network table row counts or a manually labeled accuracy benchmark.

Daemyeong unit-plan saved graph: 67 candidate texts, 16 owned annotations,
62 resolved / 1,699 physical edges, 51 ambiguous annotations and no automatic
repetition. This drawing still needs ambiguous source ownership review; the
algorithm does not claim to solve arbitrary repeated apartment subgraphs.

Run reproducibly without starting Flask or rebuilding a user cache:

```powershell
python -X utf8 scripts/audit_h_diameter_tracing.py --key '<saved drawing key>' --output outputs/new-audit.json
python -X utf8 scripts/audit_h_diameter_tracing.py --compare outputs/before.json outputs/after.json --output outputs/comparison.json
```

Use `--historical` only to explicitly inspect an invalid-stamp retained snapshot;
the report flags its use. All reports above used current saved snapshots.

## Validation contracts

Focused synthetic tests cover missing middle slots, conflicting partial donors,
target priority, fragment conflicts, frozen donors, disconnected and valve
boundaries, different head kinds/bends, loop reversal, global rotation and node
renumbering, nested attributes, explicit leaders, source ambiguity, strict text
grammar, F opt-out, persisted native anchors and source geometry invariance.

Browser checks cover plan/iso selection and dotted evidence, compacted multiple
sources, manual reference patch without full reload, authoritative state parity,
Ctrl+Z/redo, Esc and bulk attributes. Existing SDF length and nozzle-library
round trips remain regression checks, not a PIPENET desktop acceptance test.

Final focused run: 169 passed, 2 opt-in saved-drawing checks skipped; 23 additional
browser tests passed. The empty-evidence panel now safely renders without a
stale scroll target. `verified.json` is the final B1F run; `before.json` is the
pre-change baseline. This does not claim a full-repository test pass.

With the user's saved-work approval, port 5051 was restarted and returned HTTP
200. The authenticated, read-only live H/F smoke check passed without drawing
writes or browser errors. Port 5052 was left running unchanged.
