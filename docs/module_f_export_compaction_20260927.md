# Module F: compact SDF exports and smaller apparent symbols

## Scope

All Module F plan SDF, merged SDF and hydraulic-review SDF exports use an export-only
copy. Editor tables, saved edit history, hydraulic proposal IDs and source labels
are not changed. Generate new files to apply the new policy; existing downloads
are not rewritten. No PIPENET GUI verification is claimed.

Core implementation: `src/pipenet_converter/graph/export_compaction.py`.

## Reduction rules

- Combine only straight, degree-two, consistently oriented serial pipes with
  equal nominal/actual diameter, material, C, status and other physical fields.
- Preserve branches, bends, diameter/material boundaries, head attachment/stems,
  supply/pressure/flow boundaries, valve/pump links, explicit loss points,
  unresolved review points and merged-part connection anchors.
- Unknown/internal equipment placement protects the entire containing span.
  Opposing orientations, zero-length spans, non-finite values and conflicting
  diameter evidence are not simplified.
- Check real XYZ (XY mm, elevation m) AND supplied display views. Do not remove a
  corner just because an isometric projection happens to make it look straight.
- Sum declared lengths; never remeasure from display geometry. Aggregate cached
  equivalent lengths; retain each native fitting/equipment exactly once on its
  new host pipe. Explicit endpoint equipment retains its location.
- Never prune a spanning tree. Contraction of a serial vertex removes one node
  and one edge together, preserving connected components and cycle rank. Do not
  create a self-loop or collapse to an existing parallel edge.
- Keep the earliest surviving pipe label and all original rows in the audit.
  Plan/merged exports create a `.compaction.json` beside the SDF; merged ZIP
  includes it. Hydraulic-review JSON gains `export_compaction` after export and
  the diagnostic-download button returns this augmented copy.

## Drawing-only size policy

Known SDF templates do not supply a verified node/head glyph-size attribute.
Do not invent one. Use uniform drawing-coordinate expansion by **4**, without
changing physical pipe length/rise, elevation, diameter, flow or pressure.
This is a relative visual reduction, not an application-wide PIPENET preference.
Keep atmospheric nozzle display-tail length compensated by the same factor.

Normal merged iso previously used factor 3; it now uses 4 (approximately 25%
smaller apparent fixed glyphs). Plan exports and hydraulic-review exports share
the factor. Review exports now use the plan reference span as well, so a long
riser does not independently shrink the plan cluster. Merged KFP/HAS retain
their established display/physical-coordinate policies.

## Verification

`tests/test_module_f_export_compaction.py` covers serial reduction, source
immutability, diameter/material/C/control boundaries, actual and display bends,
vertical transitions, loss ownership, head stems, loop/grid cycle preservation,
solver equivalence, ordinary merged output, plan output and hydraulic review.

`scripts/audit_export_compaction.py` uses the previously saved read-only actual
network snapshot; it does not modify live drawings or start jobs in a user session.

Actual saved B1F integrated snapshot:

| Output | Physical nodes | Pipes |
|---|---:|---:|
| Before | 245 | 248 |
| Normal merged | 73 | 76 |
| Hydraulic sizing review | 89 | 92 |

Head count 12, independent cycle rank 4 and declared length 481.269771217 m are
preserved. Sizing uses more boundaries because of changed diameters and its
per-pipe total-loss representation. Atmospheric nozzle display outlet nodes
are additional SDF nodes, not included in the physical counts above.

The written review SDF was re-read and solved using its written supply pressure,
lengths and equipment losses. Maximum head flow difference was about 0.000109
L/min; maximum retained-node pressure difference was about 0.00000940 bar.
The small difference follows the existing integer-Pascal boundary serialization.
Long aggregated lengths/rises use 12 significant digits in review exports to
avoid the shared legacy writer's six-significant-digit rounding.

The test does not establish regulatory compliance, final pump selection or
successful recalculation in the installed PIPENET GUI. The proposal remains
NOT VALIDATED and does not automatically replace the original design.

Final targeted regression run: **252 passed, 3 skipped** (export reduction,
plan/merge exporters, fitting resolution, loop/grid extraction, solver, isolated
browser checks and length authority). A separate legacy real-drawing test,
`tests/test_module_f_design.py::test_design_라우트가_열려_있다`, remains blocked by
the B1F fixture's existing P69 diameter-evidence conflict. Both the pre-existing
writer and API reject this conflict; the safety check was not weakened to make
that fixture export. A 10,000-segment straight-chain reduction took about 0.44 s
locally (10,001 nodes to 2; no attachments in this synthetic performance case).

Deployment: restarted 5051 on 2026-09-28 after the user's explicit approval.
The new server PID is 87756 (previous PID 48224); authenticated `/module-f` and
its script both returned HTTP 200. The 5052 server (PID 4048) was not touched.
Reopen the saved drawing and regenerate SDFs; existing downloads are unchanged.
