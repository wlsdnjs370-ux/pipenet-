# Module H: canonical pipes and draft-segment reduction

The calculation table, editable graph and ordinary plan SDF use the same
canonical pipe labels and endpoints. CAD fragments are audit sources, not
additional exported pipes. Each original fragment belongs to exactly one
canonical pipe; compaction validates this partition and preserves the total
declared length.

## Reproduced cause

The 2026-09-30 plan snapshot contained 217 nodes, 220 pipes and 13 nozzles.
Its SDF compaction report removed no nodes. Two independent paths caused this:

- accepting a changed H basis committed the raw table as canonical without
  the usual continuous-run normalization;
- the conservative export guard treated every unspecified diameter like a
  conflicting diameter, preventing even plain collinear unmapped spans from
  joining.

Accepting a new basis now normalizes the accepted base in the same atomic
transaction. Rejected old edits are still archived, not replayed on unrelated
labels. The normalization command is undoable.

## Versioned safety rules

Existing version-1 history and Module F behavior remain unchanged. H creates
version-2 `compact_runs` commands. Only plain H `no_match` source spans with
absent bore, explicit `block_export`, matching material/C/physical fields and
finite positive lengths can join while unassigned. They remain unassigned and
blocked for calculation-ready export. Nothing copies a nearby pipe's bore.

The following still prevent contraction: branch/head/fitting/equipment nodes,
physical and display bends, elevation transitions, material/C/bore changes,
reversed directions, conflicting or partial diameter evidence, unknown losses,
generated head/riser spans and deliberate editor-created split points.

All underlying source edges, stable edit keys and diameter evidence remain in
`continuous_sources`. Current explicit references take precedence over historic
automatic evidence. The representative table counts calculation pipes only;
outside-scenario reference geometry is retained as evidence but not counted as
an output target. A pipe retaining multiple source labels appears in one combined
reference group, never once per label. Thus referenced + unreferenced = total
canonical pipes, and each table pipe corresponds to one ordinary SDF pipe.

## Verification

Focused tests cover draft reduction, refusal at meaningful boundaries, legacy
history replay, changed-basis acceptance, undo/redo, SDF/table identity and the
UI reference partition. `scripts/audit_h_plan_segments.py` obtains a read-only
snapshot; `scripts/verify_h_plan_segments.py` validates an offline copy.

For the retained snapshot: 217 → 65 nodes, 220 → 68 pipes (4 defined/reference
pipes and 64 unassigned pipes), 13 nozzles, 47 fitting rows, 4 independent cycles,
one connected component and 201.5749276 m of declared pipe length. All 220 source
fragments are accounted for exactly once. SDF values match the table within the
writer's existing six-significant-digit serialization. This is structural/draft
verification, not a completed hydraulic calculation or PIPENET GUI validation.
