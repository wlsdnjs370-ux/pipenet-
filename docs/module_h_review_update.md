# Module H review/workbench update — 2026-09-28

## Scope

Module H only opts into the new compact controls, property window and review
export. Module F retains its interface and normal diameter export validation.
Shared fixes cover ordered system/machine-room requests, reader/writer handoff,
reuse of validated parsed geometry, and canonical property-edit transactions.

## Performance and data preservation

- A request-local prepared World can be reused by pick → network construction
  only when drawing key, absolute source path/size/mtime and crop regions match.
  It is never published as an uncropped cache or reused by another request.
- H annotation fallback reads native TEXT/MTEXT and INSERT transforms without
  constructing unrelated drawing geometry. Binary DXF keeps the original reader.
  Encoding, block base points, nested/mirrored transforms and OCS stay with ezdxf.
- Native text results are shared with orientation enrichment for that source
  revision; the authoritative mapped/cropped text list still controls candidates.
- Local B1F source measurement: full native text loading 37.49 s; selective
  loading 8.19 s; all 9,412 native text records identical. After the saved working
  crop, all 527 candidate diameter points matched the previous extractor exactly.
  These are annotation measurements, **not total workflow timing guarantees**.
- System and machine-room extraction waits for queued point/layer previews.
  A bounded 2-second handoff waits for short readers, never a competing writer.
  Results from an obsolete drawing/slot cannot overwrite the current view.

## User-facing changes

- Region selection starts disarmed and uses a prominent toggle button. Selecting
  the start-node tool disarms region selection. Labels use 시작노드 정의,
  인식 결과 and 배관 속성 확정.
- Technical completion messages, long loop/review prose, rectangle zoom buttons
  and merge diagnostic tables are hidden in H. Actionable errors remain visible.
- Representative A / target B correspondence has a persistent table and canvas
  links, both during real computation and after completion. Events are bounded;
  a dropped-event gap is repaired from the actual accumulated diameter snapshot.
  Missing or conflicting evidence is never shown as a confirmed assignment.
- The integrated property window opens with the merged graph. Row selection and
  graph selection synchronize in both directions. Material, DN, C, length, node
  position and head settings use the existing atomic preview/apply/revision/undo
  engine. Length changes have the existing downstream geometry semantics, not
  an isolated table-only override. Real inner diameter follows the library.
- K-factor and minimum pressure are explicit nozzle overrides; subsequent flow
  edits retain them. SDF/SLF use a separate bound nozzle definition, preserving
  the library original. Updating an existing nozzle retains its identity and
  active/inactive status, including previously defined junction nozzles; creating
  a new nozzle still requires a valid terminal node. Table display rounding does
  not round untouched stored coordinates/properties. Old integrated
  edit/inspector panels are hidden.

## Output boundaries

- H plan-only output is explicitly a non-calculable **review SDF**. Unknown or
  conflicting bore is serialized as unassigned (`bore=0`), not a guessed pipe
  diameter. A `.review.json` retains original values, identities and evidence.
  The source tables are deep-copied, not changed. Invalid graph references,
  non-finite coordinates/lengths and invalid C still fail validation.
- Plan-only output lacks the complete actual source/pump definition. Its title
  and UI state NOT CALCULATION READY; completing diameters and the hydraulic
  boundary conditions is required before a calculation is meaningful.
- Native desktop PipeNet import of an unassigned-bore review file has **not**
  been tested here. XML structure, units, source preservation and SLF bindings
  are covered by tests; calculation readiness is deliberately not claimed.
- A fresh feasible inverse-sizing result selects “역산 결과” as the H output
  basis. Both the main SDF action and download route through the separate sized
  proposal. “원본 배관망” remains an explicit alternative with its full checks.
  Edits/stale conditions invalidate proposal output; original tables stay intact.

## Verification

Focused tests cover selective DXF equivalence, request-local/crop isolation,
reader handoff, both extraction queues, H controls, live evidence/snapshot repair,
bidirectional table selection, edit/undo, nozzle overrides, plan review export,
and primary-action proposal export. Existing F graph/editor/SDF/sizing tests are
included in regression checks. Private drawing checks use read-only source and
saved display fixtures and do not rewrite user drawing or edit files.
