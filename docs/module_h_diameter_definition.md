# Module H: drawing-first diameter definition

Implemented 2026-09-28. H explicitly opts into `drawing_first_v1`; F retains its
legacy diameter priority and UI. Source graph coordinates/connectivity are not
modified by this feature. It does not certify a fire-protection design.

## Definition order

1. Read the mapped drawing's native diameter texts (including text direction).
   H uses `configs/h_diameter_labels.json`: one permitted nominal size with an
   optional configured prefix. Mixed sizes such as `150 x 100`, quantities and
   unknown prefixes are not reduced to their first number.
   Invalid disposable text caches are reread from DXF in H. Read failures stop
   interpretation instead of silently substituting rule diameters.
2. Associate text with an actual graph edge. An explicitly linked DXF LEADER
   arrow endpoint takes priority (25 mm snap); otherwise use the native text
   bounds centre, original rotation and competing-candidate checks. The original
   insertion coordinate, raw text, layer and instance handle remain provenance.
3. Extend a uniform representative size along its continuous physical run:
   degree-two bends continue; at tees only a unique opposed pair continues.
   Supply/valve/head landmarks and ambiguous edges stop propagation. Side branches
   and the next head-to-head interval are not flooded by a single annotation.
4. Compare repeat candidates by ordered head types, endpoint connectivity and
   head-to-head segment lengths (15% tolerance) and normalized bend shape (8%,
   configurable). Require the same physical, valve-separated connected component.
   CAD fragmentation alone does not change correspondence. Each interval is
   transferred independently: a missing middle slot does not discard the other
   four known slots. Snapshot donors before copying: generated repeat values
   cannot become new donors in this pass. Direct target text always wins.
   Competing representatives or drawing
   contradictions are left unresolved. Search includes retained, non-selected
   drawing branches; it does not reintroduce geometry deleted by work-area crop.
5. No H unknown pipe is automatically filled from a head-count rule or default,
   including tree pipes and generated vertical connections. Head-count rules are
   retained only for comparison/review; they are not original drawing evidence.
6. Explicit saved user values are applied last with source/previous value retained.

Nominal drawing sizes (A/DN, mm) are distinct from actual inner diameter (mm).
The selected material's existing authoritative SLF supplies the latter. SDF bore
conversion remains in the established writer. Unknown sizes use `None` in the
core; the legacy table adapter's `0` is a blocked draft marker, never a valid bore.

## H user workflow

- After region/network definition, entering attributes automatically constructs
  the input tables. No mandatory extra table-confirmation click is needed.
- The small **관경 정의** control opens a movable panel. The H-only review-default
  diameter input is hidden and ignored by the H backend.
- Interpretation progress remains event-driven. The input view now preserves the
  white/red calculation network above the subdued original drawing; the legacy
  blue/purple/amber inference strokes are no longer painted over it. Selection
  shows links to actual native diameter text, and `근거 표시` shows all such links.
  See [pipe text references](module_h_pipe_text_reference.md) for the shared
  plan/isometric context-menu action and persistent undo/redo contract.
- Use click, pen/lasso, or rectangle selection. Shift adds/removes items. Filters
  distinguish pipes and heads. Area selection selects whole existing elements;
  it never cuts a pipe at the drawn selection boundary.
- Link a selected source text to selected pipes without retyping its value.
- Select a sized, connected, non-branching representative path, copy with Ctrl+C,
  select a target path and use Ctrl+V. Correspondence is a **manual preview** by
  normalized path length, even for different shapes. Reverse direction if needed.
  Where a target crosses a source size boundary, choose a value explicitly.
  Geometry, lengths, heads and fittings are never copied or moved.
- Drawing/manual values are protected on link/paste unless overwrite is explicitly
  checked. Bulk direct editing is an explicit user override. DN choices are
  checked against each pipe's material; a mixed invalid batch writes nothing.
- Bulk head K and minimum pressure are carried into nozzle rows and integrated
  sizing. Export creates named copies of the original SLF nozzle definition and
  binds only affected heads; original library entries remain untouched.
- Generated vertical connections are separately selectable. Continuous merged
  pipes update all underlying stable keys, not a fragile display pipe label.
- Corrections are written through to the existing per-drawing override JSON.
  The last 30 batch operations can be undone in the current session. Reopening
  restores values; the batch undo stack itself is not persisted across restart.

## Safety and integrated sizing

- Ordinary SDF cannot export unresolved sizes or conflicting diameter evidence.
- Inverse sizing remains **integrated-network only**. H defaults to unknown sizes
  with known drawing/user sizes locked; a separate option requests whole-network
  re-review. Existing completeness, fitting-loss and hydraulic solver checks remain.
- Editing invalidates the previous input table/export until a successful rebuild.
  Revision checks reject stale selections, copy previews and concurrent saved edits.
- Incompatible existing fitting/geometry edit history is not silently discarded
  during H's automatic rebuild. It remains on disk, is flagged for review, and
  prevents downstream use until resolved in the existing network editor.
- No OCR, arbitrary layer discovery, nearest-text certainty, or geometric change
  is introduced. Existing imported mappings still define the drawing scope.

## Explicit current limits

- Automatic repetition handles tee-side chains with head landmarks, not arbitrary
  branching subgraph isomorphism. A complex branched representative is copied
  path by path. Different shapes always require the manual preview.
- Ambiguous crossings, competing text, missing diameter boundaries, or incompatible
  representatives remain unresolved. A uniform run assumption is visible evidence,
  not proof of adequate flow/pressure or code compliance.
- H native metadata includes nested INSERTs (up to 12 levels), visible ATTRIBs
  and constant ATTDEFs. F retains its opt-out behavior. MLEADER and geometric
  guesses for unassociated leader lines are not supported. The selected/mapped
  text list still limits candidates; enrichment does not discover arbitrary layers.
- Coincident different text entities, multiple leader endpoints and missing
  native metadata cannot become automatic sources. Texts without uniquely owned
  physical intervals remain unresolved. Repeated two-main loop branches require
  aligned physical direction, independent of node numbering.
- SLF custom nozzle K uses L/min/sqrt(bar); user minimum pressure is gauge bar.
  SDF/SLF serialization uses absolute Pa (+101325), consistent with the file format.
  PIPENET desktop opening/calculation is a separate acceptance check, not implied
  by XML round-trip tests.

## Verification

The 2026-09-29 segment-level update supersedes the original deployment counts
below. See [segment reference v2](module_h_segment_reference_v2.md) for current
tests, before/after reports, the reduced unsafe coverage and remaining ambiguity.

Focused tests cover drawing-vs-rule priority, bent 100A main/tee isolation,
head-based repetition with fragmented CAD segments, conflicting donors, all-head
tree/loop/grid builds, persistence/rebuild/undo, compacted stable keys, stale writes,
head property SDF/SLF round trips, sizing scope and preserved edit history.
Browser tests exercise automatic build, Shift selection, bulk edit/undo, copy
preview, progress overlay and the LAN HTTP operation-ID fallback.

Deployment verification: 218 focused/core regression tests passed (3 existing
optional cases skipped); 24 H browser tests and 2 F sizing/continuous-run browser
tests passed. A read-only B1F historical-snapshot check passed: 527 native text
directions enriched, 472 owned annotations and 67 repeated edges. These counts
are an extraction diagnostic, not a claim that all remaining edges are resolved.

Read-only private drawing check (never rebuilds or overwrites user caches):

```powershell
$env:MODULE_H_REAL_BORE_CHECK='1'
python -m pytest tests/test_module_h_diameter_real_drawing.py -s
```

An invalid display-cache stamp can be read explicitly as a historical test
snapshot, reported as such. This is not a claim that the current extraction was
rebuilt or that every real drawing annotation has been manually validated.
