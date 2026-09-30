# Module F: loop/grid source extraction (phase 1)

## Available workflow

In **손질 → 배관망 방식**, explicitly choose `tree`, `loop`, or `grid`, then
run **물흐름**. The default for old saved drawings remains `tree`. Mode changes
invalidate the displayed flow/scenario, not source geometry or existing files.
Normal **손질 저장** persists the choice. Each drawing slot owns its own mode.

- `tree`: existing single-route extraction, head loads and calculation pipeline.
- `loop` / `grid`: preserve source-to-head alternate paths. The choice declares
  the drawing layout; it is **not** inferred automatically from cycle count.
  Loop mains and double-ended grid branch lines use the same general graph engine.
- **작동 헤드 후보 선정**: use the existing zone/K and reference-route distance
  ranking to select a scenario. This is not a hydraulically proven worst area.
- **경로 기록**: download `module-f-network.json`, containing the all-head network,
  the selected scenario, excluded-edge reasons, source IDs, and explicit pending
  calculation fields. The screen highlights the same preserved scenario edges.

## Algorithm and separation

`src/pipenet_converter/graph/network.py` is independent of Flask.

1. Reuse established supply/head attachment identities and the original length
   map. The legacy tree is a reference for reachability/ranking only.
2. Decompose the reachable source graph into biconnected edge blocks.
3. On the block/vertex incidence tree, prune leaves not needed by the supply or
   target heads. Keep all edges in remaining blocks, not just one route per head.
4. Repeat that selection on the same blocks for the operating-head scenario.
   Inactive nozzle stubs may be omitted; transfer branches still remain.

The result is the union of all simple source-to-target paths, computed without
enumerating those paths. A ring attached only at one articulation with no target
inside is omitted; a ring providing an alternate source-to-head route is kept.
Crossing coordinates never create a new connection. No coordinates, source edge
lengths, or original fitting ports are moved or overwritten.

`cycle_rank = E - V + 1` is an audit value for this connected extraction, not a
loop-versus-grid classifier. Physical layout classification would additionally
need reliable main/branch roles and double-ended branch attachment evidence.

## Explicit limitations / next increment

This is **2D source connectivity extraction**, not completed 3D hydraulic export.
The new branch does not run the tree-only planar chain layout, downstream head
count bore fallback, upstream monotonic sizing, or parent-based tee loss rules.
Calculation table creation, KFP/SDF emission and stale output downloads are
blocked in the new modes, including a preserved plan in an inactive merge slot.
Existing tree outputs and user files are not deleted.

The typed JSON distinguishes `plan_length_m` from unresolved `length_m`,
`rise_m`, `bore_m`, flow and elevation (all `null` until actually determined).
Endpoint A/B order is a stable reference, **not actual flow direction**.
Heads attached on retained transfer lines remain present with `selected=false`.
No cross fittings, pair-of-tees replacements, or synthetic hydraulic losses are
invented. A degree-four node still needs physical-port evidence/review.

### Head-symbol gap recovery (2026-09-23)

The stage-5 matcher already identifies same-material, opposed straight pipe arms
separated by a head symbol. The old `add_no_ring` gate then discarded a match
if its endpoints were already connected by another route. This removed genuine
loop closures **before** the preserved-network engine could see them.

The pipeline now carries those rejected stage-5 matches as `cycle_candidates`.
It does not change its legacy edges. Cache version 7 requires the evidence sidecar;
reopening an old drawing rebuilds its automatic cache once using its picked spec.
Saved manual edits are reapplied normally. No new layer names are hard-coded.

`src/pipenet_converter/graph/cycle_closures.py` supplies an additive projection
only for loop/grid modes. It rechecks endpoint coordinates, exterior arms and
the surviving return route. An existing head center splits the restored span,
so the result does not bypass the nozzle or create a false triangular loop.
Explicit deletions (including a half-neck) block recovery, and undo restores the
previous decision. Source nodes, original edge lengths and the tree source are
not mutated. Switching back to tree uses exactly the original tree behavior.

The projected graph feeds the displayed segments, hit testing/deletion, water
flow, candidate selection and review JSON. Screen counts show restored/review
candidates; `cycle_recovery` records in the JSON explain individual decisions.
Added edges are marked inferred connections, never native drawn pipes or
hydraulically validated links. Both loop and grid use this projection.

This is deliberately **not arbitrary gap healing**. Non-head stage-4 gaps,
unpicked DXF geometry, unexplained crossings, ambiguous centers or a return
route removed by the user still require manual review. The confirmed B1F
컨셉2-수정본 (1) replay restored all 25 evidenced stage-5 gaps; its 57,245 pipeline
edges and 58,739 coordinates matched the previous tree cache exactly.

The board is a simple undirected graph keyed by endpoint pairs: distinct parallel
physical pipes with the exact same endpoints need pipe-ID/multigraph support in
a subsequent increment. Non-return valves and equipment states also need explicit
directed/stateful treatment before hydraulic solving.

### Drawn head tee ports (2026-09-27)

The B1F replay exposed a second independent loss: the inline head's circle had
three real pipe endpoints, but the legacy nozzle completion attached its center
to only one. The horizontal run was preserved while the vertical return pipe
and its heads remained disconnected. Closing the 25 two-arm gaps alone could
not fix this. Head reachability and cycle counts alone were insufficient checks.

`graph/head_junctions.py` captures typed `HeadJunction` / `HeadPort` evidence
from the selected DXF material graph **before tree filtering**, as a separate
`head_junctions` cache field (version 8). It does not change tree points, edges,
head identities, fitting rules, counts or SDF conversion. Both loop and grid
consume the same source projection throughout display, flow, selection and JSON.

- Require three circumference endpoints (2 mm boundary tolerance), outward
  radial drawn arms (8 degrees), a straight opposed run and perpendicular branch.
- Require the same picked material bundle; no hard-coded layer names, nearby
  line guesses or intersection-based joins. Four ports / material conflicts are
  recorded for review, not assigned a cross fitting or an invented pair of tees.
- Reuse the established head center. Split an existing through-span at that
  center instead of duplicating it and creating a false triangle. Source XY is
  unchanged. A declared run length is distributed proportionally to the split
  spans; it is not replaced by display distance.
- Respect changed/missing arms, unknown centers and full/partial manual deletes.
  Deleting a split half must not resurrect the original unsplit bypass. Surviving
  split halves remain; undo and reopening reapply the same source evidence.
- This is physical connectivity evidence only, not a head-count load rule,
  pressure result, flow direction or tee equivalent-length determination.

On the saved B1F valve (not a fabricated diagnostic root), the replay preserved
all six top/bottom junctions of the pictured three central branches. The entire
source component changed from 270 to 285 reachable heads and cycle rank 15 to
20. Twenty head tee junctions were restored across the complete drawing, in
addition to the existing 25 straight gap closures. Pipeline coordinates/edges
and the tree flow report matched the pre-change versions. Source cache and user
edit files were hash-checked unchanged during verification. These measurements
are specific to this fixture, not a promise that arbitrary drawings are solved.

Non-head stage-4 gaps still require review: do not globally disable its no-cycle
gate (it also rejects unsupported flex-pipe shortcuts). Likewise, unsupported
four-port symbols and same-endpoint parallel pipes need explicit source evidence
and a later representation change, not silent topology repair. A complete graph
must be preserved before downstream scenario pruning; a reference tree is never
authority for deleting a physical alternate route.

Next work: evidence-backed pipe roles and connection review; cycle-safe 3D
expansion with exact edge lineage; authoritative diameters and closed nozzle
states; hydraulic mass/energy solution and scenario validation; scenario-aware
branch tee losses; validated SDF export. Do not relax the current guards merely
to make a file appear to succeed.

## Verification

Focused tests:

```powershell
python -m pytest tests/test_module_f_preserved_network.py tests/test_module_f_preserved_network_integration.py tests/test_module_f_cycle_closures.py -q
python -m pytest tests/test_module_f_head_junctions.py -q
```

Coverage: ring, three-branch grid, dormant transfer branches, disconnected visual
crossings, pendant rings, zero-length aliases, declared lengths, deterministic
IDs, exhaustive-path oracle on seeded small graphs, 2,500-node grid, HTTP flow /
selection / JSON download, save/load defaults, isolated draft publishing, tree
reversion, blocked calculation routes and inactive-slot exports, real browser UI.
Tests operate on fixtures and temporary files, not the live user's drawing.
An opt-in `MODULE_F_REAL_CYCLE_CHECK=1` test replays the confirmed local B1F DXF
in memory and writes its comparison PNG/audit JSON only to pytest's temporary
directory. It does not rewrite the live cache, saved edits or running session.
