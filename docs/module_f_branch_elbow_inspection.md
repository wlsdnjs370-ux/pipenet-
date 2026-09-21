# Module F: branch-tee / elbow inspection

## Policy

Module F's confirmed plan tables now include `tee` (branch), `elbow` (90°),
and `elbow-45`. Straight-through `tee-run` and legacy `cross-run` are omitted
before equivalent-length summation, counts, fitting tables, KFP and inspection.
They are not replaced with elbows or converted to branch tees.

Original graph degree and source ports remain available to fitting classification:
a pruned branch turn must still be recognized as a physical tee. Ordinary straight
pipe geometry, connectivity, physical lengths, rises, bores and nozzle data do not
change. Separate valves, alarm valves, FX, pumps and their losses are retained.
The shared G engine keeps its previous default; F explicitly supplies this policy.

## Merged inspection

The merged preview now serves inspection records from the actual combined table.
It maps original plan node identities through the merge label offset, retains
original omitted tee directions, and projects glyphs at the combined coordinates.
Source-pipe override evidence follows endpoint identities even on name collisions.
Merged edits disable old source directions at moved nodes and their neighbors.

Both viewers draw confirmed branch-tee/elbow glyphs continuously in red dashed
lines. Glyph arms/radii use network coordinates and scale with zoom. Native valve
and equipment symbols remain separate. Selecting a node or pipe opens the same
property card, with merged links staying in the merged viewer.

No tee/elbow is invented from a schematic riser line or background drawing:
an unclassified position remains unclassified. The riser builder currently does
not automatically generate elbow/tee rows; only confirmed fitting rows and
explicitly added fittings are drawn there. Missing source-port directions are
not guessed. This change affects Module F's canvas, not PipeNet's own renderer.

## Verification

- Unit/API tests: straight-tee omission, physical branch preservation, unchanged
  lengths/losses/equipment, node offset and inactive plan slots, edited geometry.
- Browser tests: always-on red dashes, current fitting orientation, zoom scaling,
  merged node/pipe cards and links, axis handles, rejection of old run-tee glyphs.
- Real Daemyeong regression: confirmed tables → SDF / KFP; only branch tees and
  elbows in F fitting rows, matching bores and equivalent lengths.

Rebuild confirmed tables and the merge after restarting the server; regenerate
exports. Existing downloaded files are not rewritten.

### Broader regression status (2026-09-20)

The full Module F + prune-fitting run passed 1,075 tests and reported seven
failures. The asset-version mismatch was corrected as part of deployment.
Six checks outside the changed fitting/inspection path remain unresolved:

- `test_module_f_adopt`: two old source-text expectations (rectangle-only region
  copy and a fixed 700-character UI explanation window).
- `test_module_f_consistency`: ID scanner does not read the separate network
  editor script or account for dynamically generated underlay IDs.
- `test_module_f_complete`: saved drawing graph/golden fingerprint mismatch.
- `test_module_f_junction_normalization`: historical B1F capture expects old
  node 48; replay now has 76 nodes rather than the captured 78.
- `test_module_f_kfp_editability`: conversion fixture omits the board required
  by the existing water-flow pipeline.

These were not hidden by updating drawing baselines or changing unrelated code.
Focused fitting, merged-preview, editor, browser and real-SDF/KFP tests are run
separately; a passing focused run does not imply the entire suite is green.
