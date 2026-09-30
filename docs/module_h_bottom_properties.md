# Module H bottom property dock

The plan attribute stage and integrated view use one canonical network-editor
transaction path. The plan dock opens from `전체 테이블` or More → input table,
not by selecting a pipe/head. The integrated dock still appears on entering integration;
once closed, selections cannot reopen it. Row selection and drawing selection
are linked while open, and drawing highlights still update while closed. An
explicit reopen selects the latest object. The dock is temporarily hidden
while another task drawer or bulk-selection tool owns the workspace. Escape
closes it; the header collapses, drags, or resets on double-click. Every border
and corner resizes the H windows within the viewport. No camera or window
gesture changes network geometry.

In the H plan overlay, active pipes use a solid red 7 px stroke painted last.
Unselected network/evidence colours are unchanged. Unknown diameters also use
a solid active stroke (not their ordinary dashed evidence style); this denotes
selection only, not a confirmed diameter. Turning evidence display off does not
hide an active selection. Heads use a red outline when selected.

Diameter edits, length/C edits, and head properties use the existing versioned
preview/apply operations and shared undo/redo history. They do not directly
modify display-only rows. A change to material resets the available nominal
sizes; actual inner diameter remains library-derived, in mm. Length is in m.
Input validation and missing library records remain blocking errors, not guesses.

The node/head table omits X and Y columns in both plan and integrated views.
Only Z (m) is exposed for position editing. Applying height passes the unchanged,
full-precision canonical X/Y values to the existing validated move transaction;
removing the columns does not remove or round stored coordinates. Head fields,
selection linkage and undo/redo remain available.

Native DXF metadata now retains decoded source text, layer and an entity handle
when available. Handles on virtual block entities can be unavailable; coincident
ambiguous text metadata is not invented. Continuous/compacted pipes retain all
contributing evidence separately from manual overrides. The redundant bottom
source strip (mini canvas, source text/coordinates and original-position buttons)
has been removed at user request. Its space is now used by table rows; drawing
and row selection remain bidirectional. This removes presentation only, not the
stored DXF provenance or the separate representative-evidence overlay.

H `drawing_first_v1` integration skips legacy upstream-diameter normalization.
It therefore preserves explicit drawing diameters and unknown `0` values. This
does not certify hydraulic validity; inverse sizing remains a separate workflow.
F retains the default normalizer. Stitch order records a combined-label →
source-plan-label map so colliding labels cannot restore a plan value onto a
system pipe or send an edit to the wrong original row.

Previously calculated/merged session results are not silently repaired or
reinterpreted. Reopen saved drawings after restart and rebuild attributes and
integration to obtain fresh source evidence and the preserved diameter values.
Existing edit histories remain authoritative and conflicts require review.

Focused checks: `test_module_h_bottom_properties_browser.py`,
`test_module_h_windows_browser.py`, `test_module_h_merge_bores.py` plus the H
browser/provenance and F merge/editor regression suites. Test graphs and saved
history writes are isolated in temporary directories.

H view controls sit directly below diameter definition: `등각투상` switches the
existing plan/design view and `원본 도면` shows/hides the original DXF bundles,
with a white dashed drawing-bounds frame. This is not a network-dimming control.
Plan uses original coordinates; isometric uses the server's `view.underlay`
transform at the connection elevation. Missing transforms are not approximated.
Layer visibility is respected and cached paths avoid reparsing CAD per frame.
The third checkbox, `전체 테이블`, opens the bottom property dock with its search
cleared and any folded window expanded. Unchecking closes it; close/Escape and
More → input table synchronize its check state. Folding retains the checked
state because the title bar is still visible. Table values still use canonical
editor transactions. Merely changing view controls does not edit calculations.

H's plan-design isometric view uses the same white, load-weighted corridor and
red dashed worst-path style as the plan. The legacy pink provenance glow,
exclamation markers and duplicate legend are suppressed only in this H context;
F and integration preferences are unchanged. `근거 표시` separately controls H's
source evidence overlay in both views. Actual pipes and heads follow the exact
preview nodes (including generated vertical links); original CAD source segments
and text use the server underlay transform. A chosen target displays only its
exact source-to-target link, and selection does not switch back to plan. Camera
changes and replacement previews invalidate display caches. Missing source
transforms produce no guessed alignment. Selected pipes remain solid red even
when evidence is hidden.

The old full-width bottom action strip is replaced by a compact, content-width
button cluster at bottom-left: fit, back, history and the current next action.
The canvas now fills the height below the header. Hidden status anchors remain
for existing controllers; no blank strip intercepts drawing interactions at the
bottom-right. Buttons wrap on narrow screens, and default table/drawer placement
clears the actual cluster height. All native actions and busy guards are retained.

The representative-evidence window folds to its title bar, hiding status and
contents even after border resizing. Expanding restores its previous height;
dragging the collapsed bar does not discard the expanded size. Focused tests:
`test_module_h_view_browser.py`, `test_module_h_evidence_fold_browser.py`,
`test_module_h_view_evidence_browser.py`, `test_module_h_footer_browser.py`.
