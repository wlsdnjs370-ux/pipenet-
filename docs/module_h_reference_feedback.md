# Module H reference feedback

- Source text uses the original DXF string and anchor, but is displayed at 0° in
  both plan and isometric views. Original rotation is retained for inference and
  edit provenance; no DXF or calculation coordinates are modified.
- Text, picking and dashed rounded bounds share one screen-space layout. Bounds
  include measured multiline glyphs plus 6 px padding; connector lines terminate
  at the border. The picker caches layout until source/view/viewport changes.
- Reference changes use one revision-checked `apply` transaction with
  `response_mode=reference_patch`. Source identity, pipe library, fitting losses,
  output invalidation, persistence, merge propagation and undo history still use
  the existing canonical editor. There is no graph preview or design reload.
- The response updates the committed pipe, equipment, affected-node inspection,
  attribute record and revision locally. Camera, selection and topology persist.
  A failed save leaves the shown values unchanged. Repeated clicks are blocked
  while a save is in flight; no speculative values are shown as saved.
- Esc clears pipe selection in normal inspection; existing dialog/picker cancel
  behavior takes priority. Module F does not opt into these presentation changes.
- The upload task drawer defaults to 360 × 205 px, remains resizable, and keeps
  separate geometry from normal task/history drawers.

## Subdrawing display omission

H's system/machine-room upload opts into a separate display payload. The previous
adapter retained text for extraction but omitted it from the canvas payload, and
discarded H/S hatch/solid outlines. H now includes native TEXT/MTEXT/attribute and
dimension labels, those outlines, and full display quotas. Its layer visibility
defaults show all source bundles; manual visibility controls still apply.
The original entity list and trace-layer decisions are unchanged, and F keeps its
previous display contract. External XREF contents must still be bound into DXF;
native font substitution is used, not an exact AutoCAD typography reproduction.

## Automatic matching: discussion only, not changed here

The current matcher selects a text owner using distance and original direction.
Different adjacent sizes with unknown boundaries remain unresolved. Repeated
branch transfer requires a completely defined donor profile, matching endpoint
roles/head kinds/slot count and lengths within 15%. One unresolved donor slot can
therefore prevent an otherwise useful partial match from propagating.

Proposed next increment: report extraction/ownership/boundary/repeat rejection
counts separately; match text bounding boxes to physical head-to-head/tee/valve
segments; transfer independently validated segments between equivalent branch
patterns, retaining the exact original annotation identity. Conflicting matches
must remain visible for review instead of taking the globally nearest number.
Validate against a user-confirmed sample of this drawing before wider rollout.

## Verification (2026-09-29)

- Focused combined run: 111 passed (reference API/rollback/history, plan/iso
  browser interaction, upright labels and padded bounds, upload/Access,
  subdrawing display, and shared F network editor/subdrawing behavior).
- Local Daemyeong files render 785 system labels and 232 machine-room labels.
  Native block attributes/multiline labels and layer visibility are covered;
  hidden attributes stay hidden. Original extraction entities are unchanged.
- A wider legacy F run also found seven remaining failures outside these
  changes: six source-string assertions against earlier UI/function shapes and
  the existing `REGION` strip/fallback fixture. These were not suppressed or
  rewritten to claim a passing full suite.
- Approved 5051 restart completed; authenticated read-only browser smoke check
  passed for H and F, with no drawing writes or JavaScript errors. Existing
  uploaded subdrawing payloads need to be rebuilt by opening the DXF again.
