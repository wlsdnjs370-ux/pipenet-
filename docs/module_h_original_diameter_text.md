# Original DXF diameter labels on the drawing

H's `원본 도면` checkbox now includes the original, mapped diameter text,
in white with a dark outline, independently of the `근거 표시` links. It uses
native DXF TEXT/MTEXT, not OCR, inferred values, normalized `A` suffixes, or
calculated inner bores. Prefixes such as `SP 150` and multiline text are retained.

Labels use their original CAD insertion point and full 360-degree reading
direction. In isometric view, that point and baseline use the same authoritative
underlay transform as the original drawing. Zoom/pan changes the display only;
the CAD coordinates, pipe values, units and exported calculations are unchanged.
Font size is bounded to 11–28 screen pixels for readability. This is a legible
text overlay, not CAD font/justification/MTEXT formatting reproduction.

The existing configured/mapped diameter candidates remain authoritative. Hidden
DXF layers are excluded by the native reader. Ambiguous or missing native text
is not fabricated; this does not display arbitrary architectural text or infer
missing values. Offscreen labels are culled without shifting their CAD anchors.

Original labels are sent at the beginning of diameter tracing, before edge
matching finishes. Operation-scoped snapshots restore them after polling gaps;
they remain inspectable if a later input preview fails. A new drawing clears
the previous drawing's labels. Canvas overlays do not intercept pipe clicks.

Focused tests cover native prefix/multiline/270-degree extraction, graph-result
equivalence, progress snapshot recovery/reset, plan and iso canvas positions,
zoom/pan, original/evidence toggles, read-only behavior and failed-build display.
