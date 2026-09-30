# Pipe-to-original-text references

The H input/property stage keeps the existing pipe context menu and adds
`참조 내경 변경`. The same menu is available in plan and isometric views.
Picking an original mapped DXF TEXT/MTEXT changes only that pipe's reference
and nominal diameter. The pipe's current material/library determines actual
inner diameter; existing loss recalculation and validation still apply.

`근거 표시` shows all extracted calculation pipes' text-reference links. When
unchecked, only the selected pipe's links remain. Each link ends at the actual
original text insertion point, transformed through the authoritative underlay
matrix for isometric display. Multiple contributing source texts may produce
multiple links. Missing, ambiguous, blocked, or purely typed values do not get
fabricated source links. The old A/B representative-to-target drawing links
are removed; the grouping table remains usable for inspection.

The pick overlay uses original text extent/rotation, supports zoom, and cancels
with Escape or right-click. The existing revision and busy guards prevent stale
writes. The server resolves the annotation ID from the current drawing and
ignores client-supplied sizes, text, coordinates, and material. A validated source
snapshot (identity, text, position, layer, entity handle) is stored in canonical
edit history, preserving the earlier automatic evidence for audit. Undo/redo
restores both value and reference, including changes between equal-size texts.
Subsequent typed property changes stop advertising the old explicit reference.
Neither the DXF nor unrelated pipes, lengths, or connectivity is changed.

Plan editing uses canonical calculation-node XY mapped back through the existing
CAD origin offset, not inverse isometric coordinates or guessed bounding boxes.
New split nodes therefore remain editable and visibly attached to the graph.
Background CAD geometry is rendered once at the original underlay opacity;
legacy yellow/blue inference strokes no longer paint over it. Extracted pipes
remain white with the existing red dashed route, selected pipes remain red,
and original diameter text remains legible.

Tests cover server-side identity resolution, native snapshots, material lookup,
atomic rejection, persistence/replay, same-size reference changes, undo/redo,
plan/iso right-click and text picking, and selection/all-reference visibility.
