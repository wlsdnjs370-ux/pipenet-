# Module H representative overlay focus

The former overlay connected every target record to the first point of the
first representative record. This produced a dense all-to-one fan, duplicate
representative labels, and incorrect reference endpoints for compacted pipes.

The correction is presentation-only:

- Overview shows representative geometry and at most one representative label.
- A representative-table choice inspects one target, not an implicit bulk-edit
  selection. All targets remain available individually in the existing list.
- A focused target links only to its recorded source edge. Within a compacted
  record the matching edge index selects the exact source segment.
- Missing donor geometry never falls back to an unrelated first origin.
- Labels avoid each other and viewport edges; invisible labels are omitted.
- The evidence overlay alone draws representative connectors. The attributes
  overlay continues selection highlighting, without drawing duplicate links.
- Multiple selection does not produce a fan; unresolved selection clears stale
  reference highlights. Computed sizes, source evidence and the graph are unchanged.

Browser tests cover a 40-target group, exact source-segment endpoints, missing
donors, overlapping labels, offscreen labels, live batches and preserved data.
The requested bottom-docked editable reference table is a separate UI proposal,
not part of this correction and not enabled by this change.
