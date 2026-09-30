# Inclusive pen-drawn working region

Object-definition pen crops no longer reject a crossed or retraced stroke.
The selection fills **every bounded part** of the stroke, including overlaps
and nested loops. It does not expand to a bounding box or convex hull, so
concave cut-outs and neighboring drawings remain outside. Unenclosed tails
have no selectable area. Mouse-up still closes the last point to the first.

The crop API records `fill_rule: "enclosed"` on new polygon crops. Both the
Python geometry and the browser preview split crossings and collinear overlaps
into a planar graph, walk its bounded faces, and take their union. The displayed
fill, retained entities, and persisted pipeline replay use this same rule.
Legacy region records keep their previous behavior; ordinary edit/auto head
selection regions are unchanged.

Inside the new pen region:

- Lines keep their exact interior pieces; no new pipe connections are invented.
- Circles and arcs whose centers are inside **or on the boundary** are retained
  intact, including radius and arc angles. They are not clipped into damaged
  symbols, even if part of the symbol extends beyond the pen line.
- Text retains its original contents when its insertion point is inside or on
  the boundary. Existing layer/material/head mapping still determines recognition;
  keeping an entity does not automatically make it a sprinkler pipe or head.

Original DXF files and full-world caches are not modified. Restoring the original
region is still available. Coordinate/point-count limits and a genuine nonzero
enclosed area are still required; invalid values and a straight stroke remain
clear errors. The existing post-processing guard still blocks reconstructed
pipe segments from crossing outside the retained region.

Tests cover crossed closures, bow ties, reverse/repeated outlines, overlapping
and nested loops, unclosed tails, concavity, large CAD origins, client/server
point-by-point agreement, boundary symbols, exact line clipping, unchanged DXF
bytes, browser pointer gestures, original restoration and pipeline replay.
