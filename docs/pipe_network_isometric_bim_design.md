# Pipe Network Isometric BIM Module Design

## Purpose

The former plan BIM 3D preview tried to extrude generic floor-plan geometry into walls and columns. That behavior is not suitable for sprinkler calculation review. The updated module is a pipe-network-only isometric preview:

1. Use the current DXF visible layers and optional selected area as the work scope.
2. Keep only sprinkler pipe centerlines, heads, valves, fittings, and diameter texts.
3. Build a snapped 2D network from pipe-like DXF geometry.
4. Render the result as a 3D isometric centerline graph.

This is a review and extraction aid, not an architectural BIM authoring tool.

## Input Contract

- Source: parsed DXF payload already loaded in `cad_compare_module_7.html`.
- Scope: current visible DXF layers plus locked/active selection box when present.
- Included entity classes:
  - pipe centerline: `LINE`, `LWPOLYLINE`, `ARC` on fire protection or pipe-like layers
  - sprinkler head candidates
  - valve and alarm valve candidates
  - supported diameter text near pipes
- Excluded entity classes:
  - architectural walls, doors, windows, columns
  - dimensions, hatches, grids, centerlines, annotations
  - arbitrary non-pipe linework unless the user explicitly selects a connected pipe network component

## Network Construction

- Pipe segments are clipped to the current work box.
- Segment endpoints are snapped by the UI-configured tolerance.
- Snapped endpoints become network nodes.
- Segments become pipe edges.
- Nearby diameter text is attached to edges when available.
- Heads and valves are placed as markers and attached visually to the nearest network node.
- Degree `>= 3` nodes are shown as branch/fitting candidates.

## Isometric Rendering

- The preview uses Three.js with an orthographic camera initialized to an isometric angle.
- X/Y comes from the DXF plan.
- Z is display-only:
  - main/riser pipe: configured base pipe height
  - branch pipe: base pipe height plus branch offset
  - heads: dropped from the nearest pipe node
- Pipe radius is visualized from detected or default bore.
- Color semantics:
  - teal: main or large-bore pipe
  - blue: branch pipe
  - violet: riser-like pipe
  - yellow: head
  - red: valve/alarm valve
  - orange: branch/fitting candidate

## Current Limitations

- Display Z is not calculation elevation. Calculation Z must still come from section/system rules.
- Raw coordinate matching against system diagrams is intentionally not used.
- Layer naming must be mapped or filtered by the user when a drawing uses nonstandard pipe layer names.
- Intersections without split endpoints still depend on the existing DXF graph extraction quality.

## Next Integration Step

Feed this module's pipe-only graph into the structured `pipenet_converter` models:

- `Node`: snapped graph nodes and head output nodes
- `Pipe`: centerline edges with length, bore, rise, and C-factor
- `Nozzle`: head markers attached to nearest nodes
- `Valve`/`Equipment`: valve and alarm valve candidates

The server-side converter should remain the authority for SDF writing and validation.
