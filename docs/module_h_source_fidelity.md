# Module H: native subdrawing display and shared reading

## Scope and safety boundary

H's system and machine-room upload now builds a native display snapshot and the
existing calculation entities from one loaded DXF document. F's default reader,
pipe-layer classification, path tracing, fitting rules and hydraulic algorithms
are unchanged. `SourceDisplay` is display-only: never feed its hatch outlines,
symbol strokes or tessellated curves to the pipe graph.

The previous H adapter took geometry from a reduced calculation representation,
then reopened the original DXF solely to obtain labels. It lost polyline closing
edges, bulges and some native arc transforms. It also enabled all hidden layers
and enlarged every tiny label to at least three pixels.

## Preservation

- Native virtual entities handle nested block bases, scale, rotation and OCS.
- Closed polylines include their closing segment; circular bulges remain arcs.
- Noncircular/tilted curves use display tessellation, tolerance 0.25 CAD units
  (millimetres for this project). Calculation lengths are not measured from it.
- Layer OFF/FROZEN and hidden parent blocks are retained in separate bundles,
  disabled initially. Re-enabling is a display action, not a graph edit.
- Labels retain position, rotation, alignment, width factor and native height.
  Labels smaller than 0.75 screen pixels are culled only while zoomed out; their
  data remain intact. No OCR, number synthesis or coordinate relocation occurs.
- An explicit untwisted modelspace top-view camera is restored from VPORT's
  target + display-coordinate center. Perspective/tiled/twisted views fall back
  to visible object bounds. `CAD 시점` restores the authored camera;
  `화면 맞춤` fits currently enabled bundles. Neither operation crops data.

## Caching and first-read work

`read_view(source_display=True)` has a separate fingerprint/cache namespace from
F's legacy view. The existing conservative unused-block stripping is retained,
with source-display equality tested against the original document. A single
ezdxf read supplies both the unchanged calculation traversal and the native
display traversal, including labels. Geometry and label data are cached together.
The snapshot cache includes display reader sources and ezdxf/Python versions.

## Current limitations

This is a modelspace 2D display, not a full CAD engine. Referenced raster images
and WIPEOUT masks are reported as unsupported, not silently treated as pipework.
Hatch boundaries and SOLID outlines are shown; full hatch pattern/filled painter
order, exact SHX font metrics and rich MTEXT formatting are not reproduced.
XREF content still requires a bound/exported DXF. Unsupported counts are exposed
in processing history and the H status. Restoring display geometry does not claim
that every legacy calculation geometry interpretation has been corrected.

## Verification

Focused tests cover rotations (including nonorthogonal), nested block bases,
negative OCS, signed arc sweep, closed/bulged polylines, nonuniform scale,
native label metadata, hidden-parent inheritance, camera restoration, unchanged
calculation entities, one-read behavior, warm cache, renderer invalidation and
stripped/original source equality. Browser tests check native label scale,
zoom-dependent visibility, hidden-layer fit and signed arc painting.

Run the offline real-drawing audit without opening/changing a user session:

```powershell
python -m scripts.audit_h_source_display "data/uploads/<drawing>.dxf"
```

Its report verifies source SHA-256 and exact equality of all legacy parsed fields,
and records actual timings. Generated JSON/PNGs live in
`outputs/h_source_display_review/`; these are review artifacts, not design inputs.
