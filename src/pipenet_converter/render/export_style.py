"""Drawing-only scale for readable PIPENET exports, not hydraulic dimensions."""

# Verified SDF templates do not expose a node/nozzle glyph-size attribute.
# Use the existing uniform drawing-space technique for all Module F exporters.
# Keep actual lengths, elevations, bore and pressure separate and unchanged.
# Four rather than the previous merged-only factor of three also reduces
# apparent symbols in the normal merged export by another quarter.
EXPORT_SYMBOL_SPREAD = 4.0
