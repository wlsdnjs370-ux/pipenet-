"""Drive the REAL emit path (build_input_tables + emit_sdf) with a CPVC zone
and confirm the zone pipes are written into the CPVC2 Pipe-set (schedule=CPVC2,
c-factor 150) rather than the KSD 3507 Pipe-set. Uses a zone that covers the
whole pipe bbox so ALL pipes must relocate — proves the relocation mechanics
end-to-end. Prints per-Pipe-set pipe counts.
"""
import sys
import tempfile
import xml.etree.ElementTree as ET
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import remote30_prototype as R  # noqa: E402

DXF = Path(sys.argv[1]) if len(sys.argv) > 1 else \
    Path(__file__).resolve().parent.parent / "samples" / "dxf" / \
    "대명동201동 단위세대_layer정리.dxf"


def main():
    bundle = R.parse_dxf_bundle(DXF)
    cat = {ly["name"]: ly["auto_category"] for ly in bundle.layers}
    pipe_ents = R.filter_pipenet_only(bundle)
    sel = R.select_worst30_heads(pipe_entities=pipe_ents, layer_categories=cat)
    print("selection heads=%d edges=%d source=%s"
          % (len(sel.heads), len(sel.edges), sel.source_pos is not None))

    # whole-bbox zone → every pipe midpoint inside → all CPVC
    xs, ys = [], []
    for (a, b, _) in sel.edges:
        xs += [a[0], b[0]]; ys += [a[1], b[1]]
    zone = (min(xs) - 1, min(ys) - 1, max(xs) + 1, max(ys) + 1)
    print("full-bbox zone:", tuple(round(z, 1) for z in zone))

    tables = R.build_input_tables(sel, pipe_entities=pipe_ents,
                                  project_title="cpvc-verify",
                                  cpvc_zones=[zone])
    n_cpvc = sum(1 for p in tables.pipes if str(p.get("type")) == R.CPVC_PIPE_TYPE)
    print("tables pipes=%d  type==CPVC2=%d" % (len(tables.pipes), n_cpvc))

    out = Path(tempfile.gettempdir()) / "_cpvc_verify.sdf"
    R.emit_sdf(tables, out, project_title="cpvc-verify")
    root = ET.parse(out).getroot()
    print("--- Pipe-set breakdown ---")
    for ps in root.iter("Pipe-set"):
        nm_el = ps.find("Pipe-type/Name")
        nm = nm_el.text if nm_el is not None else "(empty placeholder)"
        pipes = ps.findall("Pipe")
        cs = sorted({p.get("roughness-or-c") for p in pipes})
        print("  schedule=%-18s pipes=%3d  roughness-or-c=%s"
              % (nm, len(pipes), cs))
    # partial-zone sanity: half the bbox
    midx = (min(xs) + max(xs)) / 2.0
    half = (min(xs) - 1, min(ys) - 1, midx, max(ys) + 1)
    tables2 = R.build_input_tables(sel, pipe_entities=pipe_ents,
                                   project_title="cpvc-half",
                                   cpvc_zones=[half])
    n2 = sum(1 for p in tables2.pipes if str(p.get("type")) == R.CPVC_PIPE_TYPE)
    print("half-bbox zone → CPVC pipes=%d / %d" % (n2, len(tables2.pipes)))


if __name__ == "__main__":
    main()
