"""Check the reproducible read-only PPT extraction and DXF inventory tools."""
from pathlib import Path
from zipfile import ZipFile

import ezdxf

from scripts.audit_company_head_dxf import inventory
from scripts.extract_company_head_reference import extract


def test_reference_extraction_preserves_source_images_and_native_triangle(tmp_path):
    pptx = tmp_path / "input.pptx"
    slide = '''<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <p:sp><p:nvSpPr><p:cNvPr id="31"/></p:nvSpPr><p:spPr>
      <a:prstGeom prst="triangle"/></p:spPr><a:t>측벽식</a:t></p:sp>
      <p:pic><p:blipFill><a:blip r:embed="rId1"/>
      <a:srcRect t="69717" b="19662"/></p:blipFill></p:pic></p:sld>'''
    rels = '''<Relationships><Relationship Id="rId1" Type="x/image" Target="../media/ref.png"/></Relationships>'''
    with ZipFile(pptx, "w") as archive:
        for n in range(26, 37):
            archive.writestr(f"ppt/slides/slide{n}.xml", slide)
            archive.writestr(f"ppt/slides/_rels/slide{n}.xml.rels", rels)
        archive.writestr("ppt/media/ref.png", b"original image bytes")
    before = pptx.read_bytes()
    result = extract(pptx, tmp_path / "out")
    assert [s["slide"] for s in result["slides"]] == list(range(26, 37))
    assert result["slides"][0]["native_shapes"][0]["shape"] == "triangle"
    assert result["slides"][0]["pictures"][0]["crop_percent_1000"]["t"] == "69717"
    assert (tmp_path / "out/ref.png").read_bytes() == b"original image bytes"
    assert pptx.read_bytes() == before


def test_dxf_audit_records_conflicting_temperature_not_a_head_inference(tmp_path):
    doc = ezdxf.new()
    block = doc.blocks.new("HEAD_SAMPLE")
    block.add_circle((0, 0), 10)
    doc.modelspace().add_blockref("HEAD_SAMPLE", (100, 200))
    doc.modelspace().add_text("폐쇄형68°C 상향식 헤드")
    path = tmp_path / "sample.dxf"
    doc.saveas(path)
    before = path.read_bytes()
    result = inventory(path)
    assert result["texts"][0]["text"] == "폐쇄형68°C 상향식 헤드"
    assert result["blocks"][0]["modelspace_instances"] == 1
    assert result["blocks"][0]["geometry"][0]["radius"] == 10
    assert path.read_bytes() == before
