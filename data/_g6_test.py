# -*- coding: utf-8 -*-
"""[G6] 수용 기준 검사 추가."""
import io

p = "cad_project_editor_g/tests/test_design_input.py"
s = io.open(p, encoding="utf-8").read()

block = '''

# ─────────────────────────────────────────────────────────── G6
def g6():
    print("\\n[G6] SDF 방출")
    import os
    import xml.etree.ElementTree as ET
    from services.cad_import.design.emit import (
        AssetMissing, emit_design_sdf, resolve_standard_slf,
        resolve_template_sdf)

    tbl = g5()
    if tbl is None:
        return None

    check("템플릿 SDF 해석", resolve_template_sdf().is_file(),
          resolve_template_sdf().name[:40])
    check("표준 SLF 해석", resolve_standard_slf().is_file(),
          resolve_standard_slf().name)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out = OUT_DIR / "module_g_design.sdf"
    got = emit_design_sdf(tbl, out, project_title="Module G 검증")
    check("SDF 생성", got.is_file() and got.stat().st_size > 1000,
          f"{got.stat().st_size:,} bytes")
    slf = got.with_suffix(".slf")
    check("SLF 를 한 쌍으로 저장", slf.is_file(), f"{slf.name}")

    root = ET.parse(got).getroot()
    nodes = root.findall(".//Nodes/Node")
    pipes = root.findall(".//Links//Pipe")
    nozzles = root.findall(".//Links//Nozzle")
    check("노드 수가 테이블과 일치", len(nodes) == len(tbl.nodes),
          f"SDF {len(nodes)} / 표 {len(tbl.nodes)}")
    check("배관 수가 테이블과 일치", len(pipes) == len(tbl.pipes),
          f"SDF {len(pipes)} / 표 {len(tbl.pipes)}")
    check("노즐 수가 테이블과 일치", len(nozzles) == len(tbl.nozzles),
          f"SDF {len(nozzles)} / 표 {len(tbl.nozzles)}")

    # ★관경이 "Unset" 으로 뜨면 안 된다 — bore 속성이 실수로 들어가야 한다.
    bores = [p_.get("bore") for p_ in pipes]
    bad = [b for b in bores if b in (None, "", "Unset")]
    check("관경이 Unset 이 아니다", not bad,
          f"Unset {len(bad)} / 배관 {len(bores)} · 예 {bores[:4]}")

    # 템플릿을 썼으므로 Graphics 블록(표시 메타)이 살아 있어야 한다.
    check("템플릿의 Graphics 블록 보존",
          root.find(".//Graphics") is not None, "Graphics 있음")

    # ★자산이 없으면 «파일을 만들지 않고» 실패해야 한다.
    out2 = OUT_DIR / "should_not_exist.sdf"
    if out2.exists():
        out2.unlink()
    os.environ["REMOTE30_TEMPLATE_SDF"] = str(OUT_DIR / "no_such_template.sdf")
    try:
        emit_design_sdf(tbl, out2)
        check("자산 없으면 실패", False, "예외 없이 진행했다")
    except AssetMissing as exc:
        check("자산 없으면 실패", True, str(exc).splitlines()[0][:52])
    finally:
        os.environ.pop("REMOTE30_TEMPLATE_SDF", None)
    check("실패 시 파일을 남기지 않는다", not out2.exists(), str(out2.name))
    return got


ITEMS = {"G1": g1, "G2": g2, "G3": g3, "G4": g4, "G5": g5, "G6": g6}'''

s = s.replace('\n\nITEMS = {"G1": g1, "G2": g2, "G3": g3, "G4": g4, "G5": g5}',
              block, 1)
assert '"G6": g6' in s
io.open(p, "w", encoding="utf-8", newline="\n").write(s)
print("G6 검사 추가")
