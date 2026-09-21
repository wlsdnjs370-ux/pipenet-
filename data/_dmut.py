# -*- coding: utf-8 -*-
"""신규 검사가 진짜로 잡는지 본다 — 고친 자리를 되돌려 놓고 실패를 확인한다."""
import pathlib
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "core" / "d_iso_renderer.py"
TEST = "tests/test_module_d/test_iso_renderer.py::test_band_colours_are_the_ones_pipenet_uses"

GOOD = 'BAND_COLOURS = ("#ff0000", "#ffac00", "#00ff00", "#00ffff", "#0000ff", "#ff00ff")'
MUTANTS = {
    "순서를 뒤집은 경우": 'BAND_COLOURS = ("#ff00ff", "#0000ff", "#00ffff", "#00ff00", "#ffac00", "#ff0000")',
    "예전 무채색 계열": 'BAND_COLOURS = ("#1f4e9c", "#3d8bcd", "#59b4a8", "#c9a227", "#d9663d", "#a01f1f")',
}

original = SRC.read_text(encoding="utf-8")
try:
    assert GOOD in original
    for name, new in MUTANTS.items():
        SRC.write_text(original.replace(GOOD, new), encoding="utf-8")
        r = subprocess.run([sys.executable, "-m", "pytest", TEST, "-q"],
                           cwd=ROOT, capture_output=True, text=True)
        print(f"{name}: {'검사가 잡았다' if r.returncode else '★ 못 잡았다'}")
finally:
    SRC.write_text(original, encoding="utf-8")
    print("원래대로 되돌렸다")
