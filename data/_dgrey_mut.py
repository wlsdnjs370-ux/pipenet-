# -*- coding: utf-8 -*-
"""노드 무채색 검사가 실제로 무는지 — 렌더러를 잠깐 망가뜨려 본다."""
import pathlib
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "core" / "d_iso_renderer.py"
TEST = "tests/test_module_d/test_iso_renderer.py::test_node_bands_are_grey_not_the_link_colours"

GOOD = 'NODE_BAND_COLOURS = ("#000000", "#2a2a2a", "#545454", "#7e7e7e", "#a9a9a9", "#d4d4d4")'
MUTANTS = {
    "관로 색 재사용": "NODE_BAND_COLOURS = BAND_COLOURS",
    "계단 뒤집기": 'NODE_BAND_COLOURS = ("#d4d4d4", "#a9a9a9", "#7e7e7e", "#545454", "#2a2a2a", "#000000")',
    "간격 어긋남": 'NODE_BAND_COLOURS = ("#000000", "#333333", "#555555", "#888888", "#aaaaaa", "#dddddd")',
}

original = SRC.read_text(encoding="utf-8")
assert GOOD in original, "바꿀 줄을 못 찾았다"
try:
    for name, mutant in MUTANTS.items():
        SRC.write_text(original.replace(GOOD, mutant), encoding="utf-8")
        r = subprocess.run([sys.executable, "-m", "pytest", TEST, "-q"],
                           cwd=ROOT, capture_output=True, text=True)
        print(f"{name}: {'검사가 잡았다' if r.returncode else '통과해 버렸다 — 검사가 헐겁다'}")
finally:
    SRC.write_text(original, encoding="utf-8")
print("원본 복구")
