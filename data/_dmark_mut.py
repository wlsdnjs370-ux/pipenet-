# -*- coding: utf-8 -*-
"""기호 크기 검사가 실제로 무는지 — 렌더러를 잠깐 종이 pt 로 되돌려 본다."""
import pathlib
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "core" / "d_iso_renderer.py"
TESTS = ("tests/test_module_d/test_iso_renderer.py"
         "::test_leader_lines_are_thinner_than_pipes",
         "tests/test_module_d/test_iso_renderer.py"
         "::test_our_symbols_scale_with_the_network_not_the_page")

MUTANTS = {
    "지시선을 종이 pt 로 되돌리기": (
        "linewidths=_LEADER_WIDTH_UNITS * per_unit", "linewidths=0.25"),
    "기호 테두리를 종이 pt 로 되돌리기": (
        "    edge_pt = _MARKER_EDGE_UNITS * per_unit", "    edge_pt = 0.7"),
    "지시선이 배관보다 굵어짐": ("_LEADER_WIDTH_UNITS = 0.5", "_LEADER_WIDTH_UNITS = 1.5"),
}

original = SRC.read_text(encoding="utf-8")
try:
    for name, (good, bad) in MUTANTS.items():
        assert good in original, f"바꿀 줄을 못 찾았다: {name}"
        SRC.write_text(original.replace(good, bad), encoding="utf-8")
        r = subprocess.run([sys.executable, "-m", "pytest", *TESTS, "-q"],
                           cwd=ROOT, capture_output=True)
        print(f"{name}: {'검사가 잡았다' if r.returncode else '통과해 버렸다 — 검사가 헐겁다'}")
finally:
    SRC.write_text(original, encoding="utf-8")
print("원본 복구")
