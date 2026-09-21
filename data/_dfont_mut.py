# -*- coding: utf-8 -*-
"""글자 크기 검사가 실제로 무는지 — 렌더러를 잠깐 망가뜨려 본다."""
import pathlib
import subprocess
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
SRC = ROOT / "core" / "d_iso_renderer.py"
TEST = "tests/test_module_d/test_iso_renderer.py::test_label_height_is_fixed_in_model_units"

MUTANTS = {
    "종이 pt 로 되돌리기": ("    label_pt = _LABEL_UNITS * per_unit",
                     "    label_pt = 4.6"),
    "값 글씨 상수 어긋남": ("_LABEL_UNITS = 24.17", "_LABEL_UNITS = 26.0"),
    "주기가 typesize 를 무시": (
        "    units = _NOTE_UNITS * (item.typesize or _NOTE_TYPESIZE) / _NOTE_TYPESIZE",
        "    units = _NOTE_UNITS"),
}

original = SRC.read_text(encoding="utf-8")
try:
    for name, (good, bad) in MUTANTS.items():
        assert good in original, f"바꿀 줄을 못 찾았다: {name}"
        SRC.write_text(original.replace(good, bad), encoding="utf-8")
        r = subprocess.run([sys.executable, "-m", "pytest", TEST, "-q"],
                           cwd=ROOT, capture_output=True)
        print(f"{name}: {'검사가 잡았다' if r.returncode else '통과해 버렸다 — 검사가 헐겁다'}")
finally:
    SRC.write_text(original, encoding="utf-8")
print("원본 복구")
