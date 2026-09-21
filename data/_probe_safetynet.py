# -*- coding: utf-8 -*-
"""새로 clone 하면 안전망이 얼마나 남는가 — 정적으로 센다.

이 기계에는 추적 안 되는 작업 산물(도면·저장본)이 있어 시험이 «돈다».
새 clone·CI 에는 그것이 없다. 그때 몇 건이 남는지를 재는 것이 목적이다.

방법: 각 시험 파일의 skip 조건에서 «경로» 를 뽑아 git 추적 여부를 보고,
그 파일의 시험 수를 센다. 추측하지 않는다 — git 에 물어본다.
"""
from __future__ import annotations

import ast
import os
import re
import subprocess
import sys
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
ROOT = Path(__file__).resolve().parent.parent
os.chdir(ROOT)

tracked = set(subprocess.run(
    ["git", "ls-files"], capture_output=True, text=True,
    encoding="utf-8", errors="replace").stdout.splitlines())
tracked_norm = {os.path.normcase(p.replace("/", os.sep)) for p in tracked}


def is_tracked(p: str) -> bool:
    try:
        rel = os.path.relpath(os.path.abspath(p), ROOT)
    except ValueError:
        return False
    return os.path.normcase(rel) in tracked_norm


def count_tests(path: Path) -> int:
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"))
    except Exception:
        return -1
    n = 0
    for node in tree.body:
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) \
                and node.name.startswith("test"):
            n += 1
        if isinstance(node, ast.ClassDef) and node.name.startswith("Test"):
            n += sum(1 for b in node.body
                     if isinstance(b, (ast.FunctionDef, ast.AsyncFunctionDef))
                     and b.name.startswith("test"))
    return n


# ── 시험 파일마다: 시험 수 · 파일 안에 등장하는 «경로 상수»
PATH_RE = re.compile(
    r'["\']([^"\']*(?:제출용|samples[/\\]|data[/\\]|docs[/\\]|assets[/\\])'
    r'[^"\']*)["\']')

rows = []
for p in sorted(Path("tests").rglob("test_*.py")):
    n = count_tests(p)
    src = p.read_text(encoding="utf-8", errors="replace")
    # 파일 전체 skip(pytestmark) 인가
    whole = "pytestmark" in src and "skip" in src
    cands = set()
    for m in PATH_RE.finditer(src):
        s = m.group(1)
        if s.endswith((".dxf", ".xml", ".json", ".sdf", ".md", ".kfp")) \
                or "제출용" in s:
            cands.add(s)
    rows.append((p, n, whole, sorted(cands)))

print("■ 시험 수집 0건인 «시험» 파일 (pytest 가 못 보는 것)")
zero = [(p, n) for p, n, _w, _c in rows if n == 0]
for p, _n in zero:
    print(f"    {p}")
print(f"    → {len(zero)}개 파일")

print("\n■ 추적 안 되는 파일을 전제하는 시험")
missing_total = 0
for p, n, whole, cands in rows:
    bad = []
    for c in cands:
        # 상대·절대 모두 시도
        for base in (ROOT, ROOT / "tests"):
            full = (base / c) if not os.path.isabs(c) else Path(c)
            if full.exists():
                if not is_tracked(str(full)):
                    bad.append(c)
                break
        else:
            bad.append(c + "  (지금도 없음)")
    if bad:
        mark = "파일 전체" if whole else "일부"
        print(f"    {p}  — 시험 {n}건 · {mark} skip 위험")
        for b in sorted(set(bad))[:3]:
            print(f"        └ {b}")
        missing_total += n
print(f"    → 영향받는 시험 {missing_total}건")

print("\n■ 지금 이 기계에서의 전체 규모(대조용)")
total = sum(n for _p, n, _w, _c in rows if n > 0)
print(f"    tests/ 의 시험 함수 {total}건 · 파일 {len(rows)}개")
