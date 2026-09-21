# -*- coding: utf-8 -*-
"""평면도 추출 기능의 소스 일습을 바탕화면 zip 으로 묶는다.

진입점(routes/r30_prototype.py, remote30_prototype.py)에서 import 를 정적
추적해 프로젝트 내부 모듈만 모은다. .env·자격증명·데이터는 넣지 않는다.
"""
from __future__ import annotations

import ast
import sys
import zipfile
from datetime import datetime
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
sys.stdout.reconfigure(encoding="utf-8")

ENTRIES = ["routes/r30_prototype.py", "remote30_prototype.py"]
EXTRA = [
    "templates/remote30_prototype.html",
    "routes/__init__.py",
    "requirements.txt",
]
# 템플릿이 참조하는 self-host 자산 (three.js·폰트) — CDN 금지라 소스에 포함.
EXTRA_DIRS = ["static/vendor"]

# 프로젝트 안에 실재하는 모듈만 따라간다 — 서드파티는 무시.
# core/ 는 런타임에 sys.path 로 얹히므로 최상위 이름으로도 import 된다.
ROOTS = [BASE, BASE / "core"]


def resolve(name: str) -> Path | None:
    rel = name.replace(".", "/")
    for root in ROOTS:
        p = root / f"{rel}.py"
        if p.is_file():
            return p
        p = root / rel / "__init__.py"
        if p.is_file():
            return p
    return None


seen: set[Path] = set()
stack = [BASE / e for e in ENTRIES]
while stack:
    f = stack.pop()
    if f in seen or not f.is_file():
        continue
    seen.add(f)
    tree = ast.parse(f.read_text(encoding="utf-8"), filename=str(f))
    for node in ast.walk(tree):
        names: list[str] = []
        if isinstance(node, ast.Import):
            names = [a.name for a in node.names]
        elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
            names = [node.module] + [f"{node.module}.{a.name}" for a in node.names]
        for n in names:
            t = resolve(n)
            if t is not None:
                stack.append(t)

files = sorted(seen | {BASE / e for e in EXTRA if (BASE / e).is_file()})
assets = sorted(p for d in EXTRA_DIRS for p in (BASE / d).rglob("*") if p.is_file())

stamp = datetime.now().strftime("%Y%m%d_%H%M")
head = __import__("subprocess").run(
    ["git", "rev-parse", "--short", "HEAD"], cwd=BASE,
    capture_output=True, text=True).stdout.strip()
out = Path.home() / "Desktop" / f"평면도추출_소스_{stamp}_{head}.zip"

with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
    for f in files + assets:
        z.write(f, f.relative_to(BASE).as_posix())
    z.writestr("README.txt",
               f"평면도 추출 소스 일습\n"
               f"추출 시각 : {datetime.now():%Y-%m-%d %H:%M}\n"
               f"커밋      : {head} (branch main)\n"
               f"진입점    : {', '.join(ENTRIES)}\n"
               f"소스      : {len(files)}개 · 정적자산 {len(assets)}개\n\n"
               f"import 를 정적 추적해 모은 프로젝트 내부 모듈과, 템플릿이\n"
               f"참조하는 self-host 정적 자산(three.js·폰트)만 포함합니다.\n"
               f".env·자격증명·도면 데이터·서버 진입점은 포함하지 않았습니다.\n")

for f in files:
    print(f"  {f.relative_to(BASE).as_posix()}")
print(f"  static/vendor/** ({len(assets)}개)")
print(f"\n소스 {len(files)}개 + 자산 {len(assets)}개 · "
      f"{out.stat().st_size/1024/1024:.1f} MB")
print(out)
