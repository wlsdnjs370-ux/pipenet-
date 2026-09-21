"""바탕화면에 소스코드 zip 을 만든다.

git 추적 파일 + scripts/ 의 미추적 .py 를 담고, 학습용 이미지 데이터셋과
비밀정보(.env·자격증명)는 제외한다. 경로에 한글이 많아 zipfile 로 직접 쓴다.
"""
from __future__ import annotations

import os
import subprocess
import sys
import zipfile
from pathlib import Path

BASE = Path(r"C:\Users\admin\PycharmProjects\JupyterProject")
OUT = Path(os.environ["USERPROFILE"]) / "Desktop" / "fncadnet_source_20260731.zip"

EXCLUDE_PREFIX = ("data/triangle_head_dataset/",)
SECRET_NAMES = {".env", ".env.local", "credentials.json"}


def _git(*args: str) -> list[str]:
    out = subprocess.run(["git", "-C", str(BASE), *args],
                         capture_output=True, text=True, encoding="utf-8", check=True)
    return [ln for ln in out.stdout.splitlines() if ln]


def main() -> int:
    tracked = _git("ls-files")
    untracked_scripts = [p for p in _git("ls-files", "--others", "--exclude-standard", "*.py")
                         if p.startswith("scripts/")]

    rels = []
    for rel in tracked + untracked_scripts:
        if rel.startswith(EXCLUDE_PREFIX):
            continue
        if Path(rel).name in SECRET_NAMES:
            print("SKIP(secret):", rel)
            continue
        if (BASE / rel).is_file():
            rels.append(rel)

    OUT.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for rel in rels:
            z.write(BASE / rel, arcname=rel)

    print(f"파일 {len(rels)}개 → {OUT}")
    print(f"용량 {OUT.stat().st_size / 1024 / 1024:.1f} MB")
    return 0


if __name__ == "__main__":
    sys.exit(main())
