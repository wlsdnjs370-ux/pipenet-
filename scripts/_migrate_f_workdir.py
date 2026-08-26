# -*- coding: utf-8 -*-
"""[F-0] E 작업폴더 → G 작업폴더 1회 복사.

모듈 F 의 엔진이 E 에서 G 로 재지정되면서(지시서 D1) 작업폴더도
`cad_project_editor/docs/import` → `cad_project_editor_g/docs/import` 로 옮겨간다.
기존 F 사용자가 E 폴더에 쌓아 둔 찍은스펙·유저손질·표시캐시·설명그림은 재지정
후 보이지 않으므로 여기서 복사한다.

원칙 세 가지:
  · **원본 보존** — E 는 동결이지 삭제가 아니다. E 쪽 파일은 건드리지 않는다.
  · **덮어쓰지 않는다** — 양쪽에 다 있는 파일은 그대로 두고 «충돌» 로 보고한다.
    B1F 실측: 두 트리의 유저손질이 서로 다른 사람 편집을 담고 있다(E 923KB ·
    G 909KB). 어느 쪽이 옳은지는 기계가 정할 일이 아니다.
  · **조용히 넘어가지 않는다** — 몇 개를 복사했고 몇 개가 충돌인지 전부 찍는다.

    python scripts/_migrate_f_workdir.py            # 복사 실행
    python scripts/_migrate_f_workdir.py --dry-run  # 무엇이 복사될지 보기만
"""
from __future__ import annotations

import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "cad_project_editor" / "docs" / "import"
DST = ROOT / "cad_project_editor_g" / "docs" / "import"


def main() -> int:
    dry = "--dry-run" in sys.argv
    if not SRC.is_dir():
        print(f"원본 작업폴더가 없습니다: {SRC}")
        return 1
    DST.mkdir(parents=True, exist_ok=True)

    copied, conflicts, same = [], [], 0
    for src in sorted(SRC.rglob("*")):
        if not src.is_file():
            continue
        rel = src.relative_to(SRC)
        dst = DST / rel
        if dst.exists():
            if (dst.stat().st_size == src.stat().st_size
                    and dst.read_bytes() == src.read_bytes()):
                same += 1
            else:
                conflicts.append(rel)
            continue
        copied.append(rel)
        if not dry:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)

    tag = "[dry-run] " if dry else ""
    print(f"{tag}복사 {len(copied)}개 · 동일(건너뜀) {same}개 · "
          f"충돌(양쪽 상이 — 안 덮음) {len(conflicts)}개")
    for rel in copied[:40]:
        print(f"  + {rel}")
    if len(copied) > 40:
        print(f"  … 외 {len(copied) - 40}개")
    if conflicts:
        print("\n충돌 — 양쪽에 서로 다른 내용이 있어 G 쪽을 그대로 두었습니다.")
        print("E 쪽을 쓰려면 그 파일만 직접 복사하세요:")
        for rel in conflicts:
            s, d = SRC / rel, DST / rel
            print(f"  ! {rel}")
            print(f"      E {s.stat().st_size:>10,}B "
                  f"{__import__('datetime').datetime.fromtimestamp(s.stat().st_mtime):%m-%d %H:%M}"
                  f"  vs  G {d.stat().st_size:>10,}B "
                  f"{__import__('datetime').datetime.fromtimestamp(d.stat().st_mtime):%m-%d %H:%M}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
