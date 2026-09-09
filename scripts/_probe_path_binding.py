# -*- coding: utf-8 -*-
"""[경로 상수 함수화] 실제 경로가 조치 전과 같은가 — 지시서 §3-3.

시험 통과만으로는 부족하다. **부팅을 실제로 시켜** 다섯 경로를 찍고, 조치 전
계산식이 내던 값과 한 줄씩 맞춰 본다.

    조치 전                                   조치 후
    handoff.OUT_DIR      = work/0단계_새찍기   handoff.pick_out_dir()
    pick.io.NEW_DIR      = 위와 같은 값        pick.io.new_dir()
    disp_cache._DISP_…   = work               disp_cache._disp_cache_dir()
    handoff.default_edits_dir() = work/DWG    (그대로)
    pick.io.STD_DIR      = docs/import/…      pick.io.std_dir()   ← 여기만 달라진다

★`STD_DIR` 만 값이 달라지는 것이 **의도**다. 종전에는 cwd 상대경로라 웹서버
  (cwd = 프로젝트 루트)에서 엉뚱한 곳을 가리켰고, 부팅 고정 목록에서도 빠져
  있었다. 데스크톱(소스) 실행에서는 두 값이 여전히 같다.

    python scripts/_probe_path_binding.py
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
sys.stdout.reconfigure(encoding="utf-8", errors="replace")


def main() -> int:
    os.environ.setdefault("LOGIN_PASSWORD", "probe")
    from routes.module_f.common import IMPORT_WORK_ROOT, _boot
    _boot()
    from services.cad_import.pick import io as pio
    from services.cad_import.pipeline import disp_cache as dc
    from services.cad_import.pipeline import handoff

    work = str(IMPORT_WORK_ROOT)
    rows = [
        ("찍은스펙 폴더", handoff.pick_out_dir(),
         os.path.join(work, "0단계_새찍기")),
        ("찍기 io 가 보는 폴더", pio.new_dir(),
         os.path.join(work, "0단계_새찍기")),
        ("표시캐시 폴더", dc._disp_cache_dir(), work),
        ("유저손질 폴더", handoff.default_edits_dir(),
         os.path.join(work, "DWG")),
    ]
    print(f"\n■ 부팅 뒤 경로  (쓰기 루트 {work})")
    bad = 0
    for name, got, want in rows:
        ok = os.path.normcase(os.path.abspath(got)) == \
            os.path.normcase(os.path.abspath(want))
        bad += 0 if ok else 1
        print(f"  [{'같음' if ok else '★다름'}] {name:18} {got}")
        if not ok:
            print(f"        조치 전 값: {want}")

    old_std = os.path.join("docs", "import", "0단계_표준샘플")
    print(f"  [의도된 변경] 표준샘플 폴더   {pio.std_dir()}")
    print(f"        조치 전 값(cwd 상대): {old_std}"
          f"  →  {os.path.abspath(old_std)}")

    # 실제 파일이 그 자리에 나는가 — 경로만 맞추고 끝내지 않는다.
    probe = os.path.join(handoff.pick_out_dir(), "_probe_path_binding.tmp")
    os.makedirs(handoff.pick_out_dir(), exist_ok=True)
    with open(probe, "w", encoding="utf-8") as fh:
        fh.write("probe")
    print(f"  [쓰기] {probe} · {os.path.getsize(probe)} bytes")
    os.remove(probe)

    # 데스크톱 G — 주입이 없을 때의 값(같은 폴더를 써야 두 실행이 이어진다).
    handoff.set_write_root(None)
    print(f"  [데스크톱 G] 쓰기 루트 {handoff.import_write_root()}"
          f" · 찍은스펙 {handoff.pick_out_dir()}"
          f" · 표준샘플 {pio.std_dir()}")
    handoff.set_write_root(work)

    print(f"\n  다른 경로 {bad}건")
    return 1 if bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
