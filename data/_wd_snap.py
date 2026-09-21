# -*- coding: utf-8 -*-
"""작업폴더 스냅샷 — 시험이 사람의 저장본을 건드리는지 앞뒤로 대조한다."""
from __future__ import annotations

import glob
import json
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.abspath("."))
from routes.module_f.common import IMPORT_WORK_ROOT  # noqa: E402

SNAP = os.path.join("data", "_wd_snapshot.json")


def scan() -> dict:
    out = {}
    for pat in ("0단계_새찍기/*", "DWG/*", "_edit_disp_cache_*"):
        for f in glob.glob(str(IMPORT_WORK_ROOT / pat)):
            try:
                out[os.path.basename(f)] = os.path.getsize(f)
            except OSError:
                pass
    return out


if __name__ == "__main__":
    cur = scan()
    if sys.argv[1:2] == ["save"]:
        json.dump(cur, open(SNAP, "w", encoding="utf-8"))
        print(f"기록 {len(cur)}건")
    else:
        before = json.load(open(SNAP, encoding="utf-8"))
        added = sorted(set(cur) - set(before))
        gone = sorted(set(before) - set(cur))
        changed = sorted(k for k in set(before) & set(cur)
                         if before[k] != cur[k])
        print(f"이전 {len(before)} → 이후 {len(cur)}")
        print(f"  추가 {len(added)} {added[:3]}")
        print(f"  ★삭제 {len(gone)} {gone[:3]}")
        print(f"  변경 {len(changed)} {changed[:3]}")
        print("\n실폴더 불변 — 시험이 안 건드렸다"
              if not (added or gone or changed) else "\n★아직 건드린다")
