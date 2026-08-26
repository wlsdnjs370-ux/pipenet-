# -*- coding: utf-8 -*-
"""handoff 캐시에 치수 텍스트가 실제로 들어 있는지 직접 본다.

`world.texts == 0` 이면 관경의 도면 텍스트 경로가 통째로 죽는다(별표1 100%).
그 0 이 «도면에 글자가 없어서» 인지 «캐시가 안 담아서» 인지는 sqlite 를 열어야
갈린다 — 파이프라인을 다시 돌리면 또 수십 분이다.

    python scripts/_probe_handoff_texts.py
"""
from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
KEY = "B1F 현장조사 소화설비 평면도"


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    if str(ROOT) not in sys.path:
        sys.path.insert(0, str(ROOT))
    from routes.module_f.common import _boot
    _boot()
    from services.cad_import.pipeline import handoff

    path = handoff.handoff_path(KEY)
    print(f"handoff: {path}")
    print(f"  존재: {os.path.exists(path)}")
    if not os.path.exists(path):
        return 1
    print(f"  크기: {os.path.getsize(path):,} bytes")

    con = sqlite3.connect(path)
    try:
        tables = [r[0] for r in con.execute(
            "SELECT name FROM sqlite_master WHERE type='table'").fetchall()]
        print(f"\n  표: {', '.join(tables)}")
        cols = {t: [c[1] for c in con.execute(f"PRAGMA table_info({t})")]
                for t in tables}
        for t in tables:
            print(f"    {t}: {', '.join(cols[t])}")

        if "meta" in tables:
            mc = cols["meta"]
            meta = dict(con.execute(
                f"SELECT {mc[0]},{mc[1]} FROM meta").fetchall())
            print("\n  meta")
            for k in sorted(meta):
                print(f"    {str(k):16s} {str(meta[k])[:70]}")

        print("\n  표별 행 수")
        for t in tables:
            try:
                n = con.execute(f"SELECT COUNT(*) FROM {t}").fetchone()[0]
                print(f"    {t:6s} {n:,}")
            except sqlite3.Error as exc:
                print(f"    {t:6s} (없음: {exc})")

        rows = con.execute(
            "SELECT layer,color,x,y,h,text FROM txt LIMIT 400").fetchall()
        print(f"\n  txt 표본 {len(rows)}행")
        if rows:
            from collections import Counter
            lay = Counter(r[0] for r in rows)
            print("    레이어:", " · ".join(f"{k}({v})" for k, v in lay.most_common(8)))
            for r in rows[:15]:
                print(f"      [{r[0]}] h={r[4]:.0f} ({r[2]:.0f},{r[3]:.0f}) {r[5]!r}")

            # 치수로 읽히는 것이 있나
            from services.cad_import.design.bore import extract_dia_text_points
            allrows = con.execute(
                "SELECT layer,color,x,y,h,text FROM txt").fetchall()
            dia = extract_dia_text_points(allrows)
            print(f"\n    치수로 읽힌 것: {len(dia)} / {len(allrows)}")
            if dia:
                from collections import Counter as C2
                c = C2(d for _x, _y, d in dia)
                print("    값 분포:", " · ".join(f"{k}A×{v}" for k, v in sorted(c.items())))
    finally:
        con.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
