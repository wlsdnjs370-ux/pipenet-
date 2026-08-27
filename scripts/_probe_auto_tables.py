# -*- coding: utf-8 -*-
"""자동(A) 경로의 표가 수동(G) 경로의 표와 어디까지 같은가.

두 길의 산출이 하류에서 «같은 자리» 에 들어가므로, 칸이 다르면 화면이 조용히
빈다. 무엇이 없는지 여기서 먼저 안다.

    python scripts/_probe_auto_tables.py
"""
from __future__ import annotations

import inspect
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main() -> int:
    sys.stdout.reconfigure(errors="replace")
    for p in (str(ROOT), str(ROOT / "core")):
        if p not in sys.path:
            sys.path.insert(0, p)

    import remote30_prototype as A
    src = inspect.getsource(A.build_input_tables)

    print("A · build_input_tables")
    print(f"  dia_src 칸  : {'있음' if 'dia_src' in src else '없음'}")
    print(f"  dia_source  : {'있음' if 'dia_source' in src else '없음'}")

    keys = re.findall(r'\(\s*"([^"]+)"\s*,', src)
    meta_keys = [k for k in keys if any(w in k for w in
                                        ("관경", "앵커", "기준", "헤드", "제목"))]
    print("  meta 후보   : " + (" · ".join(dict.fromkeys(meta_keys))
                                or "(없음)"))

    # G 쪽 배관 행의 칸
    sys.path.insert(0, str(ROOT / "cad_project_editor_g"))
    from services.cad_import.design import tables as gt
    gsrc = inspect.getsource(gt.build_design_tables)
    print("\nG · build_design_tables")
    print(f"  dia_src 칸  : {'있음' if 'dia_src' in gsrc else '없음'}")
    gkeys = re.findall(r'\(\s*"([^"]+)"\s*,', gsrc)
    print("  meta 후보   : "
          + (" · ".join(dict.fromkeys(k for k in gkeys
                                      if any(w in k for w in
                                             ("관경", "앵커", "기준", "헤드"))))
             or "(없음)"))
    return 0


if __name__ == "__main__":
    sys.exit(main())
