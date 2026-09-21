# -*- coding: utf-8 -*-
"""zip 을 빈 폴더에 풀어 평면도 추출 진입점이 실제로 import 되는지 확인."""
from __future__ import annotations

import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.stdout.reconfigure(encoding="utf-8")
zp = Path(sys.argv[1])

with tempfile.TemporaryDirectory() as td:
    with zipfile.ZipFile(zp) as z:
        z.extractall(td)
    code = (
        "import sys; sys.path.insert(0, r'%s')\n"
        "import remote30_prototype as rp\n"
        "import routes.r30_prototype as rt\n"
        "print('import OK ·', len([n for n in dir(rp) if not n.startswith('_')]), '심볼')\n"
        "print('핵심 함수 :', all(hasattr(rp, f) for f in "
        "('parse_dxf_bundle','filter_pipenet_only','detect_heads',"
        "'run_stages_0_2','run_stages_3_5','select_worst30_heads_anchored',"
        "'build_stage4_entities')))\n"
        "print('라우트 등록기:', callable(rt.register))\n" % td
    )
    import os
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONNOUSERSITE": "1"}
    r = subprocess.run([sys.executable, "-c", code], capture_output=True,
                       text=True, encoding="utf-8", errors="replace",
                       cwd=td, env=env)
    print(r.stdout.strip())
    if r.returncode:
        print(r.stderr.strip()[-1500:])
    print("\n" + ("PASS" if r.returncode == 0 else "FAIL"))
    sys.exit(r.returncode)
