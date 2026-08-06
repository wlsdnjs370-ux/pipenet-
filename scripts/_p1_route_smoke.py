# -*- coding: utf-8 -*-
"""③ 라우트 배선 스모크 — combined/build 의 anchored 분기를 HTTP 로 실제 통과시킨다."""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))
sys.stdout.reconfigure(encoding="utf-8")

from characterization.golden_cases import _seed_plane_job, _server_mod, _rp  # noqa: E402

WEST_UNIT_RECT = [244500.0, -243500.0, 253500.0, -221500.0]

mod = _server_mod()
rp = _rp()
job = _seed_plane_job(mod, "plane_daemyeong", "smoke_anchored")
riser = rp.extract_riser_msp_28f((0.0, 0.0), (0.0, -3000.0))
c = mod.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

for label, plane_edit in (("비-anchored", {}),
                          ("anchored", {"zones": [WEST_UNIT_RECT]})):
    r = c.post("/api/remote30/combined/build",
               json={"plane_job_id": "smoke_anchored", "system_riser": riser,
                     "plane_edit": plane_edit})
    body = r.json or {}
    geom = body.get("geometry") or {}
    audit = body.get("extraction_audit") or {}
    print(f"{label:12s} HTTP {r.status_code}  ok={body.get('ok')}  "
          f"노드 {len(geom.get('nodes') or [])}  배관 {len(geom.get('pipes') or [])}")
    if r.status_code != 200:
        print("   ", str(body.get("message"))[:300])
        continue
    if audit:
        print(f"    audit  헤드 {audit.get('heads')}  weld {len(audit.get('welds') or [])}"
              f"  bridge {len(audit.get('bridges') or [])}"
              f"  tee_split {len(audit.get('tee_splits') or [])}")
    else:
        print("    audit  (없음 — 비-anchored)")
