# -*- coding: utf-8 -*-
"""head drop 상한 + X 교차 미절단 조치 후 대명동 서측 실측."""
from __future__ import annotations

import sys
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO))
sys.path.insert(0, str(REPO / "tests"))
sys.stdout.reconfigure(encoding="utf-8")

from characterization.golden_cases import _seed_plane_job, _server_mod, _rp  # noqa: E402

WEST = [244500.0, -243500.0, 253500.0, -221500.0]

mod = _server_mod()
rp = _rp()
job = _seed_plane_job(mod, "plane_daemyeong", "tighten_probe")
riser = rp.extract_riser_msp_28f((0.0, 0.0), (0.0, -3000.0))
c = mod.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

r = c.post("/api/remote30/combined/build",
           json={"plane_job_id": "tighten_probe", "system_riser": riser,
                 "plane_edit": {"zones": [WEST]}})
body = r.json or {}
audit = body.get("extraction_audit") or {}
geom = body.get("geometry") or {}
pl = (audit.get("ortho") or {}).get("planar") or {}

print("헤드      ", audit.get("heads"))
drops = audit.get("head_drops") or []
print("head_drops", len(drops),
      "최대", round(max((d["len_mm"] for d in drops), default=0.0), 1))
print("조각      ", audit.get("fragments"))
print("planar    ", {k: v for k, v in pl.items() if k != "dropped_nodes"})
print("최종망    ", f"노드 {len(geom.get('nodes') or [])}  배관 {len(geom.get('pipes') or [])}")
