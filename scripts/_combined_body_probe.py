# -*- coding: utf-8 -*-
"""프론트가 실제로 보내는 combined/build body 를 그대로 서버에 넣어 본다.

골든은 최소 body({plane_job_id, system_riser})만 보내 200 이 나온다. 프론트는
plane_edit·n_heads·load_mode·head_* 등을 더 얹는다 — 어느 필드가 깨뜨리는지 가른다.
"""
from __future__ import annotations

import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent
if str(BASE) not in sys.path:
    sys.path.insert(0, str(BASE))
sys.path.insert(0, str(BASE / "tests" / "characterization"))
sys.stdout.reconfigure(encoding="utf-8")

import golden_cases as gc  # noqa: E402

mod = gc._server_mod()
rp = gc._rp()
job_id = "probe_combined_plane"
gc._seed_plane_job(mod, "plane_daemyeong", job_id)
riser = rp.extract_riser_msp_28f((0.0, 0.0), (0.0, -3000.0))

c = mod.app.test_client()
with c.session_transaction() as s:
    s["authed"] = True

MINIMAL = {"plane_job_id": job_id, "system_riser": riser}
# 템플릿 7174-7200 의 _buildBody 를 그대로 옮긴 것 (기본 입력값 상태).
FULL = {
    "plane_job_id": job_id,
    "plane_edit": {"added_heads": [], "deleted_indices": [],
                   "zones": [], "branch_zones": []},
    "system_riser": riser,
    "machine_room": None,
    "pressurization": "gravity",
    "has_iso_z_scale": 1.0,
    "kfp_coord_scale": 1.0,
    "head_orientation": "pendent",
    "head_z_frac": 0.04,
    "load_mode": False,
    "n_heads": 30,
}

fails = []
for label, body in (("최소 body", MINIMAL), ("프론트 full body", FULL)):
    r = c.post("/api/remote30/combined/build", json=body)
    d = r.json if r.is_json else {}
    ok = bool(d.get("ok"))
    print(f"  {'OK ' if (r.status_code == 200 and ok) else 'FAIL'} {label}: "
          f"status={r.status_code} ok={ok} "
          f"nodes={len((d.get('geometry') or {}).get('nodes') or [])}")
    if not ok:
        print(f"       message: {str(d.get('message'))[:400]}")
        tb = str(d.get("traceback") or "")
        if tb:
            print("       traceback 끝:")
            for ln in tb.strip().splitlines()[-8:]:
                print("         " + ln)
        fails.append(label)

# full 이 깨지면 어느 필드가 범인인지 하나씩 뺀다.
if "프론트 full body" in fails:
    print("\n필드 이분 — full 에서 하나씩 제거")
    for k in [k for k in FULL if k not in ("plane_job_id", "system_riser")]:
        b = {kk: vv for kk, vv in FULL.items() if kk != k}
        r = c.post("/api/remote30/combined/build", json=b)
        d = r.json if r.is_json else {}
        print(f"  -{k:18s} status={r.status_code} ok={bool(d.get('ok'))}")

print("\n" + ("PASS — 서버 경로 정상" if not fails else "FAIL — " + ", ".join(fails)))
sys.exit(1 if fails else 0)
