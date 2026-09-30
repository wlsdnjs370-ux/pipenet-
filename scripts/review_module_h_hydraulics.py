"""Run an explicitly requested, non-destructive joint review on a known session."""
from __future__ import annotations

import argparse
import html
import json
from pathlib import Path
import re
import time

import requests


def main() -> None:
    """Preserve the previous report, calculate a new proposal, save its evidence."""
    parser = argparse.ArgumentParser()
    parser.add_argument("sid")
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--export-existing", action="store_true")
    args = parser.parse_args()
    base = "http://127.0.0.1:5051"
    session = requests.Session()
    login = session.get(base + "/login", timeout=15)
    hint = re.search(r'<span class="pw-value">([^<]+)<', login.text)
    if not hint:
        raise RuntimeError("Local application login hint unavailable")
    session.post(base + "/login", data={"password": html.unescape(hint.group(1)), "next": "/module-h"}, timeout=15).raise_for_status()

    def get(endpoint: str) -> dict:
        response = session.get(base + "/api/module-f/" + endpoint, params={"sid": args.sid}, timeout=30)
        response.raise_for_status()
        return response.json()

    old = get("merge/sizing/state")
    if get("job").get("state") == "running":
        raise RuntimeError("User task is running; proposal was not started")
    args.out.mkdir(parents=True, exist_ok=True)
    if args.export_existing:
        report = old['report']
        if not report['feasible'] or old['stale']:
            raise RuntimeError('Cannot export an infeasible or stale proposal')
        response = session.post(base + '/api/module-f/merge/sizing/emit', json={
            'sid': args.sid, 'proposal_id': report['proposal_id']}, timeout=30)
        response.raise_for_status()
        for _ in range(150):
            job = get('job')
            if job.get('state') == 'done':
                break
            if job.get('state') in ('error', 'cancelled'):
                raise RuntimeError(str(job.get('error') or job))
            time.sleep(2)
        else:
            raise RuntimeError('Export is still running')
        print(json.dumps({'files': get('merge/sizing/state')['files'],
                          'source_unchanged': get('merge/sizing/state')['basis']['fingerprint'] == old['basis']['fingerprint']}))
        return
    (args.out / "previous_review.json").write_text(json.dumps(old, ensure_ascii=False, indent=2), encoding="utf-8")
    options = dict(old["report"]["options"])
    options.update(mode="joint", diameter_scope="enlarge", duration_minutes=20)
    response = session.post(base + "/api/module-f/merge/sizing/run", json={
        "sid": args.sid, "fingerprint": old["basis"]["fingerprint"], "options": options}, timeout=30)
    response.raise_for_status()
    deadline = time.monotonic() + 600
    while time.monotonic() < deadline:
        job = get("job")
        if job.get("state") in ("done", "error", "cancelled"):
            if job["state"] != "done":
                raise RuntimeError(str(job.get("error") or job))
            break
        time.sleep(2)
    else:
        raise RuntimeError("Review still running; not cancelled")
    result = get("merge/sizing/state")
    if result["basis"]["fingerprint"] != old["basis"]["fingerprint"]:
        raise RuntimeError("Source changed while reviewing; do not use this result")
    (args.out / "joint_review.json").write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    report = result["report"]
    print(json.dumps({"status": report["status"], "feasible": report["feasible"], "pump": report["pump"],
        "violations": report["violations"], "minimum_head_flow": min(h["flow_lpm"] for h in report["nozzles"]),
        "minimum_head_pressure": min(h["pressure_bar"] for h in report["nozzles"]),
        "maximum_velocity": max(p["velocity_mps"] for p in report["pipes"]),
        "bottlenecks": [p for p in report["pipes"] if p["label"] in ("r17", "r19", "r20")],
        "source_unchanged": True}, ensure_ascii=False))


if __name__ == "__main__":
    main()
