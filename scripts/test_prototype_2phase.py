"""2-phase 파이프라인 검증 — Stage 0~2 stream → finalize POST → Stage 3~5 stream."""

from __future__ import annotations

import io
import json
import sys
import time
from pathlib import Path

import requests

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

BASE = "http://127.0.0.1:5050"
DXF = Path(r"C:\Users\admin\PycharmProjects\JupyterProject\대명동201동 단위세대_layer정리.dxf")


print(f"[1] POST run")
with DXF.open("rb") as f:
    r = requests.post(f"{BASE}/api/remote30/prototype/run",
                      files={"dxf_file": (DXF.name, f, "application/octet-stream")},
                      timeout=60)
j = r.json()
print(f"  → {j}")
job_id = j["job_id"]

print(f"\n[2] Stage 0~2 stream")
t0 = time.time()
last_event = None
with requests.get(f"{BASE}/api/remote30/prototype/stream/{job_id}", stream=True, timeout=180) as resp:
    for raw in resp.iter_lines():
        if not raw or not raw.startswith(b"data: "):
            continue
        evt = json.loads(raw[6:])
        t = evt.get("type")
        if t == "stage":
            print(f"  [{(time.time()-t0):5.1f}s] stage {evt['stage']} {evt['status']:8s}: {evt['label']}")
        elif t == "entities":
            print(f"           ↳ entities stage={evt['stage']} count={len(evt['entities'])} summary={evt.get('summary')}")
        elif t == "awaiting_finalize":
            print(f"           ↳ AWAITING_FINALIZE: head_count={evt.get('head_count')}")
            last_event = evt
            break
        elif t == "error":
            print(f"           ↳ ERROR {evt['message']}")
            break

print(f"\n[3] POST finalize — added 2 heads + 1 zone, no edits otherwise")
zone_x1, zone_y1 = 240000, -245000   # 대명동 도면 일부 영역
zone_x2, zone_y2 = 290000, -220000
payload = {
    "added_heads": [[250000, -230000], [270000, -225000]],
    "deleted_indices": [],
    "zones": [[zone_x1, zone_y1, zone_x2, zone_y2]],
}
r = requests.post(f"{BASE}/api/remote30/prototype/finalize/{job_id}", json=payload, timeout=15)
print(f"  → {r.json()}")

print(f"\n[4] Stage 3~5 stream")
t1 = time.time()
done_outputs = None
with requests.get(f"{BASE}/api/remote30/prototype/finalize_stream/{job_id}", stream=True, timeout=180) as resp:
    for raw in resp.iter_lines():
        if not raw or not raw.startswith(b"data: "):
            continue
        evt = json.loads(raw[6:])
        t = evt.get("type")
        if t == "stage":
            print(f"  [{(time.time()-t1):5.1f}s] stage {evt['stage']} {evt['status']:8s}: {evt['label']}")
        elif t == "entities":
            print(f"           ↳ stage={evt['stage']} summary={evt.get('summary')}")
        elif t == "tables_preview":
            print(f"           ↳ tables: {evt.get('counts')}")
        elif t == "done":
            done_outputs = evt.get("outputs")
            print(f"           ↳ DONE — outputs={done_outputs}")
            break
        elif t == "error":
            print(f"           ↳ ERROR {evt['message']}")
            break

print(f"\n[5] 결과: {done_outputs is not None and 'OK' or 'FAIL'}")
