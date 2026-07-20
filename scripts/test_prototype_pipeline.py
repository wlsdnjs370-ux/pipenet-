"""Remote 30 프로토타입 파이프라인 end-to-end 테스트.

대명동201동 단위세대_layer정리.dxf 업로드 → SSE stream 수신 → 결과 파일 검증.
"""

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

print(f"[1] POST run for {DXF.name}")
with DXF.open("rb") as f:
    r = requests.post(f"{BASE}/api/remote30/prototype/run",
                      files={"dxf_file": (DXF.name, f, "application/octet-stream")},
                      timeout=60)
data = r.json()
print(f"  → {data}")
if not data.get("ok"):
    sys.exit(1)
job_id = data["job_id"]

print(f"\n[2] SSE stream subscription for job {job_id}")
events_count = 0
stage_events = []
entity_counts = {}
table_counts = {}
done_outputs = None
t0 = time.time()
with requests.get(f"{BASE}/api/remote30/prototype/stream/{job_id}", stream=True, timeout=300) as resp:
    for raw in resp.iter_lines():
        if not raw:
            continue
        line = raw.decode("utf-8", errors="replace")
        if not line.startswith("data: "):
            continue
        evt = json.loads(line[6:])
        events_count += 1
        t = evt.get("type")
        if t == "stage":
            stage_events.append(evt)
            print(f"  [{(time.time()-t0):5.1f}s] stage {evt['stage']} {evt['status']:8s}: {evt['label']}")
        elif t == "entities":
            entity_counts[evt["stage"]] = len(evt["entities"])
            summ = evt.get("summary", {})
            print(f"           ↳ entities stage={evt['stage']} count={len(evt['entities'])} summary={summ}")
        elif t == "tables_preview":
            table_counts = evt["counts"]
            print(f"           ↳ tables {evt['counts']}")
        elif t == "done":
            done_outputs = evt["outputs"]
            print(f"           ↳ DONE outputs={evt['outputs']}")
            break
        elif t == "error":
            print(f"           ↳ ERROR {evt['message']}")
            break

print(f"\n[3] 총 {events_count} 이벤트 / 경과 {(time.time()-t0):.1f}s")

print("\n[4] 결과 파일 검증")
if not done_outputs:
    print("  → done 이벤트 미수신")
    sys.exit(1)

for label, fname in done_outputs.items():
    if not fname:
        continue
    if label.startswith("csv_"):
        url = f"{BASE}/api/remote30/prototype/result/{job_id}/csv/{fname}"
    else:
        url = f"{BASE}/api/remote30/prototype/result/{job_id}/{fname}"
    r = requests.get(url, timeout=30)
    if r.status_code != 200:
        print(f"  {label}: ERR HTTP {r.status_code}")
        continue
    size_kb = len(r.content) / 1024
    print(f"  {label:18s} {size_kb:8.1f} KB  {fname}")
    if label == "sdf":
        head = r.content[:500].decode("utf-8", errors="replace")
        print(f"      SDF preview:\n{head[:200]}...")
